import * as React from 'react';
import * as strings from 'DocActionsCommandSetStrings';
import type { GraphFileService } from '../services/GraphFileService';
import { MAX_DRAFT_ATTACHMENT_BYTES, type IDraftLink, type MailDraftService } from '../services/MailDraftService';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { IDocumentVersion, IDraftResult, IDriveItemRef, IFilePayload, ISendRequest, ISendStep } from '../models/ISend';
import { format } from '../utils/format';

/** Subject lists up to this many file names, then "and n more". */
const SUBJECT_NAME_LIMIT: number = 3;

export interface ISendDraft {
  state: {
    versions: IDocumentVersion[] | undefined;
    selectedVersionId: string | undefined;
    steps: ISendStep[];
    isRunning: boolean;
    result: IDraftResult | undefined;
    error: string | undefined;
  };
  actions: { selectVersion: (id: string) => void; run: () => Promise<void> };
  computed: { canRun: boolean };
}

function subjectFor(docs: ReadonlyArray<ISelectedDocument>): string {
  const names: string[] = docs.slice(0, SUBJECT_NAME_LIMIT).map(d => d.name);
  const rest: number = docs.length - names.length;
  return rest > 0 ? format(strings.SendSubjectMore, names.join(', '), rest) : names.join(', ');
}

/** Builds one Outlook draft with the requested documents attached (or linked when too large). */
export function useSendDraft(
  files: GraphFileService,
  mail: MailDraftService,
  documents: ReadonlyArray<ISelectedDocument>,
  request: ISendRequest,
  webOrigin: string
): ISendDraft {
  const [versions, setVersions] = React.useState<IDocumentVersion[] | undefined>(undefined);
  const [selectedVersionId, setSelectedVersionId] = React.useState<string | undefined>(undefined);
  const [refs, setRefs] = React.useState<Map<string, IDriveItemRef>>(new Map());
  const [steps, setSteps] = React.useState<ISendStep[]>(documents.map(d => ({ name: d.name, status: 'waiting' })));
  const [isRunning, setIsRunning] = React.useState<boolean>(false);
  const [result, setResult] = React.useState<IDraftResult | undefined>(undefined);
  const [error, setError] = React.useState<string | undefined>(undefined);

  // Previous-version mode: load the single document's history and preselect the newest previous version.
  React.useEffect(() => {
    if (request.mode !== 'version' || documents.length !== 1) {
      return;
    }
    let cancelled: boolean = false;
    files.resolve(documents[0])
      .then(async ref => {
        const list: IDocumentVersion[] = await files.listVersions(ref);
        if (!cancelled) {
          setRefs(new Map([[documents[0].uniqueId, ref]]));
          setVersions(list);
          setSelectedVersionId((list.find(v => !v.isCurrent) ?? list[0])?.id);
        }
      })
      .catch((e: Error) => { if (!cancelled) { setError(e.message); setVersions([]); } });
    return () => { cancelled = true; };
  }, [files, documents, request.mode]);

  const updateStep = (index: number, step: Partial<ISendStep>): void =>
    setSteps(current => current.map((s, i) => (i === index ? { ...s, ...step } : s)));

  const run = React.useCallback(async (): Promise<void> => {
    setIsRunning(true);
    setError(undefined);
    try {
      const draft: { id: string; webLink: string } = await mail.createDraft(subjectFor(documents));
      const links: IDraftLink[] = [];
      let attachedBytes: number = 0;
      let attached: number = 0;
      let failed: number = 0;
      for (let i = 0; i < documents.length; i++) {
        const doc: ISelectedDocument = documents[i];
        updateStep(i, { status: 'working' });
        try {
          const ref: IDriveItemRef = refs.get(doc.uniqueId) ?? await files.resolve(doc);
          const version: IDocumentVersion | undefined = versions?.find(v => v.id === selectedVersionId);
          const payload: IFilePayload = request.mode === 'version' && version
            ? await files.getVersion(ref, version, request.asPdf)
            : await files.getCurrent(ref, request.asPdf);
          if (attachedBytes + payload.bytes.byteLength > MAX_DRAFT_ATTACHMENT_BYTES) {
            links.push({ name: payload.name, url: `${webOrigin}${doc.serverRelativeUrl}` });
            updateStep(i, { status: 'linked', name: payload.name });
          } else {
            await mail.attach(draft.id, payload);
            attachedBytes += payload.bytes.byteLength;
            attached++;
            updateStep(i, { status: 'attached', name: payload.name });
          }
        } catch (e) {
          failed++;
          updateStep(i, { status: 'failed', detail: (e as Error).message });
        }
      }
      await mail.setBody(draft.id, strings.SendBodyIntro, strings.SendBodyLinks, links);
      setResult({ webLink: draft.webLink, attached, linked: links.length, failed });
    } catch (e) {
      setError((e as Error).message);
    } finally {
      setIsRunning(false);
    }
  }, [mail, files, documents, refs, versions, selectedVersionId, request, webOrigin]);

  const canRun: boolean = !isRunning && !result && (request.mode === 'current' || selectedVersionId !== undefined);

  return {
    state: { versions, selectedVersionId, steps, isRunning, result, error },
    actions: { selectVersion: setSelectedVersionId, run },
    computed: { canRun }
  };
}
