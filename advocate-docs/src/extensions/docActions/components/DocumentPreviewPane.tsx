import * as React from 'react';
import { Body1, Skeleton, SkeletonItem, makeStyles, tokens } from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import { getEmbedPreviewUrl } from '../utils/previewUrl';
import { format } from '../utils/format';

export interface IDocumentPreviewPaneProps {
  webAbsoluteUrl: string;
  document: ISelectedDocument | undefined;
}

const useStyles = makeStyles({
  root: {
    position: 'relative',
    height: '100%',
    minHeight: 0,
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: tokens.borderRadiusMedium,
    overflow: 'hidden'
  },
  frame: { width: '100%', height: '100%', border: 'none', display: 'block' },
  loading: {
    position: 'absolute',
    inset: 0,
    display: 'flex',
    flexDirection: 'column',
    gap: tokens.spacingVerticalM,
    padding: tokens.spacingVerticalXXL
  },
  empty: { display: 'flex', alignItems: 'center', justifyContent: 'center', height: '100%' }
});

/** Full-document preview of one file; loads only the document passed in (never preloads others). */
export const DocumentPreviewPane: React.FC<IDocumentPreviewPaneProps> = ({ webAbsoluteUrl, document }) => {
  const styles = useStyles();
  const [loadedId, setLoadedId] = React.useState<string | undefined>(undefined);

  if (!document) {
    return (
      <div className={styles.root} data-testid="preview-pane-empty">
        <div className={styles.empty}><Body1>{strings.NothingToPreview}</Body1></div>
      </div>
    );
  }

  const isLoading: boolean = loadedId !== document.uniqueId;
  return (
    <div className={styles.root} data-testid="preview-pane">
      {isLoading && (
        <Skeleton className={styles.loading} aria-label={strings.PreviewLoading} data-testid="preview-pane-loading">
          <SkeletonItem size={24} />
          <SkeletonItem size={16} />
          <SkeletonItem size={16} />
        </Skeleton>
      )}
      <iframe
        key={document.uniqueId}
        className={styles.frame}
        src={getEmbedPreviewUrl(webAbsoluteUrl, document.uniqueId, document.fileType)}
        title={format(strings.PreviewFrameTitle, document.name)}
        onLoad={() => setLoadedId(document.uniqueId)}
        data-testid="preview-pane-frame"
      />
    </div>
  );
};
