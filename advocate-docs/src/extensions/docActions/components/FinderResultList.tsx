import * as React from 'react';
import { Badge, Body1, Caption1, Checkbox, makeStyles, mergeClasses, tokens } from '@fluentui/react-components';
import * as strings from 'DocActionsCommandSetStrings';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { IResultSelection } from '../hooks/useResultSelection';
import { format } from '../utils/format';

export interface IFinderResultListProps {
  results: ReadonlyArray<ISelectedDocument>;
  existingIds: ReadonlySet<string>;
  selection: IResultSelection;
  previewId: string | undefined;
  onPreview: (doc: ISelectedDocument) => void;
}

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalXS, minHeight: 0 },
  list: { display: 'flex', flexDirection: 'column', overflowY: 'auto', maxHeight: '40vh', margin: 0, padding: 0, listStyleType: 'none' },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    paddingRight: tokens.spacingHorizontalS,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover }
  },
  previewing: { backgroundColor: tokens.colorNeutralBackground1Selected },
  text: { display: 'flex', flexDirection: 'column', flexGrow: 1, minWidth: 0 },
  name: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }
});

/** Finder results, one row per document: tick to pick, click the row to preview it. */
export const FinderResultList: React.FC<IFinderResultListProps> = ({ results, existingIds, selection, previewId, onPreview }) => {
  const styles = useStyles();
  const { state, actions, computed } = selection;

  const onRowKeyDown = (e: React.KeyboardEvent, doc: ISelectedDocument): void => {
    if (e.key === 'Enter') { e.preventDefault(); onPreview(doc); }
  };

  return (
    <div className={styles.root} data-testid="finder-results">
      <Checkbox
        checked={computed.allSelected ? true : computed.someSelected ? 'mixed' : false}
        disabled={computed.selectable.length === 0}
        onChange={(_, d) => actions.selectAll(d.checked === true)}
        label={format(strings.SelectAllResults, computed.selectable.length)}
        data-testid="finder-results-select-all"
      />
      <ul className={styles.list} aria-label={strings.FinderResultsLabel} data-testid="finder-results-list">
        {results.map(doc => {
          const inBundle: boolean = existingIds.has(doc.uniqueId);
          return (
            <li
              key={doc.uniqueId}
              className={mergeClasses(styles.row, doc.uniqueId === previewId && styles.previewing)}
              tabIndex={0}
              aria-current={doc.uniqueId === previewId}
              onClick={() => onPreview(doc)}
              onKeyDown={e => onRowKeyDown(e, doc)}
              data-testid="finder-results-row"
            >
              <Checkbox
                checked={inBundle || state.selectedIds.has(doc.uniqueId)}
                disabled={inBundle}
                aria-label={format(strings.PickDocument, doc.name)}
                onClick={e => e.stopPropagation()}
                onChange={(_, d) => actions.toggle(doc.uniqueId, d.checked === true)}
                data-testid="finder-results-pick"
              />
              <span className={styles.text}>
                <Body1 className={styles.name} title={doc.name} data-testid="finder-results-name">{doc.name}</Body1>
                <Caption1>{doc.fileType}</Caption1>
              </span>
              {inBundle && <Badge appearance="outline" data-testid="finder-results-in-bundle">{strings.InBundle}</Badge>}
            </li>
          );
        })}
      </ul>
    </div>
  );
};
