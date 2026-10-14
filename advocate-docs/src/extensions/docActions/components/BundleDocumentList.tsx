import * as React from 'react';
import {
  Body1, Caption1, Menu, MenuItem, MenuList, MenuPopover, MenuTrigger, Button, Tooltip, makeStyles, mergeClasses, tokens
} from '@fluentui/react-components';
import { ArrowDown20Regular, ArrowUp20Regular, Delete20Regular, MoreHorizontal20Regular } from '@fluentui/react-icons';
import * as strings from 'DocActionsCommandSetStrings';
import type { ISelectedDocument } from '../models/ISelectedDocument';
import type { IBundleDraft } from '../hooks/useBundleDraft';
import { format } from '../utils/format';

export interface IBundleDocumentListProps {
  draft: IBundleDraft;
}

const useStyles = makeStyles({
  list: { display: 'flex', flexDirection: 'column', overflowY: 'auto', minHeight: 0, margin: 0, padding: 0, listStyleType: 'none' },
  row: {
    display: 'flex',
    alignItems: 'center',
    gap: tokens.spacingHorizontalS,
    padding: `${tokens.spacingVerticalS} ${tokens.spacingHorizontalS}`,
    borderRadius: tokens.borderRadiusMedium,
    cursor: 'pointer',
    ':hover': { backgroundColor: tokens.colorNeutralBackground1Hover }
  },
  selected: { backgroundColor: tokens.colorNeutralBackground1Selected },
  order: { minWidth: tokens.spacingHorizontalXXL, color: tokens.colorNeutralForeground3 },
  text: { display: 'flex', flexDirection: 'column', flexGrow: 1, minWidth: 0 },
  name: { overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }
});

/** Ordered bundle members; clicking a row shows it in the preview pane. */
export const BundleDocumentList: React.FC<IBundleDocumentListProps> = ({ draft }) => {
  const styles = useStyles();
  const { state, actions } = draft;

  const onKeyDown = (e: React.KeyboardEvent, doc: ISelectedDocument): void => {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); actions.highlight(doc.id); }
  };

  return (
    <ol className={styles.list} aria-label={strings.BundleBuilderTitle} data-testid="bundle-list">
      {state.documents.map((doc, i) => (
        <li
          key={doc.id}
          className={mergeClasses(styles.row, doc.id === state.highlightedId && styles.selected)}
          tabIndex={0}
          aria-current={doc.id === state.highlightedId}
          onClick={() => actions.highlight(doc.id)}
          onKeyDown={e => onKeyDown(e, doc)}
          data-testid="bundle-list-row"
        >
          <Caption1 className={styles.order}>{i + 1}</Caption1>
          <span className={styles.text}>
            <Tooltip content={doc.name} relationship="description">
              <Body1 className={styles.name} data-testid="bundle-list-name">{doc.name}</Body1>
            </Tooltip>
            <Caption1>{doc.fileType}</Caption1>
          </span>
          <Menu>
            <MenuTrigger disableButtonEnhancement>
              <Button appearance="subtle" icon={<MoreHorizontal20Regular />} aria-label={format(strings.DocumentActions, doc.name)}
                onClick={e => e.stopPropagation()} data-testid="bundle-list-menu" />
            </MenuTrigger>
            <MenuPopover>
              <MenuList>
                <MenuItem icon={<ArrowUp20Regular />} disabled={i === 0} onClick={() => actions.moveUp(doc.id)}
                  data-testid="bundle-list-move-up">{strings.MoveUp}</MenuItem>
                <MenuItem icon={<ArrowDown20Regular />} disabled={i === state.documents.length - 1} onClick={() => actions.moveDown(doc.id)}
                  data-testid="bundle-list-move-down">{strings.MoveDown}</MenuItem>
                <MenuItem icon={<Delete20Regular />} onClick={() => actions.remove(doc.id)}
                  data-testid="bundle-list-remove">{strings.RemoveFromBundle}</MenuItem>
              </MenuList>
            </MenuPopover>
          </Menu>
        </li>
      ))}
    </ol>
  );
};
