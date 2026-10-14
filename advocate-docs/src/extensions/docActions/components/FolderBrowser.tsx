import * as React from 'react';
import {
  Breadcrumb, BreadcrumbButton, BreadcrumbDivider, BreadcrumbItem, Button, Caption1, Switch, makeStyles, tokens
} from '@fluentui/react-components';
import { Folder20Regular } from '@fluentui/react-icons';
import * as strings from 'DocActionsCommandSetStrings';
import type { IDocumentFinder } from '../hooks/useDocumentFinder';
import { format } from '../utils/format';

export interface IFolderBrowserProps {
  finder: IDocumentFinder;
  libraryTitle: string;
}

const useStyles = makeStyles({
  root: { display: 'flex', flexDirection: 'column', gap: tokens.spacingVerticalS, minHeight: 0 },
  list: { display: 'flex', flexDirection: 'column', overflowY: 'auto', maxHeight: '16rem' },
  folder: { justifyContent: 'flex-start' }
});

/** Browse library folders; the folder shown is the one "Find" reads. */
export const FolderBrowser: React.FC<IFolderBrowserProps> = ({ finder, libraryTitle }) => {
  const styles = useStyles();
  const { state, actions } = finder;
  const levels: string[] = [libraryTitle, ...state.folderPath];

  return (
    <div className={styles.root} data-testid="folder-browser">
      <Breadcrumb aria-label={strings.FolderPathLabel} size="small" data-testid="folder-browser-path">
        {levels.map((name, depth) => (
          <React.Fragment key={`${depth}-${name}`}>
            <BreadcrumbItem>
              <BreadcrumbButton current={depth === levels.length - 1} onClick={() => actions.goToFolderLevel(depth)}
                data-testid="folder-browser-level">{name}</BreadcrumbButton>
            </BreadcrumbItem>
            {depth < levels.length - 1 && <BreadcrumbDivider />}
          </React.Fragment>
        ))}
      </Breadcrumb>
      {state.subfolders.length === 0
        ? <Caption1 data-testid="folder-browser-empty">{strings.NoSubfolders}</Caption1>
        : (
          <div className={styles.list} role="list" data-testid="folder-browser-list">
            {state.subfolders.map(f => (
              <Button key={f.serverRelativeUrl} role="listitem" appearance="subtle" className={styles.folder}
                icon={<Folder20Regular />} onClick={() => actions.openFolder(f)} data-testid="folder-browser-folder">
                {format(strings.FolderWithCount, f.name, f.itemCount)}
              </Button>
            ))}
          </div>
        )}
      <Switch checked={state.includeSubfolders} label={strings.IncludeSubfolders}
        onChange={(_, d) => actions.setIncludeSubfolders(d.checked)} data-testid="folder-browser-recursive" />
    </div>
  );
};
