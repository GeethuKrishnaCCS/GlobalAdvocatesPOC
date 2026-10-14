import * as React from 'react';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';

export interface IAppRootProps {
  children: React.ReactNode;
}

/**
 * One FluentProvider per mounted panel. Panels render in a body-level layer above the list view,
 * on SharePoint's light dialog surface, so the light theme is used rather than the section theme.
 */
export const AppRoot: React.FC<IAppRootProps> = ({ children }) => (
  <FluentProvider theme={webLightTheme} style={{ background: 'transparent' }} data-testid="doc-actions-root">
    {children}
  </FluentProvider>
);
