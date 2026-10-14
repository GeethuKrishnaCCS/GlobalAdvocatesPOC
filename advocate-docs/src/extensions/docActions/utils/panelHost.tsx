import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { AppRoot } from '../components/AppRoot';

/**
 * Mounts one panel at a time in a body-level container (command sets have no DOM of their own).
 * Opening a new panel replaces the previous one.
 */
export class PanelHost {
  private _container: HTMLDivElement | undefined;

  public open(render: (close: () => void) => React.ReactElement): void {
    this.close();
    this._container = document.createElement('div');
    this._container.setAttribute('data-testid', 'doc-actions-panel-host');
    document.body.appendChild(this._container);
    ReactDOM.render(<AppRoot>{render(() => this.close())}</AppRoot>, this._container);
  }

  public close(): void {
    if (!this._container) {
      return;
    }
    ReactDOM.unmountComponentAtNode(this._container);
    this._container.remove();
    this._container = undefined;
  }
}
