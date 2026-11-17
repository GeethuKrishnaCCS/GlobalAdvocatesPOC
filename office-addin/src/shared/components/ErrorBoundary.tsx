import * as React from "react";
import { Button, MessageBar, MessageBarActions, MessageBarBody, MessageBarTitle } from "@fluentui/react-components";

/* global console, window */

interface IErrorBoundaryProps {
  children: React.ReactNode;
}

interface IErrorBoundaryState {
  error: Error | undefined;
}

/** Catches render errors so a crash shows a message and Reload instead of a blank pane. Async errors stay in hooks. */
export class ErrorBoundary extends React.Component<IErrorBoundaryProps, IErrorBoundaryState> {
  public state: IErrorBoundaryState = { error: undefined };

  public static getDerivedStateFromError(error: Error): IErrorBoundaryState {
    return { error };
  }

  public componentDidCatch(error: Error, info: React.ErrorInfo): void {
    console.error("[pane] render error", error, info.componentStack);
  }

  public render(): React.ReactNode {
    if (!this.state.error) return this.props.children;
    return (
      <MessageBar intent="error" layout="multiline" data-testid="error-boundary-message">
        <MessageBarBody>
          <MessageBarTitle>Something went wrong.</MessageBarTitle>
          {this.state.error.message}
        </MessageBarBody>
        <MessageBarActions>
          <Button size="small" onClick={() => window.location.reload()} data-testid="error-boundary-reload">
            Reload
          </Button>
        </MessageBarActions>
      </MessageBar>
    );
  }
}
