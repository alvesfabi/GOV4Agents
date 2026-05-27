import { MessageBar, MessageBarBody, MessageBarTitle } from '@fluentui/react-components';
import { JsonViewer } from './JsonViewer';
import { ApiError } from '../api/client';

interface ErrorPanelProps {
  error: unknown;
  title?: string;
}

export function ErrorPanel({ error, title = 'Request failed' }: ErrorPanelProps) {
  const message = error instanceof Error ? error.message : String(error);
  const body = error instanceof ApiError ? error.body : undefined;
  return (
    <MessageBar intent="error">
      <MessageBarBody>
        <MessageBarTitle>{title}</MessageBarTitle>
        <div>{message}</div>
        {body !== undefined && <JsonViewer value={body} label="Response body" />}
      </MessageBarBody>
    </MessageBar>
  );
}
