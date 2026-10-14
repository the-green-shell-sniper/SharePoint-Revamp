import { SPHttpClient } from '@microsoft/sp-http';

export interface ISubmissionUser {
  loginName: string;
  displayName: string;
}

export interface ISubmission {
  providedUrl: string;
  returnedUrl: string;
  outcome: 'Success' | 'Warning' | 'Error';
  message: string;
  submittedAt: string;
}

const LOG_SITE_URL = 'https://whallc.sharepoint.com/sites/SharePointLinkFixer';
const LOG_LIST_TITLE = 'URLFixerSubmissionLog';

export async function logSubmission(
  spHttpClient: SPHttpClient,
  user: ISubmissionUser,
  submission: ISubmission
): Promise<void> {
  if (!user.loginName) throw new Error('The signed-in SharePoint username is unavailable.');

  const response = await spHttpClient.post(
    `${LOG_SITE_URL}/_api/web/lists/getbytitle('${LOG_LIST_TITLE}')/items`,
    SPHttpClient.configurations.v1,
    {
      headers: {
        Accept: 'application/json;odata.metadata=none',
        'Content-Type': 'application/json;odata.metadata=none'
      },
      body: JSON.stringify({
        Title: 'URL submission',
        Username: user.loginName,
        Display_x0020_Name: user.displayName,
        ProvidedURL: submission.providedUrl,
        ReturnedURL: submission.returnedUrl,
        Outcome: submission.outcome,
        ResultMessage: submission.message,
        SubmittedAt: submission.submittedAt
      })
    }
  );
  const text = await response.text();
  if (!response.ok) {
    let detail: string | undefined;
    try {
      const body: { error?: { message?: string | { value?: string } } } = JSON.parse(text);
      const message = body?.error?.message;
      detail = typeof message === 'string' ? message : message?.value;
    } catch {
      throw new Error(`SharePoint submission log failed (HTTP ${response.status}). Check the log list and permissions.`);
    }
    throw new Error(`SharePoint submission log failed (HTTP ${response.status}). ${detail || 'Check the log list and permissions.'}`);
  }

  let item: { Id?: unknown };
  try {
    item = JSON.parse(text);
  } catch {
    throw new Error('SharePoint did not confirm the saved log entry. Check for a sign-in redirect.');
  }
  if (typeof item?.Id !== 'number' || !Number.isInteger(item.Id) || item.Id <= 0) {
    throw new Error('SharePoint did not return a valid log item ID.');
  }
}
