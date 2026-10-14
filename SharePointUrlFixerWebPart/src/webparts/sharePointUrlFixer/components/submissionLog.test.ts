jest.mock('@microsoft/sp-http', () => ({
  SPHttpClient: Object.assign(jest.fn(() => ({ post: jest.fn() })), { configurations: { v1: {} } })
}));
jest.mock('@microsoft/sp-core-library', () => ({ ServiceScope: { startNewRoot: jest.fn() } }));

import { ServiceScope } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { ISubmission, logSubmission } from './submissionLog';

describe('SharePoint submission log', () => {
  const user = { loginName: 'i:0#.f|membership|user@wha.example', displayName: 'Test User' };
  const submission: ISubmission = {
    providedUrl: 'https://whallc.sharepoint.com/CS/Guide.docx?x="a,b"\n',
    returnedUrl: 'https://whallc.sharepoint.com/sites/ClinicalandPatientServices/Guide.docx',
    outcome: 'Success',
    bookmarkButtonName: 'Guide',
    message: 'Updated link ready.',
    submittedAt: '2026-10-14T02:30:00.000Z'
  };
  let httpClient: SPHttpClient;
  let post: jest.SpyInstance<Promise<SPHttpClientResponse>>;

  function response(text: string, status = 201): SPHttpClientResponse {
    return { ok: status >= 200 && status < 300, status, text: async () => text } as SPHttpClientResponse;
  }

  beforeEach(() => {
    httpClient = new SPHttpClient(ServiceScope.startNewRoot());
    post = jest.spyOn(httpClient, 'post').mockResolvedValue(response('{"Id":42}'));
  });

  it('adds one independent row to the central site with complete URL strings', async () => {
    await logSubmission(httpClient, user, submission);
    expect(post).toHaveBeenCalledTimes(1);
    const [url, configuration, options] = post.mock.calls[0];
    expect(url).toBe("https://whallc.sharepoint.com/sites/SharePointLinkFixer/_api/web/lists/getbytitle('URLFixerSubmissionLog')/items");
    expect(configuration).toBe(SPHttpClient.configurations.v1);
    expect(options.headers).toEqual({
      Accept: 'application/json;odata.metadata=none',
      'Content-Type': 'application/json;odata.metadata=none'
    });
    expect(JSON.parse(options.body)).toEqual({
      Title: 'URL submission',
      Username: user.loginName,
      Display_x0020_Name: user.displayName,
      ProvidedURL: submission.providedUrl,
      ReturnedURL: submission.returnedUrl,
      BookmarkButtonName: submission.bookmarkButtonName,
      Outcome: 'Success',
      ResultMessage: submission.message,
      SubmittedAt: submission.submittedAt
    });
  });

  it('uses separate creates for simultaneous submissions', async () => {
    await Promise.all([
      logSubmission(httpClient, user, submission),
      logSubmission(httpClient, user, { ...submission, outcome: 'Warning' })
    ]);
    expect(post).toHaveBeenCalledTimes(2);
  });

  it.each([403, 404, 429, 500])('reports HTTP %s with SharePoint error details and does not retry', async (status) => {
    post.mockResolvedValueOnce(response('{"error":{"message":"Cannot save item."}}', status));
    await expect(logSubmission(httpClient, user, submission)).rejects.toThrow(
      `SharePoint submission log failed (HTTP ${status}). Cannot save item.`
    );
    expect(post).toHaveBeenCalledTimes(1);
  });

  it('reads verbose SharePoint error messages', async () => {
    post.mockResolvedValueOnce(response('{"error":{"message":{"value":"Access denied."}}}', 403));
    await expect(logSubmission(httpClient, user, submission)).rejects.toThrow('Access denied.');
  });

  it('reports non-JSON failure responses without rendering HTML', async () => {
    post.mockResolvedValueOnce(response('<html>Gateway error</html>', 502));
    await expect(logSubmission(httpClient, user, submission)).rejects.toThrow('HTTP 502');
  });

  it.each(['<html>Sign in</html>', '{}', '{"Id":"42"}', '{"Id":0}', 'null'])(
    'does not treat an unconfirmed response as a saved row: %s', async (body) => {
      post.mockResolvedValueOnce(response(body, 200));
      await expect(logSubmission(httpClient, user, submission)).rejects.toThrow('SharePoint did not');
    }
  );

  it('propagates network failures without attempting duplicate writes', async () => {
    post.mockRejectedValueOnce(new Error('Network unavailable.'));
    await expect(logSubmission(httpClient, user, submission)).rejects.toThrow('Network unavailable.');
    expect(post).toHaveBeenCalledTimes(1);
  });

  it('does not silently log an unknown username', async () => {
    await expect(logSubmission(httpClient, { ...user, loginName: '' }, submission)).rejects.toThrow(
      'The signed-in SharePoint username is unavailable.'
    );
    expect(post).not.toHaveBeenCalled();
  });
});
