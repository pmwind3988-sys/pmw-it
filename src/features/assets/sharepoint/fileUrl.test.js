import { describe, it, expect } from 'vitest';
import { absoluteFileUrl, fileApiPath } from './fileUrl.js';

const SITE = 'https://pmwgroupcom.sharepoint.com/sites/IThelpdesk';

describe('absoluteFileUrl', () => {
  /**
   * The bug this exists for: a server-relative path in an `<img src>` asks the
   * PORTAL for a file only SharePoint has.
   */
  it('puts the SharePoint host in front of a server-relative path', () => {
    expect(absoluteFileUrl(SITE, '/sites/IThelpdesk/IT%20Asset%20Photos/a.jpg'))
      .toBe('https://pmwgroupcom.sharepoint.com/sites/IThelpdesk/IT%20Asset%20Photos/a.jpg');
  });

  it('encodes the spaces a library title puts in the path', () => {
    expect(absoluteFileUrl(SITE, '/sites/IThelpdesk/IT Asset Photos/a b.jpg'))
      .toBe('https://pmwgroupcom.sharepoint.com/sites/IThelpdesk/IT%20Asset%20Photos/a%20b.jpg');
  });

  it('does not double-encode a path that arrived encoded', () => {
    expect(absoluteFileUrl(SITE, '/sites/x/IT%20Asset%20Photos/a.jpg'))
      .toBe('https://pmwgroupcom.sharepoint.com/sites/x/IT%20Asset%20Photos/a.jpg');
  });

  it('leaves an absolute URL exactly as it found it', () => {
    const url = 'https://example.com/photo.jpg';
    expect(absoluteFileUrl(SITE, url)).toBe(url);
  });

  it('hangs a bare name off the site', () => {
    expect(absoluteFileUrl(SITE, 'photo.jpg')).toBe(`${SITE}/photo.jpg`);
  });

  it('answers with nothing when there is nothing to show', () => {
    expect(absoluteFileUrl(SITE, '')).toBe('');
    expect(absoluteFileUrl(SITE, null)).toBe('');
    expect(absoluteFileUrl(SITE, '   ')).toBe('');
  });

  /** A half-formed URL is worse than none: it renders as a broken image. */
  it('answers with nothing when the site itself is unusable', () => {
    expect(absoluteFileUrl('', '/sites/x/a.jpg')).toBe('');
  });
});

describe('fileApiPath', () => {
  /**
   * The library path answers a cross-origin fetch with no CORS headers at all,
   * so the picture can only be read through the API.
   */
  it('asks the API for the bytes, not the library', () => {
    expect(fileApiPath('/sites/IThelpdesk/IT Asset Photos/a.jpg')).toBe(
      "/_api/web/GetFileByServerRelativeUrl('/sites/IThelpdesk/IT%20Asset%20Photos/a.jpg')/$value",
    );
  });

  it('takes the path out of an absolute URL', () => {
    expect(fileApiPath('https://pmwgroupcom.sharepoint.com/sites/x/lib/a.jpg')).toContain(
      "'/sites/x/lib/a.jpg'",
    );
  });

  /** An apostrophe in a file name would otherwise close the OData literal. */
  it('encodes an apostrophe rather than letting it end the argument', () => {
    expect(fileApiPath("/sites/x/lib/ali's tab.jpg")).toContain('ali%27s%20tab.jpg');
  });

  it('answers with nothing it cannot build a path from', () => {
    expect(fileApiPath('')).toBe('');
    expect(fileApiPath('photo.jpg')).toBe('');
  });
});
