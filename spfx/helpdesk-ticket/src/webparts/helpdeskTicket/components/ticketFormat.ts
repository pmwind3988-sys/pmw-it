// Pure helpers for the ticket screens: file sizes, pasted screenshots and the
// "Updated" column's wording. Nothing here touches React or SharePoint.

export const MAX_FILE_BYTES = 10 * 1024 * 1024;

export function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${Math.round(bytes / 1024)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

/** Splits picked files into the ones under the limit and the ones over it. */
export function splitBySize(files: File[]): { ok: File[]; tooBig: File[] } {
  const ok: File[] = [];
  const tooBig: File[] = [];
  files.forEach((f) => (f.size > MAX_FILE_BYTES ? tooBig : ok).push(f));
  return { ok, tooBig };
}

export function tooBigMessage(names: string[]): string {
  if (!names.length) return '';
  const list = names.join(', ');
  return names.length === 1
    ? `${list} is bigger than 10 MB, so it was not added.`
    : `${list} are bigger than 10 MB, so they were not added.`;
}

function renamed(f: File, n: number): File {
  // "image/svg+xml" -> "svg", "image/jpeg" -> "jpg"
  const sub = (f.type.split('/')[1] ?? 'png').split('+')[0];
  const ext = sub === 'jpeg' ? 'jpg' : sub || 'png';
  return new File([f], `screenshot-${n}.${ext}`, { type: f.type, lastModified: Date.now() });
}

/**
 * Images on a paste, named screenshot-<n>.<ext>. Text and other files are
 * left to the browser. Some browsers expose the image only through `items`.
 */
export function imagesFromClipboard(data: DataTransfer | undefined, nextNumber: () => number): File[] {
  if (!data) return [];
  let images = Array.from(data.files ?? []).filter((f) => f.type.indexOf('image/') === 0);
  if (!images.length && data.items) {
    const fromItems: File[] = [];
    for (let i = 0; i < data.items.length; i++) {
      const item = data.items[i];
      if (item.kind === 'file' && item.type.indexOf('image/') === 0) {
        const f = item.getAsFile();
        if (f) fromItems.push(f);
      }
    }
    images = fromItems;
  }
  return images.map((f) => renamed(f, nextNumber()));
}

const pad = (n: number): string => (n < 10 ? '0' : '') + String(n);

/**
 * "Today, 10:40 AM", "Yesterday", "9 Oct", or "9 Oct 2025" from another year.
 * Empty when the value is not a date.
 */
export function friendlyUpdated(iso: string, now: Date): string {
  const d = new Date(iso);
  if (!iso || isNaN(d.getTime())) return '';
  const dayOf = (x: Date): number => new Date(x.getFullYear(), x.getMonth(), x.getDate()).getTime();
  const days = Math.round((dayOf(now) - dayOf(d)) / 86400000);
  if (days === 0) {
    const h = d.getHours();
    const h12 = h % 12 === 0 ? 12 : h % 12;
    return `Today, ${h12}:${pad(d.getMinutes())} ${h < 12 ? 'AM' : 'PM'}`;
  }
  if (days === 1) return 'Yesterday';
  const months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  const day = `${d.getDate()} ${months[d.getMonth()]}`;
  return d.getFullYear() === now.getFullYear() ? day : `${day} ${d.getFullYear()}`;
}
