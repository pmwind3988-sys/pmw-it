import { useState } from 'react';
import { Check, AlertTriangle } from '../../../components/ui/Icons';
import { useSharePointImage } from './useSharePointImage';
import PhotoLightbox from './PhotoLightbox';
import { SHAREPOINT_SITE_URL } from '../useAssets';

/**
 * The signature somebody actually drew, shown against the handover it belongs
 * to.
 *
 * It used to be the word "signed" linking off to SharePoint. That answered
 * "did anybody sign" and nothing else: to see WHOSE hand it was, or whether
 * the person who took the laptop is the person who brought it back, you had to
 * leave the app and open a file. And because every line of one handover
 * carries the same picture, a page of five items all said "signed" without
 * ever showing that they were signed by one person on one afternoon.
 *
 * So each handover shows its own signature, out and back, as the picture. The
 * word stays as the fallback for a signature that will not load — a picture
 * that is there and unreachable must not read as a handover nobody signed.
 *
 * `when` names the moment — "handed over", "brought back" — because a row can
 * carry two of these and a reader must not have to guess which is which.
 */
export default function SignatureShot({ stored, when, by }) {
  // The site is fixed for the whole register, so it is read here rather than
  // threaded through every list that shows a handover.
  const { url, href, failed } = useSharePointImage(SHAREPOINT_SITE_URL, stored);
  const [open, setOpen] = useState(false);

  // A handover recorded without a signature is a normal thing and says so by
  // saying nothing: a row reading "unsigned" against every line would drown
  // the ones that matter.
  if (!stored) return null;

  const label = by ? `${by} — ${when}` : when;

  if (url) {
    return (
      <span className="as-sigshot">
        <button
          type="button"
          className="as-sigshot-open"
          onClick={() => setOpen(true)}
          title={`See the signature — ${label}`}
        >
          <img src={url} alt={`Signature — ${label}`} className="as-sigshot-img" />
        </button>
        <span className="as-sigshot-when"><Check size={10} /> {when}</span>
        {open && (
          <PhotoLightbox
            src={url}
            alt={`Signature — ${label}`}
            href={href}
            onClose={() => setOpen(false)}
          />
        )}
      </span>
    );
  }

  if (failed) {
    return (
      <a href={href} target="_blank" rel="noreferrer" className="as-signed" title={label}>
        <AlertTriangle size={11} /> signed — open it in SharePoint
      </a>
    );
  }

  // Being fetched. The word rather than a spinner: a list of ten rows blinking
  // ten spinners reads as a page that is broken.
  return <span className="as-signed"><Check size={11} /> {when}</span>;
}
