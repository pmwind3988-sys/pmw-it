import { useState } from 'react';
import { Camera, AlertTriangle } from '../../../components/ui/Icons';
import { useSharePointImage } from './useSharePointImage';
import PhotoLightbox from './PhotoLightbox';

/**
 * A photograph that lives in SharePoint, shown wherever the register needs it.
 *
 * A photo that cannot be shown says so and offers the link, rather than
 * leaving the broken-image glyph that made every item in the register look
 * like nobody had photographed it.
 *
 * `caption` names what the picture is OF — "this one", "the whole delivery",
 * "the delivery order". Several photographs sit side by side on an item now,
 * and an unlabelled row of them is a puzzle rather than a record.
 *
 * `empty` is what to say when there is no photograph. It is worth setting per
 * place: "no photograph of this one" and "no delivery order was scanned" are
 * different facts and a reader should not have to work out which is meant.
 */
export default function AssetPhoto({
  siteUrl,
  stored,
  alt,
  caption,
  thumb = false,
  empty = 'No photograph was taken of this one.',
}) {
  const { url, href, failed, loading } = useSharePointImage(siteUrl, stored);
  const [open, setOpen] = useState(false);

  const wrap = (body) => (
    caption ? (
      <figure className={`as-shot${thumb ? ' as-shot-thumb' : ''}`}>
        {body}
        <figcaption className="as-shot-caption">{caption}</figcaption>
      </figure>
    ) : body
  );

  if (!href) {
    return wrap(
      <p className="as-photo-none">
        <Camera size={14} /> {empty}
      </p>,
    );
  }

  if (url) {
    return wrap(
      <>
        {/* A button rather than a bare image: this opens something, and
            anything that opens something has to be reachable from a keyboard
            and announce itself to a screen reader. */}
        <button
          type="button"
          className={`as-photo-open${thumb ? ' as-photo-open-thumb' : ''}`}
          onClick={() => setOpen(true)}
          title="See the whole photograph"
        >
          <img src={url} alt={alt || 'Item photograph'} className="as-detail-photo" />
        </button>
        {open && (
          <PhotoLightbox
            src={url}
            alt={caption || alt}
            href={href}
            onClose={() => setOpen(false)}
          />
        )}
      </>,
    );
  }

  if (failed) {
    return wrap(
      <p className="as-photo-none">
        <AlertTriangle size={14} />
        <span>
          The photograph could not be loaded.{' '}
          <a href={href} target="_blank" rel="noreferrer" className="as-link">
            Open it in SharePoint
          </a>
        </span>
      </p>,
    );
  }

  return wrap(loading ? <div className="spinner" /> : null);
}
