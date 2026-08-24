import { Camera, AlertTriangle } from '../../../components/ui/Icons';
import { useSharePointImage } from './useSharePointImage';

/**
 * The photograph on a saved item.
 *
 * A photo that cannot be shown says so and offers the link, rather than
 * leaving the broken-image glyph that made every item in the register look
 * like nobody had photographed it.
 */
export default function AssetPhoto({ siteUrl, stored, alt }) {
  const { url, href, failed, loading } = useSharePointImage(siteUrl, stored);

  if (!href) {
    return (
      <p className="as-photo-none">
        <Camera size={14} /> No photograph was taken of this one.
      </p>
    );
  }

  if (url) return <img src={url} alt={alt || 'Item photograph'} className="as-detail-photo" />;

  if (failed) {
    return (
      <p className="as-photo-none">
        <AlertTriangle size={14} />
        <span>
          The photograph could not be loaded.{' '}
          <a href={href} target="_blank" rel="noreferrer" className="as-link">
            Open it in SharePoint
          </a>
        </span>
      </p>
    );
  }

  return loading ? <div className="spinner" /> : null;
}
