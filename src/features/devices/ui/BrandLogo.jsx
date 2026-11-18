import { useState } from 'react';
import { LOCAL_MARKS, initialsOfBrand, logoUrlOf } from '../brandLogo';

function LocalMark({ mark }) {
  if (mark.squares) {
    const [a, b, c, d] = mark.squares;
    return (
      <svg viewBox="0 0 24 24" width="70%" height="70%">
        <rect x="1" y="1" width="10.5" height="10.5" fill={a} />
        <rect x="12.5" y="1" width="10.5" height="10.5" fill={b} />
        <rect x="1" y="12.5" width="10.5" height="10.5" fill={c} />
        <rect x="12.5" y="12.5" width="10.5" height="10.5" fill={d} />
      </svg>
    );
  }
  return (
    <svg viewBox="0 0 24 24" width="76%" height="76%">
      <rect x="1" y="1" width="22" height="22" rx="5" fill={mark.bg} />
      <text x="12" y="17" textAnchor="middle" fontSize="14" fontWeight="700" fill="#fff"
        fontFamily="'Segoe UI', system-ui, sans-serif">{mark.letter}</text>
    </svg>
  );
}

/**
 * A small square badge: the brand's logo fetched online, or its initials when
 * there is no brand or the lookup has no logo for it. Decorative — the brand
 * is always written beside it in words.
 */
export default function BrandLogo({ brand, fallback, size = 28 }) {
  const url = logoUrlOf(brand?.slug);
  const [failed, setFailed] = useState(null);
  const showImage = url && failed !== url;
  return (
    <span className="brand-logo" style={{ width: size, height: size }} aria-hidden="true"
      title={brand && !brand.guessed ? brand.name : undefined}>
      {LOCAL_MARKS[brand?.slug] ? <LocalMark mark={LOCAL_MARKS[brand.slug]} /> : showImage
        ? <img src={url} alt="" loading="lazy" referrerPolicy="no-referrer" onError={() => setFailed(url)} />
        : <span className="brand-logo-initials">{initialsOfBrand(brand?.name || fallback)}</span>}
    </span>
  );
}
