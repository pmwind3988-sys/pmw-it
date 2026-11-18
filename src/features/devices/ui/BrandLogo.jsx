import { useState } from 'react';
import { initialsOfBrand, logoUrlOf } from '../brandLogo';

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
      {showImage
        ? <img src={url} alt="" loading="lazy" referrerPolicy="no-referrer" onError={() => setFailed(url)} />
        : <span className="brand-logo-initials">{initialsOfBrand(brand?.name || fallback)}</span>}
    </span>
  );
}
