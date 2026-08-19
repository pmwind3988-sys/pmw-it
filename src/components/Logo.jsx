import logo48 from '../assets/logo-48.png';
import logo88 from '../assets/logo-88.png';
import logo128 from '../assets/logo-128.png';

const SIZES = [
  { src: logo48, width: 48 },
  { src: logo88, width: 88 },
  { src: logo128, width: 128 },
];

/** Picks the smallest asset that still covers the rendered size. */
function pickSrc(size) {
  return (SIZES.find((entry) => entry.width >= size) || SIZES[SIZES.length - 1]).src;
}

export default function Logo({ size = 32, alt = 'PMW', className = '', style }) {
  return (
    <img
      src={pickSrc(size)}
      alt={alt}
      width={size}
      height={size}
      className={className}
      style={{ width: size, height: size, objectFit: 'contain', flexShrink: 0, display: 'block', ...style }}
    />
  );
}
