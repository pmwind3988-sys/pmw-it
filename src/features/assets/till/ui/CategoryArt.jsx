/**
 * A small flat illustration per asset category, for the till's Hand out list:
 * a soft blob, the object in bold flat colour with a dark outline, and a
 * couple of confetti marks -- so "Keyboard" is recognised before it is read.
 *
 * The palette is fixed rather than themed: each picture sits on its own blob,
 * so it reads the same on the light and the dark till. A category somebody
 * added by hand is matched by keyword (`artKeyFor`), else gets the box.
 */
import { artKeyFor } from './categoryArtKey.js';

const INK = '#1f2340';
const C = {
  coral: '#ff7a6b', sun: '#ffc94a', teal: '#2ec4b6', lilac: '#9b8cff',
  sky: '#6cc3ff', pink: '#ff9fc8', paper: '#ffffff', mint: '#bdf0e0',
};
const line = { stroke: INK, strokeWidth: 2.2, strokeLinejoin: 'round', strokeLinecap: 'round' };

const BLOBS = [
  'M32 4c12 0 26 6 27 20s-6 33-24 35S4 50 5 33 20 4 32 4Z',
  'M30 5c15-2 29 9 29 25S47 60 31 59 4 47 5 30 15 7 30 5Z',
  'M34 5c13 2 25 12 24 27S45 60 29 58 5 44 6 29 21 3 34 5Z',
];

/** Confetti: a dot, a ring and a squiggle, placed per picture. */
function Confetti({ dot = [10, 14], ring = [54, 50], squiggle = [46, 9], color = C.sun }) {
  return (
    <g>
      <circle cx={dot[0]} cy={dot[1]} r="2.6" fill={color} />
      <circle cx={ring[0]} cy={ring[1]} r="2.8" fill="none" stroke={INK} strokeWidth="1.6" />
      <path d={`M${squiggle[0]} ${squiggle[1]}q2-3 4 0t4 0`} fill="none" stroke={INK} strokeWidth="1.6" strokeLinecap="round" />
    </g>
  );
}

const ART = {
  laptop: (
    <>
      <rect x="15" y="17" width="34" height="22" rx="3" fill={C.sky} {...line} />
      <rect x="19" y="21" width="26" height="14" rx="1.5" fill={C.paper} />
      <path d="M22 29l5-4 4 3 6-5" fill="none" stroke={C.coral} strokeWidth="2" strokeLinecap="round" />
      <path d="M9 41h46l-4 6H13Z" fill={C.lilac} {...line} />
    </>
  ),
  desktop: (
    <>
      <rect x="10" y="16" width="28" height="20" rx="3" fill={C.teal} {...line} />
      <rect x="14" y="20" width="20" height="12" rx="1.5" fill={C.paper} />
      <path d="M20 36v6h8v-6M16 44h16" fill="none" {...line} />
      <rect x="42" y="14" width="12" height="32" rx="3" fill={C.coral} {...line} />
      <circle cx="48" cy="20" r="2" fill={C.paper} />
      <path d="M45 27h6M45 31h6" {...line} />
    </>
  ),
  monitor: (
    <>
      <rect x="9" y="13" width="46" height="28" rx="3" fill={C.lilac} {...line} />
      <rect x="13" y="17" width="38" height="20" rx="1.5" fill={C.paper} />
      <circle cx="24" cy="27" r="5" fill={C.sun} />
      <path d="M33 33l6-8 6 8Z" fill={C.teal} />
      <path d="M32 41v6M22 49h20" fill="none" {...line} />
    </>
  ),
  printer: (
    <>
      <path d="M20 12h24v12H20Z" fill={C.paper} {...line} />
      <rect x="10" y="22" width="44" height="20" rx="4" fill={C.coral} {...line} />
      <circle cx="47" cy="29" r="2" fill={C.sun} />
      <path d="M20 36h24v14H20Z" fill={C.paper} {...line} />
      <path d="M24 41h16M24 45h10" stroke={C.lilac} strokeWidth="2" strokeLinecap="round" />
    </>
  ),
  dock: (
    <>
      <rect x="10" y="26" width="44" height="14" rx="5" fill={C.teal} {...line} />
      <rect x="16" y="31" width="5" height="4" rx="1" fill={INK} />
      <rect x="25" y="31" width="5" height="4" rx="1" fill={INK} />
      <circle cx="37" cy="33" r="2" fill={INK} />
      <path d="M44 33h5" {...line} />
      <path d="M32 26v-6q0-4 4-4h8" fill="none" {...line} />
      <rect x="44" y="12" width="7" height="8" rx="2" fill={C.sun} {...line} />
    </>
  ),
  phone: (
    <>
      <rect x="21" y="9" width="22" height="44" rx="5" fill={C.pink} {...line} />
      <rect x="25" y="15" width="14" height="28" rx="1.5" fill={C.paper} />
      <circle cx="29" cy="22" r="2.5" fill={C.teal} />
      <circle cx="35" cy="22" r="2.5" fill={C.sun} />
      <circle cx="29" cy="29" r="2.5" fill={C.coral} />
      <circle cx="35" cy="29" r="2.5" fill={C.lilac} />
      <path d="M29 48h6" {...line} />
    </>
  ),
  tablet: (
    <>
      <rect x="12" y="11" width="40" height="42" rx="5" fill={C.sun} {...line} />
      <rect x="17" y="16" width="30" height="30" rx="1.5" fill={C.paper} />
      <path d="M22 40c4-10 8-14 10-6s6 4 10-8" fill="none" stroke={C.coral} strokeWidth="2.4" strokeLinecap="round" />
      <circle cx="32" cy="49.5" r="1.6" fill={INK} />
    </>
  ),
  network: (
    <>
      <rect x="9" y="32" width="46" height="14" rx="4" fill={C.lilac} {...line} />
      <circle cx="17" cy="39" r="2" fill={C.sun} />
      <circle cx="24" cy="39" r="2" fill={C.teal} />
      <path d="M33 39h16" {...line} />
      <path d="M20 32l-3-10M44 32l3-10" {...line} />
      <path d="M26 20q6-6 12 0M22 15q10-10 20 0" fill="none" stroke={C.coral} strokeWidth="2.4" strokeLinecap="round" />
    </>
  ),
  part: (
    <>
      <rect x="10" y="22" width="44" height="18" rx="2" fill={C.teal} {...line} />
      <rect x="15" y="26" width="8" height="10" rx="1" fill={INK} />
      <rect x="28" y="26" width="8" height="10" rx="1" fill={INK} />
      <rect x="41" y="26" width="8" height="10" rx="1" fill={INK} />
      <path d="M13 40v5M18 40v5M23 40v5M28 40v5M36 40v5M41 40v5M46 40v5M51 40v5" stroke={C.sun} strokeWidth="2.4" strokeLinecap="round" />
    </>
  ),
  keyboard: (
    <>
      <rect x="7" y="20" width="50" height="24" rx="4" fill={C.sun} {...line} />
      {[13, 21, 29, 37, 45].map((x) => <rect key={`a${x}`} x={x} y="25" width="6" height="5" rx="1" fill={C.paper} stroke={INK} strokeWidth="1.4" />)}
      {[17, 25, 33, 41].map((x) => <rect key={`b${x}`} x={x} y="32" width="6" height="5" rx="1" fill={C.paper} stroke={INK} strokeWidth="1.4" />)}
      <rect x="49" y="32" width="4" height="5" rx="1" fill={C.coral} stroke={INK} strokeWidth="1.4" />
      <path d="M32 20q0-8 8-8" fill="none" {...line} />
    </>
  ),
  mouse: (
    <>
      <path d="M32 6q4-2 4 4" fill="none" {...line} />
      <rect x="20" y="14" width="24" height="38" rx="12" fill={C.coral} {...line} />
      <path d="M20 28h24M32 14v14" {...line} />
      <rect x="30" y="18" width="4" height="7" rx="2" fill={C.sun} stroke={INK} strokeWidth="1.6" />
    </>
  ),
  cable: (
    <>
      <path d="M14 48c0-14 10-10 18-18s16-4 18-14" fill="none" stroke={C.lilac} strokeWidth="5" strokeLinecap="round" />
      <path d="M14 48c0-14 10-10 18-18s16-4 18-14" fill="none" stroke={INK} strokeWidth="1.4" strokeLinecap="round" strokeDasharray="0 9" />
      <rect x="8" y="46" width="12" height="10" rx="2" fill={C.teal} {...line} />
      <path d="M11 56v3M17 56v3" {...line} />
      <rect x="44" y="6" width="12" height="10" rx="2" fill={C.sun} {...line} />
      <path d="M47 6V3M53 6V3" {...line} />
    </>
  ),
  adapter: (
    <>
      <rect x="16" y="18" width="24" height="26" rx="5" fill={C.teal} {...line} />
      <path d="M22 18v-7M34 18v-7" {...line} />
      <path d="M30 24l-5 8h6l-4 7" fill="none" stroke={C.sun} strokeWidth="2.6" strokeLinecap="round" strokeLinejoin="round" />
      <path d="M40 34h6q6 0 6 6v12" fill="none" {...line} />
    </>
  ),
  headset: (
    <>
      <path d="M14 36v-6a18 18 0 0 1 36 0v6" fill="none" stroke={INK} strokeWidth="3" strokeLinecap="round" />
      <rect x="9" y="32" width="11" height="16" rx="5" fill={C.pink} {...line} />
      <rect x="44" y="32" width="11" height="16" rx="5" fill={C.pink} {...line} />
      <path d="M15 48q0 8 12 8" fill="none" {...line} />
      <circle cx="29" cy="56" r="3" fill={C.sun} {...line} />
    </>
  ),
  accessory: (
    <>
      <path d="M12 26l20-10 20 10v20L32 56 12 46Z" fill={C.sun} {...line} />
      <path d="M12 26l20 10 20-10M32 36v20" fill="none" {...line} />
      <path d="M22 21l20 10v6" fill="none" stroke={C.coral} strokeWidth="3" strokeLinecap="round" />
    </>
  ),
  other: (
    <>
      <path d="M12 26l20-10 20 10v20L32 56 12 46Z" fill={C.mint} {...line} />
      <path d="M12 26l20 10 20-10M32 36v20" fill="none" {...line} />
      <path d="M28 44q0-4 4-4t4-4-4-4-4 3" fill="none" stroke={C.coral} strokeWidth="2.4" strokeLinecap="round" />
    </>
  ),
};

const BLOB_COLOR = {
  laptop: C.mint, desktop: '#ffe3d6', monitor: '#fff1c4', printer: '#e7e2ff',
  dock: '#ffe3ef', phone: '#d9f0ff', tablet: '#e7e2ff', network: '#d4f5ef',
  part: '#ffe3d6', keyboard: '#e7e2ff', mouse: '#d4f5ef', cable: '#fff1c4',
  adapter: '#ffe3ef', headset: '#d9f0ff', accessory: '#d4f5ef', other: '#fff1c4',
};

export default function CategoryArt({ category, size = 48 }) {
  const key = artKeyFor(category);
  const seed = key.length % BLOBS.length;
  return (
    <svg
      className="category-art"
      viewBox="0 0 64 64"
      width={size}
      height={size}
      role="img"
      aria-label={category || 'Other'}
    >
      <path d={BLOBS[seed]} fill={BLOB_COLOR[key]} />
      <Confetti
        dot={seed ? [54, 14] : [10, 14]}
        ring={seed ? [9, 52] : [55, 52]}
        squiggle={seed ? [6, 8] : [46, 8]}
        color={seed === 2 ? C.coral : C.sun}
      />
      {ART[key]}
    </svg>
  );
}
