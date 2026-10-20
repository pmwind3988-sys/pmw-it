import { familyOf } from '../categoryFamilies';

/**
 * A category's mark: a small tinted tile with a drawing of the thing.
 *
 * Each drawing has one loop of motion in it (a laptop's screen types, a
 * printer feeds a page, a router's signal pulses), and `assets.css` keeps it
 * STILL until the row it sits on is pointed at — a register of two hundred
 * rows all moving at once would be noise. A bulk category gets a stacked edge
 * behind the tile, which is what the old box icon used to say.
 *
 * Drawn on a 32px grid in the same stroke style as `components/ui/Icons.jsx`.
 * A category somebody added that has no drawing of its own gets the open box.
 */

const GLYPHS = {
  Laptop: (
    <>
      <rect className="solid" x="6" y="7" width="20" height="13.5" rx="2" />
      <path className="a-type" d="M10 11.5h8" />
      <path className="a-type d2" d="M10 14.5h12" />
      <path className="a-type d3" d="M10 17.5h5" />
      <path className="solid" d="M3 23.5h26l-1.6 2.5H4.6z" />
    </>
  ),
  Desktop: (
    <>
      <rect className="solid" x="9" y="3" width="14" height="26" rx="2.5" />
      <path d="M13 8h6" />
      <circle cx="16" cy="19" r="4.5" />
      <g className="a-spin"><path d="M16 19v-3.2M16 19l2.8 1.6M16 19l-2.8 1.6" /></g>
      <circle className="dot a-blink" cx="16" cy="11.5" r="1" />
    </>
  ),
  Monitor: (
    <>
      <rect className="solid" x="3" y="5" width="26" height="17" rx="2" />
      <path className="a-scan" d="M6 9h20" />
      <path d="M16 22v4M11 27h10" />
      <circle className="dot" cx="25.5" cy="19" r=".9" />
    </>
  ),
  Printer: (
    <>
      <g className="a-feed">
        <rect className="paper" x="9" y="18" width="14" height="10" rx="1" />
        <path d="M12 22.5h8M12 25.5h5" />
      </g>
      <path className="paper" d="M9 11V4.5h14V11" />
      <rect className="solid" x="4" y="11" width="24" height="11" rx="2" />
      <path d="M8 18.5h16" />
      <circle className="dot a-blink" cx="23.5" cy="14.5" r="1.1" />
    </>
  ),
  'Docking Station': (
    <>
      <g className="a-plug">
        <path d="M16 3v4" />
        <rect className="solid" x="13" y="7" width="6" height="5" rx="1" />
      </g>
      <rect className="solid" x="3" y="13" width="26" height="10" rx="2.5" />
      <path d="M7.5 18h3M13.5 18h3" />
      <circle className="dot a-led" cx="23.5" cy="18" r="1.3" />
      <path d="M7 23v2.5M25 23v2.5" />
    </>
  ),
  Phone: (
    <>
      <path className="a-wave" d="M4.5 12v8M27.5 12v8" />
      <g className="a-buzz">
        <rect className="solid" x="9" y="3" width="14" height="26" rx="3" />
        <path d="M14 25.5h4" />
      </g>
      <circle className="hotdot a-ping" cx="22.5" cy="5" r="3" />
    </>
  ),
  Tab: (
    <>
      <rect className="solid" x="3" y="7" width="26" height="18" rx="3" />
      <rect className="soft a-swipe" x="7" y="11" width="9" height="10" rx="1.5" />
      <path d="M20 12.5h5M20 16h3" />
    </>
  ),
  Network: (
    <>
      <path className="a-wifi w3" d="M8.5 9.5a10.5 10.5 0 0 1 15 0" />
      <path className="a-wifi w2" d="M11 12.5a7 7 0 0 1 10 0" />
      <path className="a-wifi" d="M13.5 15.5a3.5 3.5 0 0 1 5 0" />
      <rect className="solid" x="4" y="19" width="24" height="8" rx="2" />
      <circle className="dot a-blink" cx="9" cy="23" r="1" />
      <circle className="dot" cx="13" cy="23" r="1" />
      <circle className="dot" cx="17" cy="23" r="1" />
    </>
  ),
  'PC Part': (
    <>
      <path d="M12 4v4M16 4v4M20 4v4M12 24v4M16 24v4M20 24v4M4 12h4M4 16h4M4 20h4M24 12h4M24 16h4M24 20h4" />
      <rect className="solid" x="8" y="8" width="16" height="16" rx="2.5" />
      <rect className="a-core" x="12.5" y="12.5" width="7" height="7" rx="1" />
    </>
  ),
  Keyboard: (
    <>
      <rect className="solid" x="2.5" y="9" width="27" height="15" rx="2.5" />
      <rect className="soft a-press" x="6" y="12.5" width="2.6" height="2.4" rx=".6" />
      <rect className="soft a-press p3" x="10.2" y="12.5" width="2.6" height="2.4" rx=".6" />
      <rect className="soft a-press p5" x="14.4" y="12.5" width="2.6" height="2.4" rx=".6" />
      <rect className="soft a-press p2" x="18.6" y="12.5" width="2.6" height="2.4" rx=".6" />
      <rect className="soft a-press p4" x="22.8" y="12.5" width="2.6" height="2.4" rx=".6" />
      <rect className="soft a-press p6" x="10" y="18.5" width="12" height="2.2" rx=".6" />
    </>
  ),
  Mouse: (
    <>
      <rect className="solid" x="9" y="3.5" width="14" height="25" rx="7" />
      <path d="M16 3.5v4M9.3 14h13.4" />
      <rect className="dot a-scroll" x="15" y="8" width="2" height="4" rx="1" />
    </>
  ),
  Cable: (
    <>
      <path d="M5 25C5 17 11 17 16 16S27 15 27 7" />
      <path className="a-pulse hot" d="M5 25C5 17 11 17 16 16S27 15 27 7" />
      <rect className="solid" x="2.5" y="25" width="5" height="5" rx="1" />
      <rect className="solid" x="24.5" y="2" width="5" height="5" rx="1" />
    </>
  ),
  Adapter: (
    <>
      <path d="M12.5 3.5v5M19.5 3.5v5M16 27v3" />
      <rect className="solid" x="7" y="8.5" width="18" height="18.5" rx="3" />
      <path className="hot a-bolt" d="M17.5 12l-4 6h4.5l-2.5 5" />
    </>
  ),
  Accessory: (
    <g className="a-swing">
      <path d="M12 10V7.5a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2V10" />
      <rect className="solid" x="5" y="10" width="22" height="16" rx="3" />
      <path d="M5 16.5h22M14.5 16.5V19h3v-2.5" />
    </g>
  ),
  Other: (
    <>
      <path className="solid" d="M5.5 13.5h21V26a2 2 0 0 1-2 2h-17a2 2 0 0 1-2-2z" />
      <path d="M16 13.5v5" />
      <g className="a-lid"><rect className="solid" x="4" y="9.5" width="24" height="4" rx="1" /></g>
      <path className="hot a-spark" d="M24 2.5v4M22 4.5h4" />
    </>
  ),
};

export default function CategoryLogo({ category, bulk = false }) {
  const glyph = GLYPHS[category] ?? GLYPHS.Other;
  return (
    <span className={`cat-logo cat-${familyOf(category)}${bulk ? ' cat-bulk' : ''}`}>
      <svg viewBox="0 0 32 32" aria-hidden="true" focusable="false">{glyph}</svg>
    </span>
  );
}
