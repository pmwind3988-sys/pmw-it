/**
 * The sign-in screen's idle animation: a board, a chip, and packets relayed out
 * to four nodes. No words, nothing to read, nothing to click.
 *
 * All motion is CSS (see `src/styles/auth.css`) on plain divs — nothing here
 * needs exporting as an asset. It is held still under `prefers-reduced-motion`;
 * the composition still reads.
 *
 * The same tree is dropped behind the card on a phone, where there is no column
 * to give it; `className` is how the caller strips the desktop margin that
 * would otherwise crop it there.
 */
const DIRECTIONS = ['up', 'down', 'left', 'right'];

export default function IdleAnimation({ className = '' }) {
  return (
    <div className={`idle ${className}`.trim()} aria-hidden="true">
      <div className="idle-dots" />

      <div className="idle-stage">
        <div className="idle-board">
          {DIRECTIONS.map((dir) => (
            <div key={`trace-${dir}`} className={`idle-trace ${dir === 'up' || dir === 'down' ? 'v' : 'h'} ${dir}`} />
          ))}

          {DIRECTIONS.map((dir) => (
            <div key={`node-${dir}`} className={`idle-node ${dir}`} />
          ))}

          <div className="idle-ripple" />
          <div className="idle-ripple b" />

          <div className="idle-pins" />
          <div className="idle-chip">
            <div className="idle-core" />
          </div>

          {DIRECTIONS.map((dir) => (
            <div key={`packet-${dir}`} className={`idle-packet ${dir}`} />
          ))}
        </div>
      </div>
    </div>
  );
}
