import { useMemo, useState } from 'react';
import { tillSearch } from '../tillSearch';

/**
 * "Same as one we have": an unknown box is a model the register already
 * holds. Picking it puts the box on that model's line and remembers the box's
 * barcode, so the next one is recognised on sight.
 */
export default function SameAsSearch({ assets, onPick }) {
  const [query, setQuery] = useState('');
  const results = useMemo(() => tillSearch('in', query, { assets }), [query, assets]);

  return (
    <div className="till-sameas">
      <input
        className="till-inline"
        value={query}
        onChange={(event) => setQuery(event.target.value)}
        placeholder="Same as one we have? Type the model"
        aria-label="Same as a model we already have"
      />
      {results.length > 0 && (
        <div className="till-chips">
          {results.map((result) => (
            <button key={result.id} type="button" className="till-chip" onClick={() => onPick(result.asset)}>
              {result.name}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
