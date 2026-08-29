import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { useSemantic } from '../useSemantic';
import { UNSORTED_ID, UNSORTED_LABEL } from './buckets';
import { SENTIMENT } from './sentiment';

/**
 * Every separated issue, with the category it was given and a dropdown
 * to disagree.
 *
 * An excluded row stays on screen, struck through, rather than
 * disappearing. Removing it outright would leave someone unable to undo
 * a misclick and unable to see how much they had excluded.
 *
 * Severity and Tone are two different questions and both columns are here.
 * Severity is how strongly it is put; tone is which way it points. A survey
 * asking what is wrong still collects "the new laptops are excellent", and a
 * high-severity compliment reads exactly like a complaint without this.
 */

const TONE_CLASS = {
  [SENTIMENT.NEGATIVE]: 'sa-tone-bad',
  [SENTIMENT.POSITIVE]: 'sa-tone-good',
  [SENTIMENT.NEUTRAL]: 'sa-tone-flat',
};
export default function IssueTable() {
  const { analysis, buckets, retagFragment, toggleNoise } = useSemantic();

  if (!analysis || analysis.fragments.length === 0) {
    return <EmptyState>No issues were found in these answers.</EmptyState>;
  }

  const options = [...buckets, { id: UNSORTED_ID, label: UNSORTED_LABEL }];

  return (
    <Card className="sa-table-card">
      <div className="sa-table-scroll">
        <table className="sa-table">
          <thead>
            <tr>
              <th className="sa-num">Row</th>
              <th>Issue</th>
              <th>Category</th>
              <th className="sa-num">Severity</th>
              <th>Tone</th>
              <th>Use</th>
            </tr>
          </thead>
          <tbody>
            {analysis.fragments.map((fragment) => (
              <tr key={fragment.id} className={fragment.noise ? 'sa-issue-noise' : undefined}>
                <td className="sa-num">{fragment.row + 1}</td>
                <td className="sa-issue-text">{fragment.text}</td>
                <td>
                  <select
                    className="sa-select"
                    aria-label={`Category for row ${fragment.row + 1}`}
                    value={fragment.bucketId}
                    onChange={(e) => retagFragment(fragment.id, e.target.value)}
                  >
                    {options.map((b) => <option key={b.id} value={b.id}>{b.label}</option>)}
                  </select>
                </td>
                <td className="sa-num">{Math.round((fragment.severity ?? 0) * 100)}</td>
                <td>
                  {/* A signal, not a verdict — the same caveat severity
                      carries. Nothing here knows that "it works, but only if
                      you restart it twice" is a complaint. */}
                  <span className={TONE_CLASS[fragment.sentiment?.sentiment] ?? 'sa-tone-flat'}>
                    {fragment.sentiment?.sentiment ?? SENTIMENT.NEUTRAL}
                  </span>
                </td>
                <td>
                  <label className="sa-field">
                    <input
                      type="checkbox"
                      checked={!fragment.noise}
                      aria-label={`Count row ${fragment.row + 1}`}
                      onChange={() => toggleNoise(fragment.id)}
                    />
                    <span>{fragment.noise ? 'Excluded' : 'Counted'}</span>
                  </label>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}
