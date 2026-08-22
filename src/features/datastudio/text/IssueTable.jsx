import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { useDataStudio } from '../useDataStudio';
import { UNSORTED_ID, UNSORTED_LABEL } from './buckets';

/**
 * Every separated issue, with the category it was given and a dropdown
 * to disagree.
 *
 * An excluded row stays on screen, struck through, rather than
 * disappearing. Removing it outright would leave someone unable to undo
 * a misclick and unable to see how much they had excluded.
 */
export default function IssueTable() {
  const { analysis, buckets, retagFragment, toggleNoise } = useDataStudio();

  if (!analysis || analysis.fragments.length === 0) {
    return <EmptyState>No issues were found in these answers.</EmptyState>;
  }

  const options = [...buckets, { id: UNSORTED_ID, label: UNSORTED_LABEL }];

  return (
    <Card className="ds-table-card">
      <div className="ds-table-scroll">
        <table className="ds-table">
          <thead>
            <tr>
              <th className="ds-num">Row</th>
              <th>Issue</th>
              <th>Category</th>
              <th className="ds-num">Severity</th>
              <th>Use</th>
            </tr>
          </thead>
          <tbody>
            {analysis.fragments.map((fragment) => (
              <tr key={fragment.id} className={fragment.noise ? 'ds-issue-noise' : undefined}>
                <td className="ds-num">{fragment.row + 1}</td>
                <td className="ds-issue-text">{fragment.text}</td>
                <td>
                  <select
                    className="ds-select"
                    aria-label={`Category for row ${fragment.row + 1}`}
                    value={fragment.bucketId}
                    onChange={(e) => retagFragment(fragment.id, e.target.value)}
                  >
                    {options.map((b) => <option key={b.id} value={b.id}>{b.label}</option>)}
                  </select>
                </td>
                <td className="ds-num">{Math.round((fragment.severity ?? 0) * 100)}</td>
                <td>
                  <label className="ds-field">
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
