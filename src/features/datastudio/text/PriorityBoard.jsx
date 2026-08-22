import { Card, EmptyState } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { useDataStudio } from '../useDataStudio';

/**
 * The ranked list, and the one screen that has to explain itself.
 *
 * "People" comes before "Mentions" in the header because that is the
 * order the score works in: five people with one mild complaint each
 * outrank one person with five furious ones. Leading with the mention
 * count would suggest the opposite ordering and make the list look
 * wrong to anyone who read it carefully.
 */
export default function PriorityBoard() {
  const { analysis, togglePin, toggleSuppress } = useDataStudio();

  if (!analysis || analysis.priority.length === 0) {
    return <EmptyState>Nothing to rank yet.</EmptyState>;
  }

  return (
    <Card className="ds-table-card">
      <p className="ds-summary">
        Ranked by how many different people raised it, scaled by how strongly they wrote.
        Severity is a signal from the wording, not a judgement.
      </p>
      <div className="ds-table-scroll">
        <table className="ds-table">
          <thead>
            <tr>
              <th className="ds-num">#</th>
              <th>Issue</th>
              <th>From</th>
              <th className="ds-num">People</th>
              <th className="ds-num">Mentions</th>
              <th className="ds-num">Severity</th>
              <th aria-label="Actions" />
            </tr>
          </thead>
          <tbody>
            {analysis.priority.map((item, i) => (
              <tr
                key={`${item.kind}:${item.id}`}
                className={item.suppressed ? 'ds-issue-noise' : undefined}
              >
                <td className="ds-num">{i + 1}</td>
                <td>{item.label}</td>
                <td>
                  <span className="ds-summary">
                    {item.kind === 'bucket' ? 'Category' : 'Theme'}
                  </span>
                </td>
                <td className="ds-num">{item.respondents}</td>
                <td className="ds-num">{item.count}</td>
                <td className="ds-num">{Math.round(item.meanSeverity * 100)}</td>
                <td className="ds-priority-actions">
                  <Button variant="ghost" size="sm" onClick={() => togglePin(item.id)}>
                    {item.pinned ? 'Unpin' : 'Pin'}
                  </Button>
                  <Button variant="ghost" size="sm" onClick={() => toggleSuppress(item.id)}>
                    {item.suppressed ? 'Restore' : 'Hide'}
                  </Button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}
