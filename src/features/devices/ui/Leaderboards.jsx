import { Link } from 'react-router-dom';
import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { formatMYT } from '../../../utils/malaysiaTime';
import { leaderboards } from '../stats/deviceStats';

function Board({ title, blurb, rows, render, emptyText }) {
  return (
    <Card className="lb-card todo-board">
      <div className="chart-head">
        <h3>{title}</h3>
        <p>{blurb}</p>
      </div>
      {rows.length === 0 ? (
        <EmptyState>{emptyText}</EmptyState>
      ) : (
        <ul className="lb-list">
          {rows.map((device) => (
            <li key={device.id ?? device.computerName}>
              <span className="lb-name">
                {device.id != null ? (
                  <Link to={`/devices/${device.id}`}>{device.computerName}</Link>
                ) : device.computerName}
                {device.owner && <span className="lb-owner"> · {device.owner}</span>}
              </span>
              <span className="lb-figure">{render(device)}</span>
            </li>
          ))}
        </ul>
      )}
    </Card>
  );
}

/** The two boards the To do view keeps: the cheap upgrade and the stale scan. */
export default function Leaderboards({ devices }) {
  const boards = leaderboards(devices);

  return (
    <div className="lb-grid todo-boards">
      <Board
        title="Cheap RAM upgrades"
        blurb="8 GB or less with a slot free — one stick, not a new machine."
        rows={boards.upgradeCandidates}
        render={(d) => `${d.installedRamGB} GB in ${d.ramSlotsUsed} of ${d.ramSlotsTotal} slots`}
        emptyText="Nothing upgradable on the cheap."
      />

      <Board
        title="Needs a re-scan"
        blurb="Reports that failed, and machines not seen in six months."
        rows={boards.rescanNeeded}
        render={(d) => (d.scanComplete === false
          ? 'Scan incomplete'
          : `Last seen ${formatMYT(d.scannedOn, 'date')}`)}
        emptyText="Every machine has a recent, complete scan."
      />
    </div>
  );
}
