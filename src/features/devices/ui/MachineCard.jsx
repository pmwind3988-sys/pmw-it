import { Link } from 'react-router-dom';
import { Laptop, Monitor } from '../../../components/ui/Icons';
import { statusOf, SPARE, RETIRED, IN_REPAIR } from '../lifecycle/status';
import { PARTS } from '../standards/defaultStandard.js';
import { GRADE_SLUG } from '../standards/gradeColors';

const STATUS_PILL = {
  [SPARE]: { text: 'In IT Stash', tone: 'stash' },
  [RETIRED]: { text: 'Retired', tone: 'retired' },
  [IN_REPAIR]: { text: 'In repair', tone: 'repair' },
};

/** Only the parts that are not fine: Critical and Needs attention. Everything else is left out. */
const notFine = (device) => PARTS
  .map(({ key, label }) => ({ label, grade: device.parts?.[key]?.grade }))
  .filter(({ grade }) => grade === 'Critical' || grade === 'Needs attention');

export default function MachineCard({ device }) {
  const status = statusOf(device);
  const owner = device.owner || 'No owner recorded';
  const pill = STATUS_PILL[status];
  const parts = notFine(device);
  const Glyph = device.deviceType === 'Desktop' ? Monitor : Laptop;
  const partsText = parts.length
    ? parts.map(({ label, grade }) => `${label} ${grade}`).join(', ')
    : 'all parts fine';

  return (
    <Link
      to={`/devices/${device.id}`}
      className="mc"
      aria-label={`${owner}, ${device.computerName ?? 'no computer name'}${pill ? `, ${pill.text}` : ''}, ${partsText}`}
    >
      <span className="mc-head">
        <span className="mc-who">
          <strong>{owner}</strong>
          <span className="mc-name">{device.computerName}</span>
        </span>
        {pill && <span className={`mc-status mc-status-${pill.tone}`}>{pill.text}</span>}
      </span>
      <span className="mc-type"><Glyph size={18} /> {device.deviceType ?? 'Unknown'}</span>
      {parts.length > 0
        ? (
          <span className="pc">
            {parts.map(({ label, grade }) => (
              <span key={label} className={`pc-chip g-${GRADE_SLUG[grade]}`}>{label} · {grade}</span>
            ))}
          </span>
        )
        : <span className="mc-fine">All parts fine</span>}
    </Link>
  );
}
