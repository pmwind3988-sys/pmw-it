import { Link } from 'react-router-dom';
import { Laptop, Monitor } from '../../../components/ui/Icons';
import { statusOf, SPARE, RETIRED } from '../lifecycle/status';
import { formatMYT } from '../../../utils/malaysiaTime';

const TONE = { Critical: 'crit', 'Needs Attention': 'attn', Moderate: 'mod', Optimal: 'ok' };

/** Fixed full scales, so two cards side by side compare at a glance: 15th gen, 32 GB, 1 TB. */
const bars = (device) => [
  { key: 'CPU', value: device.cpuGeneration ?? '—', share: (device.cpuGenerationRank ?? 0) / 15 },
  { key: 'RAM', value: device.installedRamGB ? `${device.installedRamGB} GB` : '—', share: (device.installedRamGB ?? 0) / 32 },
  { key: 'Disk', value: device.storageTotalGB ? `${device.storageTotalGB} GB` : '—', share: (device.storageTotalGB ?? 0) / 1024 },
];

export default function MachineCard({ device }) {
  const status = statusOf(device);
  const tone = TONE[device.fitStatus] ?? 'unk';
  let who = device.owner ?? 'No owner recorded';
  if (status === SPARE) who = 'In IT Stash';
  if (status === RETIRED) who = `Retired${device.statusChangedOn ? ` ${formatMYT(device.statusChangedOn, 'date')}` : ''}`;
  const Glyph = device.deviceType === 'Desktop' ? Monitor : Laptop;

  return (
    <Link to={`/devices/${device.id}`} className={`mc mc-${tone}`} aria-label={`${who}, ${device.computerName}, ${device.fitStatus ?? 'not judged'}`}>
      <span className="mc-head">
        <span className="mc-who">
          <strong>{who}</strong>
          <span className="mc-name">{device.computerName}</span>
        </span>
        <span className="mc-type"><Glyph size={18} /> {device.deviceType ?? 'Unknown'}</span>
      </span>
      <span className="mc-bars">
        {bars(device).map((bar) => (
          <span key={bar.key} className="mc-bar">
            <span className="mc-bar-key">{bar.key}</span>
            <span className="mc-bar-track"><span className="mc-bar-fill" style={{ width: `${Math.min(1, bar.share) * 100}%` }} /></span>
            <span className="mc-bar-value">{bar.value}</span>
          </span>
        ))}
      </span>
      <span className={`mc-verdict mc-verdict-${tone}`}>{device.fitStatus ?? 'Not judged'}</span>
    </Link>
  );
}
