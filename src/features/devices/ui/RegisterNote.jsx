import { Link } from 'react-router-dom';
import Button from '../../../components/ui/Button';
import { assetFor, holderOf, machineMove, describeMove } from '../../assets/machineLink';
import { ACTIONS } from '../lifecycle/planLifecycle';

const BUTTON = {
  [ACTIONS.ASSIGN]: (move) => `Give it to ${move.holder.name}`,
  [ACTIONS.CHANGE_OWNER]: (move) => `Change owner to ${move.holder.name}`,
  [ACTIONS.TO_STASH]: () => 'Put it in the IT Stash',
};

/**
 * The asset-register row that is this same machine (`machineLink.js`), and --
 * when the two disagree -- one press to make the device list follow the
 * register, which is where handovers are recorded.
 */
export default function RegisterNote({ device, assets, onAction, busy }) {
  const asset = assetFor(device, assets);
  if (!asset) return null;
  const holder = holderOf(asset);
  const move = machineMove(asset, device);
  const name = asset.title || [asset.manufacturer, asset.model].filter(Boolean).join(' ') || 'Asset';

  return (
    <div className={`mh-register${move ? ' mh-register-off' : ''}`} role={move ? 'status' : undefined}>
      <p className="mh-register-line">
        <span className="mh-def-k">Asset register</span>{' '}
        <Link to={`/assets/${asset.id}`}>{name}</Link>
        <span className="mh-muted"> · {holder ? `with ${holder.name}` : (asset.status || 'In stock')}</span>
      </p>
      {move && (
        <div className="mh-register-fix">
          <p>{describeMove(move, device)}</p>
          <Button size="sm" loading={busy} onClick={() => onAction(move.action, move.input)}>
            {BUTTON[move.action](move)}
          </Button>
        </div>
      )}
    </div>
  );
}
