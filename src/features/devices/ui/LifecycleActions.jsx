import { useEffect, useRef, useState } from 'react';
import { useNavigate } from 'react-router-dom';
import Button from '../../../components/ui/Button';
import { ChevronDown } from '../../../components/ui/Icons';
import AssignDialog from './AssignDialog';
import { ACTIONS, actionsFor } from '../lifecycle/planLifecycle';
import { statusOf, inFleet, SPARE } from '../lifecycle/status';
import { useConfirm } from '../../../components/ui/useConfirm';

const LABEL = {
  [ACTIONS.CHANGE_OWNER]: 'Change owner',
  [ACTIONS.TO_REPAIR]: 'Send to repair',
  [ACTIONS.BACK_IN_USE]: 'Back in use',
  [ACTIONS.TO_STASH]: 'Back to the IT Stash',
  [ACTIONS.RETIRE]: 'Retire…',
  [ACTIONS.ASSIGN]: 'Assign / bring back',
};

const TILL = 'till';

/**
 * The one obvious next step as a button, and everything else under More. The
 * till is the same move with a checklist attached, so it is offered beside
 * the status actions; Retire is always last and always asks first.
 */
export default function LifecycleActions({
  device, owners, locations, departments, onAction, busy, showTones, onToggleTones,
}) {
  const [dialog, setDialog] = useState(null);
  // Which control started the work, so only THAT one spins while `busy`.
  const [pressed, setPressed] = useState(null);
  const [menuOpen, setMenuOpen] = useState(false);
  const wrapRef = useRef(null);
  const moreRef = useRef(null);
  const { ask, dialog: confirm } = useConfirm();
  const navigate = useNavigate();

  const status = statusOf(device);
  const tillMode = status === SPARE ? 'out' : (inFleet(device) ? 'back' : null);
  const tillHref = tillMode && device.id != null
    ? `/assets/till?mode=${tillMode}&machine=${encodeURIComponent(device.id)}`
    : null;
  const allowed = actionsFor(device) ?? [];

  // A spare goes out on a checklist; everything else starts with its first status action.
  const primary = tillHref && status === SPARE ? TILL : allowed[0] ?? null;
  const menuActions = allowed
    .filter((action) => action !== primary)
    .sort((a, b) => (a === ACTIONS.RETIRE) - (b === ACTIONS.RETIRE));
  const pressedInMenu = pressed !== null && pressed !== primary;

  useEffect(() => {
    if (!menuOpen) return undefined;
    const onPointer = (event) => {
      if (wrapRef.current && !wrapRef.current.contains(event.target)) setMenuOpen(false);
    };
    const onKey = (event) => {
      if (event.key !== 'Escape') return;
      setMenuOpen(false);
      moreRef.current?.focus();
    };
    document.addEventListener('pointerdown', onPointer);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('pointerdown', onPointer);
      document.removeEventListener('keydown', onKey);
    };
  }, [menuOpen]);

  const press = async (action) => {
    setMenuOpen(false);
    if (action === ACTIONS.CHANGE_OWNER || action === ACTIONS.ASSIGN) {
      setPressed(action);
      setDialog(action);
      return;
    }
    if (action === ACTIONS.RETIRE) {
      const yes = await ask({
        title: `Retire ${device.computerName}?`,
        body: 'It goes to the Graveyard and stops counting in the fleet. Its record and history are kept, and it can be brought back.',
        confirmLabel: 'Retire',
        cancelLabel: 'Keep it in use',
      });
      if (!yes) return;
    }
    setPressed(action);
    onAction(action, {});
  };

  const renderPrimary = () => {
    if (!primary) return null;
    if (primary === TILL) {
      return (
        <Button variant="primary" size="sm" disabled={busy} onClick={() => navigate(tillHref)}>
          Hand out on a checklist
        </Button>
      );
    }
    return (
      <Button variant="primary" size="sm" disabled={busy && pressed !== primary}
        loading={busy && pressed === primary} onClick={() => press(primary)}>
        {LABEL[primary]}
      </Button>
    );
  };

  const tillMenuItem = tillHref && primary !== TILL && (
    <button type="button" role="menuitem" className="la-item" disabled={busy}
      onClick={() => { setMenuOpen(false); navigate(tillHref); }}>
      {tillMode === 'back' ? 'Take back at the till' : 'Hand out on a checklist'}
    </button>
  );

  return (
    <div className="la" ref={wrapRef}>
      {renderPrimary()}
      <div className="la-more">
        <Button ref={moreRef} variant="secondary" size="sm" icon={ChevronDown}
          aria-haspopup="menu" aria-expanded={menuOpen}
          disabled={busy && !pressedInMenu} loading={busy && pressedInMenu}
          onClick={() => setMenuOpen((open) => !open)}>
          More
        </Button>
        {menuOpen && (
          <div className="la-menu" role="menu" aria-label="More actions">
            {menuActions.filter((action) => action !== ACTIONS.RETIRE).map((action) => (
              <button key={action} type="button" role="menuitem" disabled={busy}
                className="la-item"
                onClick={() => press(action)}>
                {LABEL[action]}
              </button>
            ))}
            {tillMenuItem}
            {onToggleTones && (
              <button type="button" role="menuitem" className="la-item la-item-divider"
                onClick={() => { setMenuOpen(false); onToggleTones(); }}>
                Colour values: {showTones ? 'on' : 'off'}
              </button>
            )}
            {/* Retire goes last, apart from the rest: the one choice here that takes a machine out of the fleet. */}
            {menuActions.filter((action) => action === ACTIONS.RETIRE).map((action) => (
              <button key={action} type="button" role="menuitem" disabled={busy}
                className="la-item la-item-danger la-item-divider"
                onClick={() => press(action)}>
                {LABEL[action]}
              </button>
            ))}
          </div>
        )}
      </div>
      {dialog && (
        <AssignDialog
          title={dialog === ACTIONS.ASSIGN ? `Bring back ${device.computerName}` : `Change owner of ${device.computerName}`}
          device={device} owners={owners} locations={locations} departments={departments} busy={busy}
          onCancel={() => setDialog(null)}
          onSubmit={(input) => { setDialog(null); onAction(dialog, input); }}
        />
      )}
      {confirm}
    </div>
  );
}
