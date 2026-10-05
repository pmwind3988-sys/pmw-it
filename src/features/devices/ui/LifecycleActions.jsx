import { useState } from 'react';
import Button from '../../../components/ui/Button';
import AssignDialog from './AssignDialog';
import { ACTIONS, actionsFor } from '../lifecycle/planLifecycle';
import { useConfirm } from '../../../components/ui/useConfirm';

const LABEL = {
  [ACTIONS.CHANGE_OWNER]: 'Change owner',
  [ACTIONS.TO_REPAIR]: 'To repair',
  [ACTIONS.BACK_IN_USE]: 'Back in use',
  [ACTIONS.TO_STASH]: 'To IT Stash',
  [ACTIONS.RETIRE]: 'Retire…',
  [ACTIONS.ASSIGN]: 'Assign / bring back',
};

/** The buttons a machine's status allows, and the one dialog they share. */
export default function LifecycleActions({
  device, owners, locations, departments, onAction, busy,
}) {
  const [dialog, setDialog] = useState(null);
  // Which button started the work, so only THAT one spins while `busy`.
  const [pressed, setPressed] = useState(null);
  const { ask, dialog: confirm } = useConfirm();

  const press = async (action) => {
    if (action === ACTIONS.CHANGE_OWNER || action === ACTIONS.ASSIGN) {
      setPressed(action);
      setDialog(action);
      return;
    }
    setPressed(action);
    if (action === ACTIONS.RETIRE) {
      const yes = await ask({
        title: `Retire ${device.computerName}?`,
        body: 'It goes to the Graveyard and stops counting in the fleet. Its record and history are kept, and it can be brought back.',
        confirmLabel: 'Retire',
        cancelLabel: 'Keep it in use',
      });
      if (!yes) return;
    }
    onAction(action, {});
  };

  return (
    <div className="la">
      {actionsFor(device).map((action, index) => (
        <Button key={action} variant={index === 0 ? 'primary' : 'secondary'} size="sm"
          className={action === ACTIONS.RETIRE ? 'la-danger' : undefined}
          disabled={busy && pressed !== action} loading={busy && pressed === action}
          onClick={() => press(action)}>
          {LABEL[action]}
        </Button>
      ))}
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
