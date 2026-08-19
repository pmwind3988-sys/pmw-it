import { UserPlus, UserMinus, ClipboardList, ShieldCheck } from './Icons';

/**
 * Colours are keyed off the request type, so the same green/red reading carries
 * from the dashboard cards through the list rows to the form's mode badge.
 */
const TYPE_MAP = {
  onboarding: { c: '#12a150', Icon: UserPlus, label: 'Onboarding' },
  in: { c: '#12a150', Icon: UserPlus, label: 'Onboarding' },
  offboarding: { c: '#dc2626', Icon: UserMinus, label: 'Offboarding' },
  out: { c: '#dc2626', Icon: UserMinus, label: 'Offboarding' },
  'individual request': { c: '#0078d4', Icon: ClipboardList, label: 'Individual request' },
};

export function RequestTypeBadge({ type, showIcon = true }) {
  const key = String(type || '').toLowerCase();
  const cfg = TYPE_MAP[key];
  if (!cfg) {
    return (
      <span className="ui-pill" style={{ background: 'var(--it-canvas)', color: 'var(--it-ink-soft)' }}>
        <span>{type || '—'}</span>
      </span>
    );
  }
  const { c, Icon, label } = cfg;
  return (
    <span
      className="ui-pill"
      style={{ background: `${c}14`, color: c, borderColor: `${c}45` }}
      title={label}
    >
      {showIcon && <Icon size={12} />}
      <span>{label}</span>
    </span>
  );
}

/**
 * Who is signed in, in the bar. The SI shell puts a role badge here; this app
 * has no roles of its own — everyone reaches it through the same Microsoft 365
 * work account — so the badge says which account rather than which role.
 *
 * `compact` drops the text below 400px and leaves the mark. On a 360px phone,
 * beside the hamburger and the theme toggle, the name is what pushes the row
 * past the viewport.
 */
export function AccountBadge({ name, compact = false }) {
  const label = name || 'Signed in';
  return (
    <span
      className="ui-pill"
      style={{
        background: 'var(--it-brand-wash)',
        color: 'var(--it-brand)',
        borderColor: 'color-mix(in srgb, var(--it-brand) 30%, transparent)',
      }}
      aria-label={label}
      title={label}
    >
      <ShieldCheck size={12} />
      <span className={compact ? 'hide-below-xs' : ''}>{label}</span>
    </span>
  );
}
