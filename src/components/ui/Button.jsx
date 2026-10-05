import Spinner from './Spinner';

/**
 * The one button. Variants and sizes are class names on `.ui-btn` — see
 * `src/styles/shell.css`.
 *
 * `loading` is the universal busy state: the spinner takes the icon's place,
 * the label stays so the reader still knows WHAT is happening, and the button
 * is disabled so a second press cannot send the same write twice.
 */
export default function Button({
  children,
  variant = 'primary',
  size = 'md',
  icon: Icon,
  className = '',
  loading = false,
  disabled,
  ...props
}) {
  return (
    <button
      className={`ui-btn ui-btn-${size} ui-btn-${variant}${loading ? ' ui-btn-loading' : ''} ${className}`.trim()}
      disabled={disabled || loading}
      aria-busy={loading || undefined}
      {...props}
    >
      {loading ? <Spinner size={14} /> : Icon && <Icon size={14} />}
      {children}
    </button>
  );
}
