/**
 * The icon set, inline.
 *
 * The SI shell this layout comes from draws its icons from `lucide-react`.
 * Pulling a whole icon package in for the twenty-odd glyphs used here would be
 * the largest dependency in this app, so the glyphs are transcribed as paths
 * instead — same 24px stroke grid, same visual weight, no install step.
 */

const BASE = {
  viewBox: '0 0 24 24',
  fill: 'none',
  stroke: 'currentColor',
  strokeWidth: 2,
  strokeLinecap: 'round',
  strokeLinejoin: 'round',
};

function make(displayName, children) {
  const Icon = ({ size = 16, ...rest }) => (
    <svg width={size} height={size} aria-hidden="true" focusable="false" {...BASE} {...rest}>
      {children}
    </svg>
  );
  Icon.displayName = displayName;
  return Icon;
}

export const LayoutDashboard = make('LayoutDashboard', (
  <>
    <rect x="3" y="3" width="7" height="9" rx="1" />
    <rect x="14" y="3" width="7" height="5" rx="1" />
    <rect x="14" y="12" width="7" height="9" rx="1" />
    <rect x="3" y="16" width="7" height="5" rx="1" />
  </>
));

export const ClipboardList = make('ClipboardList', (
  <>
    <rect x="8" y="2" width="8" height="4" rx="1" />
    <path d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />
    <path d="M8 11h8M8 15h5" />
  </>
));

export const FilePlus = make('FilePlus', (
  <>
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <polyline points="14 2 14 8 20 8" />
    <path d="M12 18v-6M9 15h6" />
  </>
));

export const CheckSquare = make('CheckSquare', (
  <>
    <polyline points="9 11 12 14 22 4" />
    <path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />
  </>
));

export const Menu = make('Menu', <path d="M3 6h18M3 12h18M3 18h18" />);

export const X = make('X', <path d="M18 6 6 18M6 6l12 12" />);

export const Search = make('Search', (
  <>
    <circle cx="11" cy="11" r="8" />
    <path d="m21 21-4.35-4.35" />
  </>
));

export const LogOut = make('LogOut', (
  <>
    <path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" />
    <polyline points="16 17 21 12 16 7" />
    <path d="M21 12H9" />
  </>
));

export const Sun = make('Sun', (
  <>
    <circle cx="12" cy="12" r="4.5" />
    <path d="M12 1.5v2M12 20.5v2M4.2 4.2l1.4 1.4M18.4 18.4l1.4 1.4M1.5 12h2M20.5 12h2M4.2 19.8l1.4-1.4M18.4 5.6l1.4-1.4" />
  </>
));

export const Moon = make('Moon', <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />);

export const ChevronRight = make('ChevronRight', <polyline points="9 18 15 12 9 6" />);

export const ArrowLeft = make('ArrowLeft', (
  <>
    <path d="M19 12H5" />
    <polyline points="12 19 5 12 12 5" />
  </>
));

export const RefreshCw = make('RefreshCw', (
  <>
    <path d="M21 12a9 9 0 0 1-15.5 6.2L3 16" />
    <path d="M3 12a9 9 0 0 1 15.5-6.2L21 8" />
    <polyline points="3 21 3 16 8 16" />
    <polyline points="21 3 21 8 16 8" />
  </>
));

export const AlertTriangle = make('AlertTriangle', (
  <>
    <path d="M10.3 3.6 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.6a2 2 0 0 0-3.4 0z" />
    <path d="M12 9v4M12 17h.01" />
  </>
));

export const UserPlus = make('UserPlus', (
  <>
    <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
    <circle cx="9" cy="7" r="4" />
    <path d="M19 8v6M22 11h-6" />
  </>
));

export const UserMinus = make('UserMinus', (
  <>
    <path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />
    <circle cx="9" cy="7" r="4" />
    <path d="M22 11h-6" />
  </>
));

export const Users = make('Users', (
  <>
    <path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2" />
    <circle cx="9" cy="7" r="4" />
    <path d="M23 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" />
  </>
));

export const Calendar = make('Calendar', (
  <>
    <rect x="3" y="4" width="18" height="18" rx="2" />
    <path d="M16 2v4M8 2v4M3 10h18" />
  </>
));

export const Clock = make('Clock', (
  <>
    <circle cx="12" cy="12" r="9" />
    <polyline points="12 7 12 12 15.5 14" />
  </>
));

export const Laptop = make('Laptop', (
  <>
    <rect x="3" y="4" width="18" height="12" rx="2" />
    <path d="M2 20h20" />
  </>
));

export const Share2 = make('Share2', (
  <>
    <circle cx="18" cy="5" r="3" />
    <circle cx="6" cy="12" r="3" />
    <circle cx="18" cy="19" r="3" />
    <path d="m8.6 13.5 6.8 4M15.4 6.5l-6.8 4" />
  </>
));

export const Plus = make('Plus', <path d="M12 5v14M5 12h14" />);

export const Filter = make('Filter', <polygon points="22 3 2 3 10 12.46 10 19 14 21 14 12.46 22 3" />);

export const Pencil = make('Pencil', (
  <>
    <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
    <path d="M18.5 2.5a2.12 2.12 0 0 1 3 3L12 15l-4 1 1-4z" />
  </>
));

export const ShieldCheck = make('ShieldCheck', (
  <>
    <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z" />
    <polyline points="9 12 11 14 15 10" />
  </>
));

export const Copy = make('Copy', (
  <>
    <rect x="9" y="9" width="13" height="13" rx="2" />
    <path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1" />
  </>
));

export const Download = make('Download', (
  <>
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <path d="M12 15V3" />
  </>
));

export const Building = make('Building', (
  <>
    <rect x="4" y="2" width="16" height="20" rx="2" />
    <path d="M9 22v-4h6v4M8 6h.01M12 6h.01M16 6h.01M8 10h.01M12 10h.01M16 10h.01M8 14h.01M12 14h.01M16 14h.01" />
  </>
));

export const Inbox = make('Inbox', (
  <>
    <polyline points="22 12 16 12 14 15 10 15 8 12 2 12" />
    <path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z" />
  </>
));

export const Check = make('Check', <polyline points="20 6 9 17 4 12" />);

export const HardDrive = make('HardDrive', (
  <>
    <line x1="22" y1="12" x2="2" y2="12" />
    <path d="M5.45 5.11 2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z" />
    <line x1="6" y1="16" x2="6.01" y2="16" />
    <line x1="10" y1="16" x2="10.01" y2="16" />
  </>
));

export const Cpu = make('Cpu', (
  <>
    <rect x="4" y="4" width="16" height="16" rx="2" />
    <rect x="9" y="9" width="6" height="6" />
    <path d="M9 2v2M15 2v2M9 20v2M15 20v2M2 9h2M2 15h2M20 9h2M20 15h2" />
  </>
));

export const MemoryStick = make('MemoryStick', (
  <>
    <path d="M4 15V7a2 2 0 0 1 2-2h12a2 2 0 0 1 2 2v8" />
    <path d="M2 15h20" />
    <path d="M8 11V9M12 11V9M16 11V9" />
    <path d="M6 19v-4M10 19v-4M14 19v-4M18 19v-4" />
  </>
));
