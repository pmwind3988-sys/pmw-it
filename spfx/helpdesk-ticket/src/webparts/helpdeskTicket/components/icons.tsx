// Inline stroke icons on the 24px grid. Decorative unless a label is given.
import * as React from 'react';

interface IconProps {
  size?: number;
  strokeWidth?: number;
  label?: string;
  className?: string;
}

function Svg(props: IconProps & { children: React.ReactNode }): React.ReactElement {
  const { size = 16, strokeWidth = 2, label, className, children } = props;
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={strokeWidth}
      strokeLinecap="round"
      strokeLinejoin="round"
      className={className}
      role={label ? 'img' : undefined}
      aria-label={label}
      aria-hidden={label ? undefined : true}
      focusable="false"
    >
      {children}
    </svg>
  );
}

export const TicketIcon = (p: IconProps): React.ReactElement => (
  <Svg size={p.size ?? 26} strokeWidth={1.8} {...p}>
    <path d="M3 7a2 2 0 0 1 2-2h14a2 2 0 0 1 2 2v3a2 2 0 0 0 0 4v3a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-3a2 2 0 0 0 0-4z" />
    <path d="M9 5v14" strokeDasharray="2 2" />
  </Svg>
);

export const LockIcon = (p: IconProps): React.ReactElement => (
  <Svg {...p}>
    <rect x="4" y="11" width="16" height="10" rx="2" />
    <path d="M8 11V7a4 4 0 0 1 8 0v4" />
  </Svg>
);

export const ClipIcon = (p: IconProps): React.ReactElement => (
  <Svg size={p.size ?? 16} {...p}>
    <path d="M21.4 11.6 12.2 20.8a6 6 0 0 1-8.5-8.5l9.2-9.2a4 4 0 0 1 5.7 5.7l-9.2 9.2a2 2 0 0 1-2.8-2.8l8.5-8.5" />
  </Svg>
);

export const CloseIcon = (p: IconProps): React.ReactElement => (
  <Svg size={p.size ?? 12} strokeWidth={2.5} {...p}>
    <path d="M6 6l12 12M18 6 6 18" />
  </Svg>
);

export const CheckIcon = (p: IconProps): React.ReactElement => (
  <Svg size={p.size ?? 24} strokeWidth={2.5} {...p}>
    <path d="M20 6 9 17l-5-5" />
  </Svg>
);

export const ChevronIcon = (p: IconProps): React.ReactElement => (
  <Svg size={p.size ?? 16} {...p}>
    <path d="M9 6l6 6-6 6" />
  </Svg>
);

export const RefreshIcon = (p: IconProps): React.ReactElement => (
  <Svg size={p.size ?? 18} {...p}>
    <path d="M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7" />
  </Svg>
);

export const SearchIcon = (p: IconProps): React.ReactElement => (
  <Svg size={p.size ?? 16} {...p}>
    <circle cx="11" cy="11" r="7" />
    <path d="M20 20l-4-4" />
  </Svg>
);
