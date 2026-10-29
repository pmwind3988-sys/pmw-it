// Glyphs the queue needs beyond the helpdesk icons. Same 24px stroke grid.
import * as React from 'react';

interface IconProps {
  size?: number;
  className?: string;
}

function Svg(props: IconProps & { children: React.ReactNode }): React.ReactElement {
  const { size = 16, className, children } = props;
  return (
    <svg
      width={size}
      height={size}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeWidth={2}
      strokeLinecap="round"
      strokeLinejoin="round"
      className={className}
      aria-hidden="true"
      focusable="false"
    >
      {children}
    </svg>
  );
}

export const DownloadIcon = (p: IconProps): React.ReactElement => (
  <Svg {...p}>
    <path d="M12 4v11M7 10l5 5 5-5M5 20h14" />
  </Svg>
);

export const ChevronDownIcon = (p: IconProps): React.ReactElement => (
  <Svg {...p}>
    <path d="M6 9l6 6 6-6" />
  </Svg>
);

/** Up for ascending, down for descending; both faint when the column is not the sort. */
export const SortIcon = (p: IconProps & { direction?: 'asc' | 'desc' }): React.ReactElement => {
  if (p.direction === 'asc') {
    return (
      <Svg {...p}>
        <path d="M12 19V5M6 11l6-6 6 6" />
      </Svg>
    );
  }
  if (p.direction === 'desc') {
    return (
      <Svg {...p}>
        <path d="M12 5v14M6 13l6 6 6-6" />
      </Svg>
    );
  }
  return (
    <Svg {...p}>
      <path d="M8 9l4-4 4 4M8 15l4 4 4-4" />
    </Svg>
  );
};
