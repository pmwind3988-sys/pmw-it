// CSS tokens -> an ECharts theme -- spec §10.6, §10.7.
//
// Every colour is read from the live custom properties rather than
// written here. That is the project rule, and it is also what makes the
// dark-mode flip correct for free: the tokens change, the theme is
// rebuilt from them, and nothing in this file has to know which mode is
// on.
//
// The series palette is read the same way, from `--sa-series-1..8` in
// semantic.css, which is where its colour-blindness reasoning lives.

import echarts from './echartsCore.js';

export const SERIES_SLOTS = 8;

const THEME_NAME = 'pmw-studio';

function token(name, fallback = '') {
  if (typeof document === 'undefined') return fallback;
  const value = getComputedStyle(document.documentElement).getPropertyValue(name);
  return value.trim() || fallback;
}

export function seriesPalette() {
  return Array.from({ length: SERIES_SLOTS }, (_, i) => token(`--sa-series-${i + 1}`))
    .filter(Boolean);
}

function prefersReducedMotion() {
  return typeof window !== 'undefined'
    && typeof window.matchMedia === 'function'
    && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
}

export function buildTheme() {
  const ink = token('--it-ink', '#101828');
  const inkSoft = token('--it-ink-soft', '#64748b');
  const line = token('--it-line', '#e3e9f0');
  const panel = token('--it-panel', '#ffffff');

  // Honoured here rather than per chart: a dashboard of a dozen tiles
  // all easing at once is precisely the motion someone with the
  // preference set is asking not to see.
  const animate = !prefersReducedMotion();

  const axis = {
    axisLine: { lineStyle: { color: line } },
    axisTick: { lineStyle: { color: line } },
    axisLabel: { color: inkSoft },
    splitLine: { lineStyle: { color: line, type: 'dashed' } },
    splitArea: { show: false },
  };

  return {
    color: seriesPalette(),
    // Transparent, not `--it-panel`: the tile card already paints the
    // panel colour, and a second opaque layer inside it puts a visible
    // square over the card's rounded corners.
    backgroundColor: 'transparent',
    textStyle: { color: ink },
    animation: animate,
    animationDuration: animate ? 300 : 0,
    title: {
      textStyle: { color: ink },
      subtextStyle: { color: inkSoft },
    },
    legend: {
      textStyle: { color: inkSoft },
      inactiveColor: line,
    },
    tooltip: {
      backgroundColor: panel,
      borderColor: line,
      textStyle: { color: ink },
      axisPointer: {
        lineStyle: { color: inkSoft },
        crossStyle: { color: inkSoft },
      },
    },
    categoryAxis: { ...axis },
    valueAxis: { ...axis },
    timeAxis: { ...axis },
    logAxis: { ...axis },
  };
}

/**
 * Registers the theme under a name and returns it.
 *
 * ECharts bakes a theme at `init` and offers no way to swap it on a live
 * instance, so a theme flip means re-registering (the tokens have
 * changed) and re-initialising. `EChart` does exactly that.
 */
export function registerStudioTheme() {
  echarts.registerTheme(THEME_NAME, buildTheme());
  return THEME_NAME;
}
