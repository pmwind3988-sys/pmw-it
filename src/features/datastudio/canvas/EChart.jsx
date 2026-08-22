import { useEffect, useRef } from 'react';
import echarts from './echartsCore.js';
import { registerStudioTheme } from './echartsTheme.js';
import { useTheme } from '../../../context/ThemeContext.jsx';

/**
 * A thin React wrapper around one ECharts instance.
 *
 * `onEvents` must be memoised by callers with `useMemo`, or the binding
 * effect tears down and re-binds every render.
 *
 * `onInit` hands the instance back to the caller. Export needs a real
 * chart object to call `getDataURL` on, and nothing else exposes one.
 */
export default function EChart({ option, onEvents, onInit, className }) {
  const hostRef = useRef(null);
  const chartRef = useRef(null);
  const { isDarkMode } = useTheme();

  // The latest option and callback, readable from the init effect
  // without making either a dependency of it -- otherwise every option
  // change would dispose and rebuild the whole chart.
  //
  // Synced in an effect rather than assigned during render, which React
  // forbids. This effect is declared FIRST on purpose: effects run in
  // declaration order, so the refs are current before the init effect
  // below reads them, on the first mount and on every theme flip.
  const optionRef = useRef(option);
  const onInitRef = useRef(onInit);

  useEffect(() => {
    optionRef.current = option;
    onInitRef.current = onInit;
  });

  useEffect(() => {
    const host = hostRef.current;
    if (!host) return undefined;

    // Re-registering on every init is what makes the theme flip work:
    // `buildTheme` re-reads the CSS custom properties, which the flip
    // has just changed.
    const name = registerStudioTheme();
    const chart = echarts.init(host, name, { renderer: 'canvas' });
    chartRef.current = chart;

    // The option effect below will not re-run on a theme flip, because
    // the option object did not change -- so a rebuilt instance has to
    // be given its option here or the tile comes back blank.
    if (optionRef.current) chart.setOption(optionRef.current, { notMerge: true });
    onInitRef.current?.(chart);

    const ro = new ResizeObserver(() => chart.resize());
    ro.observe(host);

    return () => {
      ro.disconnect();
      chart.dispose();
      chartRef.current = null;
    };
  }, [isDarkMode]); // theme flip rebuilds the instance -- themes are baked at init

  useEffect(() => {
    // `notMerge` is required, not tidiness: merging leaves the previous
    // option's extra series in place when a tile's series count shrinks,
    // so a chart that used to have four lines keeps drawing the fourth.
    chartRef.current?.setOption(option, { notMerge: true });
  }, [option]);

  useEffect(() => {
    const chart = chartRef.current;
    if (!chart || !onEvents) return undefined;
    const entries = Object.entries(onEvents);
    entries.forEach(([evt, handler]) => chart.on(evt, handler));
    return () => entries.forEach(([evt, handler]) => chart.off(evt, handler));
  }, [onEvents, isDarkMode]);

  return <div ref={hostRef} className={className} />;
}
