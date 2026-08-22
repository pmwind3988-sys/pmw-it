import { useCallback, useEffect, useMemo } from 'react';
import { createMaskCache } from '../engine/filterMask.js';
import { TILE_SIZES } from '../dataStudioStore.js';
import { useDataStudio } from '../useDataStudio';
import ChartTile from './ChartTile.jsx';
import {
  ChevronLeft, ChevronRight, Copy, Pencil, Trash2, Maximize2, Download,
} from '../../../components/ui/Icons';

function TileChrome({ tile, index, total, children, onExport }) {
  const {
    moveTile, cycleTileSize, removeTile, duplicateTile, setEditingTile,
  } = useDataStudio();

  return (
    <section
      className={`ds-tile ds-tile-${tile.chart}`}
      style={{ '--ds-tile-span': TILE_SIZES[tile.size] ?? TILE_SIZES.M }}
      aria-label={tile.title}
    >
      <header className="ds-tile-head">
        <h3 className="ds-tile-title">{tile.title}</h3>
        {/* Real buttons with labels, not icon divs: keyboard reordering
            is the whole reordering story here -- there is no drag
            engine -- so these have to be reachable and announced. */}
        <div className="ds-tile-actions">
          <button
            type="button"
            aria-label={`Move ${tile.title} earlier`}
            disabled={index === 0}
            onClick={() => moveTile(tile.id, -1)}
          >
            <ChevronLeft size={14} />
          </button>
          <button
            type="button"
            aria-label={`Move ${tile.title} later`}
            disabled={index === total - 1}
            onClick={() => moveTile(tile.id, 1)}
          >
            <ChevronRight size={14} />
          </button>
          <button
            type="button"
            aria-label={`Resize ${tile.title} (currently ${tile.size ?? 'M'})`}
            onClick={() => cycleTileSize(tile.id)}
          >
            <Maximize2 size={14} />
          </button>
          <button
            type="button"
            aria-label={`Edit ${tile.title}`}
            onClick={() => setEditingTile(tile.id)}
          >
            <Pencil size={14} />
          </button>
          <button
            type="button"
            aria-label={`Duplicate ${tile.title}`}
            onClick={() => duplicateTile(tile.id)}
          >
            <Copy size={14} />
          </button>
          {onExport && (
            <button
              type="button"
              aria-label={`Download ${tile.title} as an image`}
              onClick={() => onExport(tile)}
            >
              <Download size={14} />
            </button>
          )}
          <button
            type="button"
            aria-label={`Remove ${tile.title}`}
            onClick={() => removeTile(tile.id)}
          >
            <Trash2 size={14} />
          </button>
        </div>
      </header>
      {children}
    </section>
  );
}

export default function CanvasGrid({ onExport, onChartInit }) {
  const {
    tiles, dataset, globalFilters, selection, selectMark, clearSelection,
  } = useDataStudio();

  // One cache for the whole canvas. Because it keys on whether the
  // selection APPLIES rather than on the tile id (see filterMask.js),
  // N tiles share at most two arrays: one for the selection's source
  // tile and one for everyone else.
  const cache = useMemo(() => createMaskCache(), []);

  // An unfiltered mask for tiles that opted out of filtering. Built once
  // per dataset rather than per tile.
  const unfiltered = useMemo(
    () => (dataset ? new Uint8Array(dataset.rowCount).fill(1) : null),
    [dataset],
  );

  // Escape clears the cross-filter from anywhere on the canvas. Without
  // it the only way out is finding the mark you clicked, which on a
  // filtered dashboard may no longer be on screen.
  useEffect(() => {
    if (!selection) return undefined;
    const onKey = (e) => {
      if (e.key === 'Escape') clearSelection();
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [selection, clearSelection]);

  // Resolved up front rather than inside the render loop, so the render
  // body stays a plain mapping over data.
  const masks = useMemo(() => {
    if (!dataset) return new Map();
    return new Map(tiles.map((tile) => [
      tile.id,
      tile.respondsToFilters === false
        ? unfiltered
        : cache.get(dataset, globalFilters, selection, tile.id),
    ]));
  }, [cache, dataset, tiles, globalFilters, selection, unfiltered]);

  const handleSelect = useCallback(({ tileId, column, value, additive }) => {
    selectMark({ tileId, column, value, additive });
  }, [selectMark]);

  if (!dataset) return null;

  return (
    <div className="ds-canvas">
      {tiles.map((tile, index) => (
        <TileChrome
          key={tile.id}
          tile={tile}
          index={index}
          total={tiles.length}
          onExport={onExport}
        >
          <ChartTile
            tile={tile}
            dataset={dataset}
            mask={masks.get(tile.id)}
            selection={selection}
            onSelect={handleSelect}
            onChartInit={onChartInit}
          />
        </TileChrome>
      ))}
    </div>
  );
}
