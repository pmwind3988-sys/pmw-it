import { Card, EmptyState } from '../../../components/ui/Surfaces';
import { useDataStudio } from '../useDataStudio';

/**
 * The groupings the model found on its own.
 *
 * A theme's name is four distinctive words, not a sentence -- the model
 * measures sameness, it does not write. The name field is editable
 * because that is the honest interface for a starting point, and the
 * example lines under each theme are what somebody reads in order to
 * decide what to call it.
 */
export default function ThemeList() {
  const { analysis, renameTheme, mergeThemes } = useDataStudio();

  if (!analysis) return null;
  if (analysis.themes.length === 0) {
    return (
      <EmptyState>
        Too few answers to find themes in. Categories and the issue list still work.
      </EmptyState>
    );
  }

  const textOf = (id) => analysis.fragments.find((f) => f.id === id)?.text ?? '';

  return (
    <div className="ds-theme-list">
      {analysis.themes.map((theme) => (
        <Card key={theme.id} className="ds-theme">
          <div className="ds-bucket-head">
            <input
              className="ds-input ds-bucket-name"
              value={theme.name}
              aria-label="Theme name"
              onChange={(e) => renameTheme(theme.id, e.target.value)}
            />
            <span className="ds-summary">
              {`${theme.count} issues · ${theme.respondents} people`}
            </span>
            <select
              className="ds-select"
              aria-label={`Merge ${theme.name} into another theme`}
              value=""
              onChange={(e) => e.target.value && mergeThemes(theme.id, e.target.value)}
            >
              <option value="">Merge into…</option>
              {analysis.themes
                .filter((other) => other.id !== theme.id)
                .map((other) => (
                  <option key={other.id} value={other.id}>{other.name}</option>
                ))}
            </select>
          </div>
          <ul className="ds-theme-examples">
            {theme.fragmentIds.slice(0, 3).map((id) => <li key={id}>{textOf(id)}</li>)}
          </ul>
        </Card>
      ))}
    </div>
  );
}
