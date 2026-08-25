import { Card } from '../../../components/ui/Surfaces';
import Button from '../../../components/ui/Button';
import { Plus, Trash2 } from '../../../components/ui/Icons';
import { useSemantic } from '../useSemantic';

/**
 * Where the categories are defined.
 *
 * The DESCRIPTION is the field that does the work, which is why it gets
 * the larger control and why the toolbar says so. Answers are matched
 * against these sentences, not against the names -- rename "SAP / ERP"
 * to "The Big System" and nothing about what lands in it changes.
 * Saying that on screen saves somebody an afternoon of renaming things
 * and wondering why nothing moved.
 */
export default function BucketEditor() {
  const {
    buckets, updateBucket, addBucket, removeBucket,
    textSettings, setTextSetting, analysis,
  } = useSemantic();

  const countOf = (id) => analysis?.buckets.find((b) => b.id === id)?.count ?? 0;

  return (
    <Card className="sa-text-card">
      <div className="sa-toolbar">
        <span className="sa-summary">
          Answers are matched against each category&apos;s description, not its name.
        </span>
        <span className="sa-toolbar-spacer" />
        <label className="sa-field">
          <span>Confidence</span>
          <input
            type="range"
            min="0.1"
            max="0.7"
            step="0.01"
            value={textSettings.threshold}
            onChange={(e) => setTextSetting('threshold', Number(e.target.value))}
          />
          <span className="sa-summary">{textSettings.threshold.toFixed(2)}</span>
        </label>
        <Button variant="ghost" size="sm" icon={Plus} onClick={addBucket}>
          Add a category
        </Button>
      </div>

      <ul className="sa-bucket-list">
        {buckets.map((bucket) => (
          <li key={bucket.id} className="sa-bucket">
            <div className="sa-bucket-head">
              <input
                className="sa-input sa-bucket-name"
                value={bucket.label}
                aria-label="Category name"
                onChange={(e) => updateBucket(bucket.id, { label: e.target.value })}
              />
              <span className="sa-summary">{countOf(bucket.id)} issues</span>
              <Button
                variant="ghost"
                size="sm"
                icon={Trash2}
                aria-label={`Remove ${bucket.label}`}
                onClick={() => removeBucket(bucket.id)}
              />
            </div>
            <textarea
              className="sa-input sa-bucket-description"
              rows={2}
              aria-label={`What belongs in ${bucket.label}`}
              placeholder="Describe in a sentence what belongs here."
              value={bucket.description}
              onChange={(e) => updateBucket(bucket.id, { description: e.target.value })}
            />
          </li>
        ))}
      </ul>
    </Card>
  );
}
