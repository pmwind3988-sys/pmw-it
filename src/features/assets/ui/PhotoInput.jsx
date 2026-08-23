import { useRef, useState } from 'react';
import { Camera, X } from '../../../components/ui/Icons';
import { shrinkFile } from '../scan/capturePhoto';
import { savePhoto, deletePhoto } from '../store/assetDb';
import { newId } from '../draft/draftAsset';
import { usePhotoUrl } from './usePhotoUrl';

/**
 * One photograph, on a draft or on the delivery.
 *
 * `capture="environment"` asks a phone to open the back camera straight into
 * the camera app rather than the file browser, which is the difference between
 * two taps and five. On a desktop it is ignored and the file picker opens,
 * which is the right behaviour there.
 */
export default function PhotoInput({ photoId, onChange, label = 'Photo', compact = false }) {
  const inputRef = useRef(null);
  const url = usePhotoUrl(photoId);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');

  const pick = async (event) => {
    const file = event.target.files?.[0];
    // Cleared straight away so choosing the SAME file twice still fires a
    // change event — otherwise retaking a bad photo appears to do nothing.
    event.target.value = '';
    if (!file) return;

    setBusy(true);
    setError('');
    try {
      const blob = await shrinkFile(file);
      const id = newId();
      await savePhoto(id, blob);
      if (photoId) await deletePhoto(photoId).catch(() => {});
      onChange(id);
    } catch {
      setError('That photo could not be saved. There may be no room left.');
    } finally {
      setBusy(false);
    }
  };

  const clear = async () => {
    if (photoId) await deletePhoto(photoId).catch(() => {});
    onChange(null);
  };

  return (
    <div className={`as-photo${compact ? ' as-photo-compact' : ''}`}>
      <input
        ref={inputRef}
        type="file"
        accept="image/*"
        capture="environment"
        onChange={pick}
        hidden
      />

      {url ? (
        <div className="as-photo-shot">
          <img src={url} alt={label} />
          <button type="button" className="as-photo-clear" onClick={clear} aria-label="Remove photo">
            <X size={13} />
          </button>
        </div>
      ) : (
        <button
          type="button"
          className="as-photo-add"
          onClick={() => inputRef.current?.click()}
          disabled={busy}
        >
          <Camera size={compact ? 15 : 18} />
          {!compact && <span>{busy ? 'Saving…' : label}</span>}
        </button>
      )}

      {error && <p className="as-photo-error">{error}</p>}
    </div>
  );
}
