/**
 * Telling the person holding the phone what just happened, without asking them
 * to look at the screen — they are looking at a box.
 *
 * Three sounds, deliberately different from each other rather than three
 * volumes of the same one: accepted is a rising blip, a code already scanned is
 * a low flat one, and nothing-found is silent. The spec calls for a re-read to
 * be ANSWERED (§4.4): silence and "the scanner is broken" are the same
 * experience from behind a camera.
 */

let context = null;

/**
 * Created on the first sound rather than at import: browsers refuse an
 * AudioContext made outside a user gesture, and one made too early sits
 * suspended and plays nothing for the rest of the session.
 */
function audio() {
  if (context) return context;
  const Ctor = window.AudioContext ?? window.webkitAudioContext;
  if (!Ctor) return null;
  context = new Ctor();
  return context;
}

function tone(frequency, durationMs, { volume = 0.09, slideTo } = {}) {
  const ctx = audio();
  if (!ctx) return;
  // A context suspended by autoplay policy resumes on the first gesture, and
  // scanning is always downstream of one.
  if (ctx.state === 'suspended') ctx.resume().catch(() => {});

  const oscillator = ctx.createOscillator();
  const gain = ctx.createGain();
  const now = ctx.currentTime;
  const seconds = durationMs / 1000;

  oscillator.type = 'sine';
  oscillator.frequency.setValueAtTime(frequency, now);
  if (slideTo) oscillator.frequency.linearRampToValueAtTime(slideTo, now + seconds);

  // Ramped rather than switched: a square-edged gain change is an audible
  // click on phone speakers, and thirty of them in a scanning session is
  // genuinely unpleasant.
  gain.gain.setValueAtTime(0, now);
  gain.gain.linearRampToValueAtTime(volume, now + 0.01);
  gain.gain.linearRampToValueAtTime(0, now + seconds);

  oscillator.connect(gain).connect(ctx.destination);
  oscillator.start(now);
  oscillator.stop(now + seconds + 0.02);
}

function buzz(pattern) {
  navigator.vibrate?.(pattern);
}

export function signalAccepted() {
  tone(880, 90, { slideTo: 1320 });
  buzz(30);
}

export function signalDuplicate() {
  tone(300, 150);
  buzz([20, 60, 20]);
}

export function signalDone() {
  tone(660, 120, { slideTo: 990 });
  buzz([25, 40, 25]);
}
