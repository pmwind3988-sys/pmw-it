/**
 * Whether a machine has a graphics card of its own.
 *
 * The scan lists every display adapter Windows knows about, integrated ones
 * included, so "has a GPU" is never the question — the question is whether one
 * of them is a card rather than a corner of the processor.
 */

/** Cards. A GeForce, a Quadro, a Radeon RX or Pro, an Arc. */
const DEDICATED = /geforce|\brtx\b|\bgtx\b|quadro|nvidia|\btesla\b|radeon\s+(?:rx|pro)|firepro|\barc\s+a\d/i;

/** Graphics baked into the processor, or the fallback driver Windows ships. */
const INTEGRATED = /uhd|\bhd graphics|iris|vega\b|radeon\s+graphics|microsoft basic|standard vga/i;

/** `Radeon(TM) Graphics` and `Radeon Graphics` are the same adapter. */
const plain = (name) => name.replace(/\((?:tm|r|c)\)/gi, ' ').replace(/\s+/g, ' ');

export function gpuClass(gpuList = []) {
  const gpus = gpuList.map((name) => String(name).trim()).filter(Boolean);
  if (!gpus.length) return { gpuClass: 'Unknown', dedicatedGpu: null, dedicatedGpuName: null };

  const card = gpus.find((name) => DEDICATED.test(plain(name)));
  if (card) return { gpuClass: 'Dedicated', dedicatedGpu: true, dedicatedGpuName: card };

  if (gpus.some((name) => INTEGRATED.test(plain(name)))) {
    return { gpuClass: 'Integrated', dedicatedGpu: false, dedicatedGpuName: null };
  }

  // An adapter nobody here recognises is not evidence either way.
  return { gpuClass: 'Unknown', dedicatedGpu: null, dedicatedGpuName: null };
}
