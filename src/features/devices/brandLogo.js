import { hardwareIdsOf } from './derive/deriveHardwareIds.js';

/**
 * Which brand a part or a program belongs to, pure. The logo itself is looked
 * up online (Simple Icons' CDN) by slug when the page draws it; anything the
 * lookup does not have falls back to an initials tile, so a wrong or missing
 * slug costs nothing worse than two letters.
 */

// [pattern, slug, name] — first match wins, so specific names sit above generic ones.
const KNOWN = [
  [/\bintel\b/i, 'intel', 'Intel'],
  [/\bamd\b|ryzen|radeon|athlon|threadripper|advanced micro/i, 'amd', 'AMD'],
  [/nvidia|geforce|quadro|\brtx\b|\bgtx\b/i, 'nvidia', 'NVIDIA'],
  [/qualcomm|snapdragon/i, 'qualcomm', 'Qualcomm'],
  [/\bapple\b/i, 'apple', 'Apple'],
  [/\bwindows\b/i, 'windows', 'Windows'],
  [/microsoft teams/i, 'microsoftteams', 'Microsoft Teams'],
  [/onedrive/i, 'microsoftonedrive', 'OneDrive'],
  [/microsoft edge/i, 'microsoftedge', 'Microsoft Edge'],
  [/\boutlook\b/i, 'microsoftoutlook', 'Outlook'],
  [/\bexcel\b/i, 'microsoftexcel', 'Excel'],
  [/\bword\b/i, 'microsoftword', 'Word'],
  [/powerpoint/i, 'microsoftpowerpoint', 'PowerPoint'],
  [/microsoft 365|office 16|microsoft office|office professional|office standard/i, 'microsoftoffice', 'Microsoft Office'],
  [/visual studio code/i, 'visualstudiocode', 'VS Code'],
  [/visual studio/i, 'visualstudio', 'Visual Studio'],
  [/\.net|dotnet/i, 'dotnet', '.NET'],
  [/sql server/i, 'microsoftsqlserver', 'SQL Server'],
  [/\bmicrosoft\b/i, 'microsoft', 'Microsoft'],
  [/google chrome|\bchrome\b/i, 'googlechrome', 'Google Chrome'],
  [/google drive/i, 'googledrive', 'Google Drive'],
  [/\bgoogle\b/i, 'google', 'Google'],
  [/firefox/i, 'firefoxbrowser', 'Firefox'],
  [/thunderbird/i, 'thunderbird', 'Thunderbird'],
  [/\bopera\b/i, 'opera', 'Opera'],
  [/\bbrave\b/i, 'brave', 'Brave'],
  [/anydesk/i, 'anydesk', 'AnyDesk'],
  [/teamviewer/i, 'teamviewer', 'TeamViewer'],
  [/\bzoom\b/i, 'zoom', 'Zoom'],
  [/whatsapp/i, 'whatsapp', 'WhatsApp'],
  [/telegram/i, 'telegram', 'Telegram'],
  [/\bslack\b/i, 'slack', 'Slack'],
  [/dropbox/i, 'dropbox', 'Dropbox'],
  [/7-zip/i, '7zip', '7-Zip'],
  [/winrar/i, 'winrar', 'WinRAR'],
  [/vlc/i, 'vlcmediaplayer', 'VLC'],
  [/notepad\+\+/i, 'notepadplusplus', 'Notepad++'],
  [/autocad|autodesk|revit|inventor/i, 'autodesk', 'Autodesk'],
  [/solidworks|dassault/i, 'dassaultsystemes', 'Dassault Systèmes'],
  [/sketchup/i, 'sketchup', 'SketchUp'],
  [/foxit/i, 'foxit', 'Foxit'],
  [/\bjava\b|oracle/i, 'openjdk', 'Java'],
  [/python/i, 'python', 'Python'],
  [/node\.?js/i, 'nodedotjs', 'Node.js'],
  [/\bgit\b/i, 'git', 'Git'],
  [/eset/i, 'eset', 'ESET'],
  [/kaspersky/i, 'kaspersky', 'Kaspersky'],
  [/avast/i, 'avast', 'Avast'],
  [/\bavg\b/i, 'avg', 'AVG'],
  [/norton/i, 'norton', 'Norton'],
  [/bitdefender/i, 'bitdefender', 'Bitdefender'],
  [/mcafee/i, 'mcafee', 'McAfee'],
  [/sophos/i, 'sophos', 'Sophos'],
  [/malwarebytes/i, 'malwarebytes', 'Malwarebytes'],
  [/fortinet|forticlient/i, 'fortinet', 'Fortinet'],
  [/cisco|webex/i, 'cisco', 'Cisco'],
  [/realtek/i, 'realtek', 'Realtek'],
  [/logitech|logi\b/i, 'logitech', 'Logitech'],
  [/\bdell\b/i, 'dell', 'Dell'],
  [/hewlett|\bhp\b|\bhpe\b/i, 'hp', 'HP'],
  [/lenovo|thinkpad|thinkcentre/i, 'lenovo', 'Lenovo'],
  [/asus|rog\b/i, 'asus', 'ASUS'],
  [/\bacer\b/i, 'acer', 'Acer'],
  [/\bmsi\b|micro-star/i, 'msi', 'MSI'],
  [/gigabyte|aorus/i, 'gigabyte', 'Gigabyte'],
  [/asrock/i, 'asrock', 'ASRock'],
  [/samsung/i, 'samsung', 'Samsung'],
  [/kingston/i, 'kingstontechnology', 'Kingston'],
  [/western digital|\bwdc\b|\bwd\b/i, 'westerndigital', 'Western Digital'],
  [/seagate/i, 'seagate', 'Seagate'],
  [/crucial|micron/i, 'micron', 'Micron'],
  [/sk ?hynix|\bhynix\b/i, 'skhynix', 'SK hynix'],
  [/toshiba|kioxia/i, 'toshiba', 'Toshiba'],
  [/sandisk/i, 'sandisk', 'SanDisk'],
  [/\blg\b/i, 'lg', 'LG'],
  [/philips/i, 'philips', 'Philips'],
  [/\bbenq\b/i, 'benq', 'BenQ'],
  [/viewsonic/i, 'viewsonic', 'ViewSonic'],
  [/vmware/i, 'vmware', 'VMware'],
  [/virtualbox/i, 'virtualbox', 'VirtualBox'],
  [/steam\b|valve/i, 'steam', 'Steam'],
  [/spotify/i, 'spotify', 'Spotify'],
  [/canon/i, 'canon', 'Canon'],
  [/epson/i, 'epson', 'Epson'],
  [/brother/i, 'brother', 'Brother'],
  [/\bsap\b/i, 'sap', 'SAP'],
  [/adobe|acrobat|photoshop/i, 'adobe', 'Adobe'],
];

/** Memory makers write their JEDEC-ish codes; the common ones, as the scan prints them. */
const MEMORY_CODES = { '80ad': 'SK hynix', '00ce': 'Samsung', 'ce00': 'Samsung', '2c00': 'Micron', '0198': 'Kingston' };

/**
 * The known brand behind any of the given texts (model, maker, publisher),
 * tried in order: { slug, name } or null.
 */
export function brandOf(...texts) {
  for (const raw of texts) {
    const text = String(raw ?? '').trim();
    if (!text) continue;
    const code = MEMORY_CODES[text.toLowerCase().replace(/^0x/, '').slice(0, 4)];
    const probe = code ?? text;
    const hit = KNOWN.find(([pattern]) => pattern.test(probe));
    if (hit) return { slug: hit[1], name: hit[2] };
  }
  return null;
}

/**
 * A program's brand: the known list first, then a GUESS at Simple Icons' slug
 * from the publisher's first word (or the name's), which the CDN either has or
 * answers 404 for — the tile then shows initials.
 */
export function softwareBrandOf(name, publisher) {
  const known = brandOf(name, publisher);
  if (known) return known;
  const word = String(publisher || name || '').trim().split(/[\s,(]+/)[0] ?? '';
  const slug = word.toLowerCase().replace(/[^a-z0-9]/g, '');
  return slug.length >= 3 ? { slug, name: word, guessed: true } : null;
}

/** Up to two letters for the fallback tile. */
export function initialsOfBrand(text) {
  const words = String(text ?? '').replace(/[^\p{L}\p{N}\s]/gu, ' ').trim().split(/\s+/).filter(Boolean);
  if (words.length === 0) return '?';
  if (words.length === 1) return words[0].slice(0, 2).toUpperCase();
  return (words[0][0] + words[1][0]).toUpperCase();
}

/** The logo, in the brand's own colour. */
export function logoUrlOf(slug) {
  return slug ? `https://cdn.simpleicons.org/${encodeURIComponent(slug)}` : null;
}

/**
 * The brand to show beside each Health tile. The tile's value carries the
 * model for CPU and graphics; RAM and storage read the parts behind them.
 */
export function partBrandOf(key, device, value) {
  const { memoryModules, driveSerials } = hardwareIdsOf(device);
  switch (key) {
    case 'cpu': return brandOf(value, device.cpuModel);
    case 'graphics': return brandOf(value);
    case 'windows': return brandOf('Windows');
    case 'ram': return brandOf(...memoryModules.map((m) => m?.maker));
    case 'storage': return brandOf(...driveSerials.map((d) => d?.model), value);
    default: return brandOf(value);
  }
}
