import { useEffect, useState } from 'react';
import QRCode from 'qrcode';
import { Copy, Check } from '../../../../components/ui/Icons';

/**
 * The checklist link a handover or return was made with: the employee scans
 * the code on their own phone right here, or IT copies the link and sends it.
 * Once they sign, the signature lands on the handover rows by itself.
 */
export default function ChecklistLinkCard({ url, mode }) {
  const [qr, setQr] = useState('');
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    let live = true;
    QRCode.toDataURL(url, { margin: 1, width: 200 }).then((data) => { if (live) setQr(data); }).catch(() => {});
    return () => { live = false; };
  }, [url]);

  const copy = async () => {
    try {
      await navigator.clipboard.writeText(url);
      setCopied(true);
      setTimeout(() => setCopied(false), 1800);
    } catch {
      // No clipboard here: the link is on screen to be selected by hand.
    }
  };

  return (
    <div className="till-linkcard">
      <strong>{mode} checklist — waiting for their signature</strong>
      <span>Let them scan this, or copy the link and send it. Their signature is added to this record when they sign.</span>
      {qr && <img src={qr} alt="QR code for the checklist link" width="200" height="200" />}
      <div className="till-linkcard-row">
        <span className="till-mono">{url}</span>
        <button type="button" className="till-chip" onClick={copy}>
          {copied ? <Check size={14} /> : <Copy size={14} />} {copied ? 'Copied' : 'Copy'}
        </button>
      </div>
    </div>
  );
}
