import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';
import { environment } from '../environments/environment';
// @ts-ignore – qz-tray ships no TypeScript declarations
import qz from 'qz-tray';

/**
 * Single source of truth for the QZ Tray digital-signature handshake.
 * Wires qz.security.setCertificatePromise / setSignaturePromise against the
 * backend's GET /qz/certificate and POST /qz/sign endpoints so QZ Tray can
 * verify every request came from this app and skip the "Untrusted Website"
 * popup entirely. All print services should go through this instead of
 * talking to `qz` directly.
 */
@Injectable({ providedIn: 'root' })
export class QzTrayService {
  private securityConfigured = false;

  constructor(private http: HttpClient) {}

  /** Ensures a trusted, connected websocket session. Safe to call before every print. */
  async connect(): Promise<void> {
    this.configureSecurity();

    if (qz.websocket.isActive()) return;

    try {
      await qz.websocket.connect();
    } catch (err: any) {
      // With rejectOnFailure:true, a failed certificate fetch now rejects here
      // with its own message instead of silently connecting as "anonymous" —
      // surface it as-is rather than masking it with a generic message.
      if (err?.message) throw new Error(err.message);
      throw new Error('QZ Tray is not running on this PC. Install and start QZ Tray, then try again.');
    }
  }

  async disconnect(): Promise<void> {
    if (qz.websocket.isActive()) {
      try { await qz.websocket.disconnect(); } catch { /* ignore */ }
    }
  }

  async findPrinters(): Promise<string[]> {
    await this.connect();
    try {
      const found = await qz.printers.find();
      return Array.isArray(found) ? found : (found ? [found] : []);
    } catch {
      throw new Error('QZ Tray could not list printers on this PC.');
    }
  }

  /** Matches a user-facing printer name against what QZ Tray reports as installed. */
  resolvePrinterName(target: string, allPrinters: string[]): string | undefined {
    const wanted = target.trim();
    // CUPS (Linux) queue names can't contain spaces, so "TSC TTP-244 Pro" is
    // installed as "TSC_TTP-244_Pro" — compare with spaces/underscores unified.
    const norm = (s: string) => s.toLowerCase().replace(/[\s_]+/g, ' ').trim();
    return (
      allPrinters.find(p => p === wanted) ??
      allPrinters.find(p => norm(p) === norm(wanted)) ??
      allPrinters.find(p =>
        norm(p).includes(norm(wanted)) ||
        norm(wanted).includes(norm(p))
      )
    );
  }

  /**
   * Connects, resolves the printer name, and sends the print job.
   * `data` is whatever QZ Tray's `qz.print()` expects (raw TSPL, pixel/image, etc).
   */
  async print(printerName: string, data: any[], configOptions: Record<string, any> = {}): Promise<void> {
    await this.connect();

    const allPrinters = await this.findPrinters();
    const resolved = this.resolvePrinterName(printerName, allPrinters);
    if (!resolved) {
      throw new Error(
        `Printer "${printerName}" not found. Installed: ${allPrinters.join(', ') || '(none)'}`
      );
    }

    // On Linux/macOS QZ Tray submits raw jobs through CUPS, which auto-detects
    // TSPL as text/plain and runs it through its text filter — the printer then
    // prints the command text itself instead of executing it. altPrinting makes
    // QZ Tray use `lp -o raw` so the bytes reach the printer untouched.
    const isRawJob = data.some(d => d?.type === 'raw');
    const isWindowsClient = /win/i.test(navigator.platform || navigator.userAgent);
    const options = isRawJob && !isWindowsClient && configOptions['altPrinting'] === undefined
      ? { ...configOptions, altPrinting: true }
      : configOptions;

    const config = qz.configs.create(resolved, options);
    try {
      await qz.print(config, data);
    } catch (err: any) {
      throw new Error(err?.message || 'QZ Tray failed to send the job to the printer.');
    }
  }

  // ── Security handshake ──────────────────────────────────────────────────────

  private configureSecurity(): void {
    if (this.securityConfigured) return;
    this.securityConfigured = true;

    qz.security.setCertificatePromise((resolve: (cert: string) => void, reject: (err: any) => void) => {
      firstValueFrom(
        this.http.get(`${environment.apiUrl}/qz/certificate`, { responseType: 'text' })
      ).then(resolve).catch(err => {
        // qz-tray's default behavior on a rejected cert promise is to SILENTLY
        // connect with certificate:null (shows as "anonymous"/"Invalid
        // Certificate" in QZ Tray, with no visible error in this app) unless
        // {rejectOnFailure: true} is passed below. Logging here is what makes
        // that failure diagnosable instead of silently degrading.
        console.error('[QZ Tray] Failed to fetch certificate from', `${environment.apiUrl}/qz/certificate`, err);
        reject(err);
      });
    }, { rejectOnFailure: true });

    // Backend signs with RSA-SHA512 (crypto.createSign('SHA512')) — QZ Tray
    // must be told to expect SHA512, otherwise signature verification fails
    // silently and QZ Tray falls back to showing the untrusted-site popup.
    qz.security.setSignatureAlgorithm('SHA512');

    qz.security.setSignaturePromise((toSign: string) => {
      return (resolve: (sig: string) => void, reject: (err: any) => void) => {
        firstValueFrom(
          this.http.post<{ signature: string }>(`${environment.apiUrl}/qz/sign`, { request: toSign })
        ).then(r => resolve(r.signature)).catch(err => {
          console.error('[QZ Tray] Failed to sign request via', `${environment.apiUrl}/qz/sign`, err);
          reject(err);
        });
      };
    });

    qz.websocket.setErrorCallbacks((err: any) => console.error('[QZ Tray]', err));
  }
}
