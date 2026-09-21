// src/features/shipment-dashboard/shipment-dashboard.component.ts
//
// Ported from the standalone "BANDB — Shipment Intelligence Dashboard" tool.
// See src/features/legacy-tool-base.ts for the porting pattern this follows.
//
// The original tool is a single huge vanilla-JS HTML file that lets the user
// upload up to 3 courier export files (EKART/BANDB, BlueDart, India Post),
// auto-detects which courier each file belongs to, and then renders three
// large per-courier dashboards plus a combined cross-courier overview. This
// component keeps that logic almost verbatim, translated into methods on
// `this`, with `document.getElementById`/`querySelectorAll` replaced by the
// scoped `this.q()`/`this.qa()` helpers from LegacyToolBase, and every
// `onclick="fn(...)"` attribute in the template replaced by a `data-action`/
// `data-arg` pair handled through a single delegated click listener (see
// `wireEvents()`/`onRootClick()`/`dispatchAction()` below) — this covers both
// the static template buttons and the many buttons this tool generates at
// runtime via `innerHTML` (master nav tabs, TAT tabs, insight action buttons).
import { AfterViewInit, Component, ElementRef, OnDestroy, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import { LegacyToolBase } from '../legacy-tool-base';

declare var Chart: any;

type CourierType = 'bandb' | 'bluedart' | 'indiapost';
type Row = Record<string, any>;

interface CourierSignature {
  label: string;
  badge: string;
  icon: string;
  required?: string[];
  aliasGroups?: string[][];
}

@Component({
  selector: 'app-feature-shipment-dashboard',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './shipment-dashboard.component.html',
})
export class ShipmentDashboardComponent extends LegacyToolBase implements AfterViewInit, OnDestroy {
  @ViewChild('root', { static: true }) rootRef!: ElementRef<HTMLElement>;

  private today = new Date();
  private readonly gridColor = 'rgba(37,42,58,0.8)';
  private readonly tickColor = '#6b7280';

  // ── GLOBAL DATA (BANDB / EKART) ──
  private ALL: Row[] = [];
  private tatRows: Row[] = [];
  private nd5Rows: Row[] = [];
  private dispRows: Row[] = [];
  private delRows: Row[] = [];
  private attRows: Row[] = [];
  private chartMap: Record<string, any> = {};
  private activeTat = 'all';
  private dispDateRows: Row[] = [];
  private delDateRows: Row[] = [];

  // Last month
  private lmRows: Row[] = [];
  private lmYear: number | null = null;
  private lmMonth: number | null = null;
  private readonly STATUS_LIST = [
    'Order Placed', 'Picked Up', 'In Transit', 'Out for Delivery', 'Delivered',
    'Undelivered', 'RTO In Transit', 'RTO Delivered', 'Shipment Delayed', 'Cancelled',
  ];

  // Zone / City
  private znField = 'Zone';
  private znGroups: string[] = [];

  // RTO month-wise
  private rtoMonthKeys: string[] = [];

  // Combined overview charts
  private combinedVolumeChart: any = null;
  private combinedRateChart: any = null;

  // Per-slot upload + build state
  private SLOTS: Record<CourierType, Row[] | null> = { bandb: null, bluedart: null, indiapost: null };
  private LOADED_TYPES: CourierType[] = [];

  private readonly SIGNATURES: Record<CourierType, CourierSignature> = {
    bandb: { label: 'EKART / BANDB', badge: 'BB', icon: '🏭', required: ['Shipment ID', 'Shipment Status', 'Zone', 'Payment Type'] },
    bluedart: { label: 'BlueDart', badge: 'BD', icon: '📦', aliasGroups: [['waybill', 'airway', 'tracking', 'docket'], ['origin'], ['destination'], ['weight'], ['status']] },
    indiapost: { label: 'India Post', badge: 'IP', icon: '📮', aliasGroups: [['article number', 'article no', 'consignment number'], ['booking office'], ['destination office'], ['article type'], ['event description', 'status']] },
  };

  // ── BLUEDART MODULE STATE ──
  private BD_ALL: Row[] = [];
  private BD_ROWS: any[] = [];
  private BD_COLS: Record<string, string | null> = {};
  private bdChartMap: Record<string, any> = {};
  private BD_CACHE: any = {};
  private BD_FILTERED: any[] = [];

  // ── INDIA POST MODULE STATE ──
  private IP_ALL: Row[] = [];
  private IP_ROWS: any[] = [];
  private IP_COLS: Record<string, string | null> = {};
  private ipChartMap: Record<string, any> = {};
  private IP_CACHE: any = {};
  private IP_FILTERED: any[] = [];

  // ═══════════════════════════════════════════════════════
  // LIFECYCLE
  // ═══════════════════════════════════════════════════════

  ngAfterViewInit(): void {
    this.bindRoot(this.rootRef);
    this.q('live-date').textContent = this.today.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
    this.wireEvents();
  }

  ngOnDestroy(): void {
    Object.values(this.chartMap).forEach((c: any) => c?.destroy());
    Object.values(this.bdChartMap).forEach((c: any) => c?.destroy());
    Object.values(this.ipChartMap).forEach((c: any) => c?.destroy());
    this.combinedVolumeChart?.destroy();
    this.combinedRateChart?.destroy();
  }

  // ═══════════════════════════════════════════════════════
  // EVENT WIRING (static filter inputs by id + one delegated
  // click listener that covers every data-action button, static
  // or dynamically generated — replaces the source's onclick="…")
  // ═══════════════════════════════════════════════════════

  private wireEvents(): void {
    (['bandb', 'bluedart', 'indiapost'] as CourierType[]).forEach(type => {
      this.q<HTMLInputElement>('inp-' + type).addEventListener('change', (e) => {
        const file = (e.target as HTMLInputElement).files?.[0];
        void this.handleSlotFile(type, file);
      });
    });

    // BANDB filters
    this.onInput('tat-search', () => this.renderTatTable());
    this.onChange('tat-zone', () => this.renderTatTable());
    this.onChange('tat-pay', () => this.renderTatTable());
    this.onInput('disp-search', () => this.filterDispatch());
    this.onChange('disp-pay', () => this.filterDispatch());
    this.onChange('disp-zone', () => this.filterDispatch());
    this.onInput('del-search', () => this.filterDelivery());
    this.onChange('del-pay', () => this.filterDelivery());
    this.onChange('del-zone', () => this.filterDelivery());
    this.onChange('del-tat', () => this.filterDelivery());
    this.onChange('del-ndr', () => this.filterDelivery());
    this.onInput('pend-search', () => this.filterPending());
    this.onChange('pend-status', () => this.filterPending());
    this.onChange('pend-zone', () => this.filterPending());
    this.onChange('pend-pay', () => this.filterPending());
    this.onInput('att-search', () => this.filterAttempts());
    this.onChange('att-status', () => this.filterAttempts());
    this.onChange('att-zone', () => this.filterAttempts());
    this.onChange('att-pay', () => this.filterAttempts());
    this.onInput('lm-search', () => this.filterLastMonth());
    this.onChange('lm-status', () => this.filterLastMonth());
    this.onChange('lm-pay', () => this.filterLastMonth());
    this.onChange('lm-zone', () => this.filterLastMonth());
    this.onInput('zn-search', () => this.filterZoneOrders());
    this.onChange('zn-zone', () => this.filterZoneOrders());
    this.onChange('zn-status', () => this.filterZoneOrders());
    this.onChange('zn-pay', () => this.filterZoneOrders());
    this.onInput('rto-search', () => this.filterRtoOrders());
    this.onChange('rto-month-sel-filter', () => this.filterRtoOrders());
    this.onChange('rto-status-sel', () => this.filterRtoOrders());
    this.onChange('rto-pay', () => this.filterRtoOrders());
    this.onChange('rto-stuck-sel', () => this.filterRtoOrders());

    // BlueDart filters
    this.onInput('bd-search', () => this.renderBDTable());
    this.onChange('bd-status-filter', () => this.renderBDTable());
    this.onChange('bd-origin-filter', () => this.renderBDTable());
    this.onChange('bd-dest-filter', () => this.renderBDTable());

    // India Post filters
    this.onInput('ip-search', () => this.renderIPTable());
    this.onChange('ip-status-filter', () => this.renderIPTable());
    this.onChange('ip-office-filter', () => this.renderIPTable());

    // Delegated click handling for every data-action button (static + generated)
    this.hostEl.addEventListener('click', (e) => this.onRootClick(e));
  }

  private onInput(id: string, fn: () => void): void {
    this.q<HTMLInputElement>(id).addEventListener('input', fn);
  }
  private onChange(id: string, fn: () => void): void {
    this.q<HTMLSelectElement>(id).addEventListener('change', fn);
  }

  private onRootClick(e: MouseEvent): void {
    const target = e.target as HTMLElement | null;
    const actionEl = target?.closest<HTMLElement>('[data-action]');
    if (!actionEl) return;
    const action = actionEl.getAttribute('data-action') || '';
    const arg = actionEl.getAttribute('data-arg') || undefined;
    this.dispatchAction(action, arg, actionEl);
  }

  private dispatchAction(action: string, arg: string | undefined, el: HTMLElement): void {
    switch (action) {
      case 'resetDashboard': this.resetDashboard(); break;
      case 'browseSlot': this.q<HTMLInputElement>('inp-' + arg).click(); break;
      case 'buildDashboard': this.buildDashboard(); break;
      case 'masterSwitch': this.masterSwitch(arg || ''); break;
      case 'switchTab': this.switchTab(arg || ''); break;
      case 'setTatTab': this.setTatTab(arg || 'all'); break;
      case 'downloadTatAll': this.downloadTatAll(); break;
      case 'downloadTatFiltered': this.downloadTatFiltered(); break;
      case 'loadDispatchDate': this.loadDispatchDate(); break;
      case 'downloadDispatch': this.downloadDispatch(); break;
      case 'loadDeliveryDate': this.loadDeliveryDate(); break;
      case 'downloadDelivery': this.downloadDelivery(); break;
      case 'downloadPending': this.downloadPending(); break;
      case 'downloadAddressIssue': this.downloadAddressIssue(); break;
      case 'loadLastMonthSelected': this.loadLastMonthSelected(); break;
      case 'downloadLastMonth': this.downloadLastMonth(); break;
      case 'downloadLastMonthSummary': this.downloadLastMonthSummary(); break;
      case 'downloadZoneSummary': this.downloadZoneSummary(); break;
      case 'downloadZoneOrders': this.downloadZoneOrders(); break;
      case 'downloadRtoMonthly': this.downloadRtoMonthly(); break;
      case 'downloadRtoOrders': this.downloadRtoOrders(); break;
      case 'downloadAttempts': this.downloadAttempts(); break;
      case 'switchTabBD': this.switchTabBD(arg || 'overview', el); break;
      case 'downloadBD': this.downloadBD(); break;
      case 'switchTabIP': this.switchTabIP(arg || 'overview', el); break;
      case 'downloadIP': this.downloadIP(); break;
      default: break;
    }
  }

  // ═══════════════════════════════════════════════════════
  // SHARED UTILITIES (dates, formatting, badges)
  // ═══════════════════════════════════════════════════════

  private parseDate(ds: any): Date | null {
    if (!ds || ds === 'Not Picked' || ds === '--') return null;
    const d = new Date(String(ds).replace(/(\d{2}) (\w{3}) (\d{4})/, '$2 $1 $3'));
    if (isNaN(d.getTime())) return null;
    return d;
  }
  private daysBetween(ds: any): number | null {
    const d = this.parseDate(ds);
    if (!d) return null;
    return Math.floor((this.today.getTime() - d.getTime()) / 86400000);
  }
  private daysTaken(pickupStr: any, deliveryStr: any): number | null {
    const p = this.parseDate(pickupStr), dd = this.parseDate(deliveryStr);
    if (!p || !dd) return null;
    return Math.round((dd.getTime() - p.getTime()) / 86400000);
  }
  /** Shared numeric key for the "dd Mon yyyy"-style date-string sorts used throughout the source. */
  private dateSortKey(s: string): number {
    const d = new Date(String(s).replace(/(\d{2}) (\w{3}) (\d{4})/, '$2 $1 $3'));
    return d.getTime();
  }
  private fc(v: any): string { return '₹' + (parseFloat(v) || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 }); }
  private fk(v: any): string { const n = parseFloat(v) || 0; return n >= 1000 ? '₹' + (n / 1000).toFixed(1) + 'K' : this.fc(n); }
  private sBadge(s: any): string {
    const m: Record<string, string> = {
      'Delivered': 'b-del', 'In Transit': 'b-it', 'Out for Delivery': 'b-ofd',
      'Undelivered': 'b-ndr', 'RTO In Transit': 'b-rto', 'RTO Delivered': 'b-rtodel', 'Cancelled': 'b-can',
      'Order Placed': 'b-op', 'Picked Up': 'b-pk', 'Shipment Delayed': 'b-sd',
    };
    return `<span class="badge ${m[s] || 'b-it'}">${s || '—'}</span>`;
  }
  private pBadge(p: any): string { return `<span class="badge ${p === 'COD' ? 'b-cod' : 'b-pre'}">${p || '—'}</span>`; }
  private tagChips(t: any): string {
    if (!t || !String(t).trim() || t === '=""') return '';
    return String(t).split(',').map(x => {
      x = x.trim(); let c = 't-lbl';
      if (x.includes('Address')) c = 't-addr';
      else if (x.includes('Escalation')) c = 't-esc';
      else if (x.includes('Feedback')) c = 't-fb';
      return `<span class="tag-chip ${c}">${x}</span>`;
    }).join('');
  }
  private daysCell(days: number | null): string {
    if (days === null) return '—';
    const c = days >= 10 ? 'var(--red)' : days >= 7 ? 'var(--orange)' : 'var(--accent)';
    return `<span style="font-weight:900;color:${c}">${days}d</span>`;
  }
  private daysTakenCell(days: number | null): string {
    if (days === null) return '—';
    let c = 'var(--green)';
    if (days >= 3 && days <= 4) c = 'var(--orange)';
    if (days >= 5) c = 'var(--red)';
    return `<span style="font-weight:900;color:${c}">${days} day${days === 1 ? '' : 's'}</span>`;
  }
  private destroyChart(id: string): void {
    if (this.chartMap[id]) { this.chartMap[id].destroy(); delete this.chartMap[id]; }
  }

  // ═══════════════════════════════════════════════════════
  // COURIER SIGNATURES (used to sanity-check each slot's file)
  // ═══════════════════════════════════════════════════════

  private normH(h: any): string {
    return String(h || '').toLowerCase().trim().replace(/[_\-\/]+/g, ' ').replace(/\s+/g, ' ').trim();
  }
  private headerMatchesAlias(headers: string[], aliases: string[]): boolean {
    return headers.some(h => { const nh = this.normH(h); return aliases.some(a => nh.includes(this.normH(a))); });
  }
  private detectCourierType(headers: string[]): CourierType | 'unknown' {
    const norm = headers.map(h => this.normH(h));
    const bandbHit = this.SIGNATURES.bandb.required!.every(req => norm.includes(this.normH(req)));
    if (bandbHit) return 'bandb';
    let bdScore = 0, ipScore = 0;
    this.SIGNATURES.bluedart.aliasGroups!.forEach(g => { if (this.headerMatchesAlias(headers, g)) bdScore++; });
    this.SIGNATURES.indiapost.aliasGroups!.forEach(g => { if (this.headerMatchesAlias(headers, g)) ipScore++; });
    if (this.headerMatchesAlias(headers, ['article number', 'booking office', 'destination office'])) ipScore += 2;
    if (this.headerMatchesAlias(headers, ['waybill', 'airway', 'docket'])) bdScore += 2;
    if (bdScore === 0 && ipScore === 0) return 'unknown';
    return bdScore >= ipScore ? 'bluedart' : 'indiapost';
  }

  // ═══════════════════════════════════════════════════════
  // PER-SLOT UPLOAD (BANDB / BlueDart / India Post)
  //
  // The source's parseAnyFile(file, cb) callback helper is ported as a single
  // async method: the .xlsx/.xls branch keeps using XLSX.read/sheet_to_json,
  // the CSV branch is rewritten on top of XLSX (string mode) instead of
  // PapaParse, since this app already depends on `xlsx` everywhere else.
  // ═══════════════════════════════════════════════════════

  private async parseAnyFile(file: File): Promise<Row[]> {
    const name = file.name.toLowerCase();
    if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(new Uint8Array(buf), { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      return XLSX.utils.sheet_to_json<Row>(ws, { defval: '' });
    }
    return this.parseCsvFile(file);
  }

  private async parseCsvFile(file: File): Promise<Row[]> {
    const text = await file.text();
    const wb = XLSX.read(text, { type: 'string' });
    const rows = XLSX.utils.sheet_to_json<Row>(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    return rows.filter(r => Object.values(r).some(v => String(v).trim() !== ''));
  }

  private async handleSlotFile(expectedType: CourierType, file: File | undefined): Promise<void> {
    if (!file) return;
    this.q('detect-error').style.display = 'none';
    const zone = this.q('slot-' + expectedType);
    const status = this.q('status-' + expectedType);
    zone.classList.remove('loaded', 'error');
    status.textContent = `⏳ Reading ${file.name}...`;
    const rows = await this.parseAnyFile(file);
    if (!rows || !rows.length) {
      zone.classList.add('error');
      status.textContent = '⚠️ No rows found in this file.';
      return;
    }
    const headers = Object.keys(rows[0]);
    const detected = this.detectCourierType(headers);
    if (detected !== 'unknown' && detected !== expectedType) {
      zone.classList.add('error');
      status.textContent = `⚠️ This looks like a ${this.SIGNATURES[detected].label} file, not ${this.SIGNATURES[expectedType].label}. Uploaded anyway — check the slot.`;
      this.SLOTS[expectedType] = rows;
    } else {
      zone.classList.add('loaded');
      status.textContent = `✅ ${file.name} — ${rows.length.toLocaleString()} rows`;
      this.SLOTS[expectedType] = rows;
    }
    this.updateBuildButton();
  }

  private updateBuildButton(): void {
    const n = Object.values(this.SLOTS).filter(Boolean).length;
    const btn = this.q<HTMLButtonElement>('build-btn');
    if (n > 0) { btn.style.opacity = '1'; btn.style.pointerEvents = 'auto'; btn.textContent = `📊 Build Dashboard (${n} file${n > 1 ? 's' : ''})`; }
    else { btn.style.opacity = '.4'; btn.style.pointerEvents = 'none'; btn.textContent = '📊 Build Dashboard'; }
  }

  // ═══════════════════════════════════════════════════════
  // BUILD DASHBOARD (routes loaded slots into their modules)
  // ═══════════════════════════════════════════════════════

  private buildDashboard(): void {
    this.LOADED_TYPES = (Object.keys(this.SLOTS) as CourierType[]).filter(t => this.SLOTS[t]);
    if (!this.LOADED_TYPES.length) return;
    this.q('upload-section').style.display = 'none';
    this.q('reset-btn').style.display = 'inline-block';
    this.q('hdr-title').textContent = this.LOADED_TYPES.length > 1 ? 'Multi-Courier Shipments' : this.SIGNATURES[this.LOADED_TYPES[0]].label;
    this.q('hdr-sub').textContent = 'Shipment Intelligence Dashboard';
    this.q('hdr-badge').textContent = this.LOADED_TYPES.length > 1 ? 'MC' : this.SIGNATURES[this.LOADED_TYPES[0]].badge;

    if (this.SLOTS.bandb) { this.ALL = this.SLOTS.bandb; this.processAll(); }
    if (this.SLOTS.bluedart) { this.BD_ALL = this.SLOTS.bluedart; this.processBluedart(); }
    if (this.SLOTS.indiapost) { this.IP_ALL = this.SLOTS.indiapost; this.processIndiaPost(); }

    const nav = this.q('master-nav');
    const tabs: { key: string; label: string }[] = [];
    if (this.LOADED_TYPES.length > 1) tabs.push({ key: 'combined', label: '🔗 Combined Overview' });
    if (this.SLOTS.bandb) tabs.push({ key: 'bandb', label: '🏭 EKART / BANDB' });
    if (this.SLOTS.bluedart) tabs.push({ key: 'bluedart', label: '📦 BlueDart' });
    if (this.SLOTS.indiapost) tabs.push({ key: 'indiapost', label: '📮 India Post' });
    nav.innerHTML = tabs.map((t, i) => `<button class="master-tab${i === 0 ? ' active' : ''}" data-key="${t.key}" data-action="masterSwitch" data-arg="${t.key}">${t.label}</button>`).join('');
    nav.style.display = 'flex';

    if (this.LOADED_TYPES.length > 1) this.renderCombined();
    this.masterSwitch(tabs[0].key);
  }

  private masterSwitch(key: string): void {
    this.qa('.master-tab').forEach(b => b.classList.toggle('active', b.getAttribute('data-key') === key));
    this.q('app-combined').style.display = key === 'combined' ? 'block' : 'none';
    this.q('app-bandb').style.display = key === 'bandb' ? 'block' : 'none';
    this.q('app-bluedart').style.display = key === 'bluedart' ? 'block' : 'none';
    this.q('app-indiapost').style.display = key === 'indiapost' ? 'block' : 'none';
  }

  private courierStats(type: CourierType): { n: number; del: number; rate: number; pending: number; rto: number } {
    if (type === 'bandb') {
      const n = this.ALL.length, del = this.ALL.filter(r => r['Shipment Status'] === 'Delivered').length;
      const rto = this.ALL.filter(r => /RTO/i.test(r['Shipment Status'] || '')).length;
      return { n, del, rate: n ? del / n * 100 : 0, pending: n - del - rto, rto };
    }
    if (type === 'bluedart') {
      const n = this.BD_ROWS.length, del = this.BD_ROWS.filter(r => /delivered/i.test(r.status)).length;
      const rto = this.BD_ROWS.filter(r => /rto|return/i.test(r.status)).length;
      return { n, del, rate: n ? del / n * 100 : 0, pending: n - del - rto, rto };
    }
    const n = this.IP_ROWS.length, del = this.IP_ROWS.filter(r => /delivered/i.test(r.status)).length;
    return { n, del, rate: n ? del / n * 100 : 0, pending: n - del, rto: 0 };
  }

  private renderCombined(): void {
    this.q('combined-date').textContent = 'Report generated ' + this.today.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
    const stats = this.LOADED_TYPES.map(t => ({ type: t, label: this.SIGNATURES[t].label, ...this.courierStats(t) }));
    const totalN = stats.reduce((s, x) => s + x.n, 0);
    const totalDel = stats.reduce((s, x) => s + x.del, 0);
    const blendedRate = totalN ? ((totalDel / totalN) * 100).toFixed(1) : '0.0';
    const kpis: any[] = [
      { label: 'Total Shipments (All Couriers)', val: totalN.toLocaleString(), icon: '📦', ka: 'var(--accent)' },
      { label: 'Total Delivered', val: totalDel.toLocaleString(), sub: blendedRate + '% blended rate', icon: '✅', ka: 'var(--green)' },
      { label: 'Couriers Loaded', val: String(stats.length), icon: '🔗', ka: 'var(--accent2)' },
      { label: 'Total Pending / RTO', val: stats.reduce((s, x) => s + x.pending + x.rto, 0).toLocaleString(), icon: '⚠️', ka: 'var(--orange)' },
    ];
    this.q('combined-kpi-grid').innerHTML = kpis.map(k => `<div class="kpi" style="--ka:${k.ka}">
      <div class="kpi-icon">${k.icon}</div><div class="kpi-label">${k.label}</div>
      <div class="kpi-val">${k.val}</div>${k.sub ? `<div class="kpi-sub">${k.sub}</div>` : ''}
    </div>`).join('');

    const cfg = { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#9ca3af', font: { size: 11 } } } } };
    this.combinedVolumeChart?.destroy();
    this.combinedRateChart?.destroy();
    this.combinedVolumeChart = new Chart(this.q('combined-volume-chart'), { type: 'bar', data: { labels: stats.map(s => s.label), datasets: [{ label: 'Total Shipments', data: stats.map(s => s.n), backgroundColor: ['#4f8ef7', '#7c5cfc', '#22c55e'], borderRadius: 8 }] }, options: { ...cfg, scales: { y: { beginAtZero: true, ticks: { color: '#9ca3af' } }, x: { ticks: { color: '#9ca3af' } } } } });
    this.combinedRateChart = new Chart(this.q('combined-rate-chart'), { type: 'bar', data: { labels: stats.map(s => s.label), datasets: [{ label: 'Delivery Rate %', data: stats.map(s => Number(s.rate.toFixed(1))), backgroundColor: ['#22c55e', '#f59e0b', '#06b6d4'], borderRadius: 8 }] }, options: { ...cfg, scales: { y: { beginAtZero: true, max: 100, ticks: { color: '#9ca3af' } }, x: { ticks: { color: '#9ca3af' } } } } });

    this.q('combined-tbody').innerHTML = stats.map(s => `<tr>
      <td><strong>${this.SIGNATURES[s.type].icon} ${s.label}</strong></td>
      <td>${s.n.toLocaleString()}</td><td>${s.del.toLocaleString()}</td>
      <td>${s.rate.toFixed(1)}%</td><td>${s.pending.toLocaleString()}</td><td>${s.rto.toLocaleString()}</td>
    </tr>`).join('');
  }

  private resetDashboard(): void {
    location.reload();
  }

  // ═══════════════════════════════════════════════════════
  // BANDB MODULE — PROCESS ALL DATA
  // ═══════════════════════════════════════════════════════

  private switchTab(name: string): void {
    const tabs = ['overview', 'tat', 'dispatch', 'delivered', 'pending', 'attempts', 'lastmonth', 'zone', 'rto', 'thoughts'];
    this.qa<HTMLElement>('#app-bandb .tab-btn').forEach((b, i) => b.classList.toggle('active', tabs[i] === name));
    this.qa<HTMLElement>('#app-bandb .page').forEach(p => p.classList.remove('active'));
    this.q('page-' + name).classList.add('active');
  }

  private processAll(): void {
    const df = this.ALL, n = df.length;
    const notDelSt = new Set(['In Transit', 'Picked Up', 'Order Placed', 'Out for Delivery', 'Undelivered', 'RTO In Transit', 'Shipment Delayed']);

    this.tatRows = df.filter(r => r['TAT Breach'] && String(r['TAT Breach']).trim());

    this.nd5Rows = df.filter(r => {
      if (!notDelSt.has(r['Shipment Status'])) return false;
      const d = this.daysBetween(r['Pickup Date']); return d !== null && d >= 5;
    });

    this.attRows = df.filter(r => {
      if (r['Shipment Status'] === 'Delivered') return false;
      const att = parseInt(r['Attempts']) || 0;
      return att >= 2;
    });

    this.dispRows = df.filter(r => r['Pickup Date'] && r['Pickup Date'] !== 'Not Picked' && r['Pickup Date'] !== '--');
    this.delRows = df.filter(r => r['Delivery Date'] && r['Delivery Date'] !== '--' && String(r['Delivery Date']).trim());

    const sm: Record<string, number> = {};
    df.forEach(r => { const s = r['Shipment Status'] || 'Unknown'; sm[s] = (sm[s] || 0) + 1; });

    let tot = 0, cod = 0;
    df.forEach(r => { tot += parseFloat(r['Total Amount']) || 0; cod += parseFloat(r['COD Amount']) || 0; });

    const del = sm['Delivered'] || 0;
    const tatBr = this.tatRows.length;

    this.q('rpt-date').textContent = `Report for ${this.today.toLocaleDateString('en-IN', { day: 'numeric', month: 'long', year: 'numeric' })} · ${n.toLocaleString()} shipments`;

    this.renderKPIs({
      n, del, inT: sm['In Transit'] || 0, ofd: sm['Out for Delivery'] || 0,
      ndr: sm['Undelivered'] || 0, rto: sm['RTO In Transit'] || 0,
      tot, cod, tatBr, nd5: this.nd5Rows.length, cancelled: sm['Cancelled'] || 0,
    });

    this.renderOverviewCharts(df, sm);
    this.setupTatPage();
    this.setupDispatchPage();
    this.setupDeliveryPage();
    this.setupPendingPage();
    this.setupAttemptsPage();
    this.setupLastMonthPage();
    this.setupZonePage();
    this.setupRtoPage();
    this.renderInsights(df, sm, tot, cod, tatBr, this.nd5Rows.length, n);
  }

  // ── KPIs ──
  private renderKPIs(p: { n: number; del: number; inT: number; ofd: number; ndr: number; rto: number; tot: number; cod: number; tatBr: number; nd5: number; cancelled: number }): void {
    const { n, del, inT, ofd, ndr, rto, tot, cod, tatBr, nd5 } = p;
    const rate = ((del / n) * 100).toFixed(1);
    const rtoRate = n ? ((rto / n) * 100).toFixed(2) : '0.00';
    const kpis: any[] = [
      { l: 'Total Shipments', v: n.toLocaleString(), s: 'All records', i: '📦', a: 'linear-gradient(90deg,#4f8ef7,#7c5cfc)' },
      { l: 'Delivered', v: del.toLocaleString(), s: `${rate}% delivery rate`, i: '✅', a: 'linear-gradient(90deg,#22c55e,#16a34a)' },
      { l: 'In Transit', v: inT.toLocaleString(), s: 'Moving to destination', i: '🚚', a: 'linear-gradient(90deg,#4f8ef7,#06b6d4)' },
      { l: 'Out for Delivery', v: ofd.toLocaleString(), s: 'Last-mile today', i: '🏃', a: 'linear-gradient(90deg,#7c5cfc,#a855f7)' },
      { l: 'Undelivered (NDR)', v: ndr.toLocaleString(), s: 'Needs reattempt', i: '⚠️', a: 'linear-gradient(90deg,#f59e0b,#d97706)' },
      { l: '⚡ Pending 5+ Days', v: nd5.toLocaleString(), s: 'Critical — take action', i: '🔴', a: 'linear-gradient(90deg,#ef4444,#dc2626)' },
      { l: 'TAT Breached', v: tatBr.toLocaleString(), s: 'Delayed vs SLA', i: '⏱', a: 'linear-gradient(90deg,#f59e0b,#ef4444)' },
      { l: 'RTO In Transit', v: rto.toLocaleString(), s: `${rtoRate}% of total shipments`, i: '↩️', a: 'linear-gradient(90deg,#ef4444,#7c5cfc)' },
      { l: 'Total Revenue', v: this.fk(tot), s: this.fc(tot), i: '💰', a: 'linear-gradient(90deg,#22c55e,#06b6d4)' },
      { l: 'COD Exposure', v: this.fk(cod), s: 'Cash on delivery', i: '💵', a: 'linear-gradient(90deg,#f59e0b,#22c55e)' },
    ];
    this.q('kpi-grid').innerHTML = kpis.map(k => `
      <div class="kpi" style="--ka:${k.a}">
        <div class="kpi-icon">${k.i}</div>
        <div class="kpi-label">${k.l}</div>
        <div class="kpi-val">${k.v}</div>
        <div class="kpi-sub">${k.s}</div>
      </div>`).join('');
  }

  // ── OVERVIEW CHARTS ──
  private renderOverviewCharts(df: Row[], sm: Record<string, number>): void {
    // Trend
    const tm: Record<string, number> = {};
    df.forEach(r => { const d = r['Pickup Date']; if (d && d !== 'Not Picked' && d !== '--') tm[d] = (tm[d] || 0) + 1; });
    const tl = Object.keys(tm).sort((a, b) => this.dateSortKey(a) - this.dateSortKey(b));
    this.destroyChart('trend');
    this.chartMap['trend'] = new Chart(this.q('trend-chart'), { type: 'line', data: { labels: tl.map(l => l.replace(' 2026', '')), datasets: [{ label: 'Dispatched', data: tl.map(l => tm[l]), borderColor: '#4f8ef7', backgroundColor: 'rgba(79,142,247,.1)', fill: true, tension: .4, pointRadius: 4, pointBackgroundColor: '#4f8ef7' }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor, font: { size: 10 } }, grid: { color: this.gridColor } }, y: { ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    // Delivery trend
    const dm: Record<string, number> = {};
    df.forEach(r => { const d = r['Delivery Date']; if (d && d !== '--' && String(d).trim()) dm[d] = (dm[d] || 0) + 1; });
    const dl = Object.keys(dm).sort((a, b) => this.dateSortKey(a) - this.dateSortKey(b));
    this.destroyChart('del-trend');
    this.chartMap['del-trend'] = new Chart(this.q('del-trend-chart'), { type: 'bar', data: { labels: dl.map(l => l.replace(' 2026', '')), datasets: [{ label: 'Delivered', data: dl.map(l => dm[l]), backgroundColor: 'rgba(34,197,94,.8)', borderRadius: 5 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    // Payment
    const pm: Record<string, number> = {};
    df.forEach(r => { const p = r['Payment Type'] || 'Unknown'; pm[p] = (pm[p] || 0) + 1; });
    this.destroyChart('pay');
    this.chartMap['pay'] = new Chart(this.q('pay-chart'), { type: 'doughnut', data: { labels: Object.keys(pm), datasets: [{ data: Object.values(pm), backgroundColor: ['#4f8ef7', '#f59e0b'], borderWidth: 0, hoverOffset: 8 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 11 }, padding: 12 } } } } });

    // Zone
    const zm: Record<string, number> = {};
    df.forEach(r => { const z = r['Zone'] || '?'; zm[z] = (zm[z] || 0) + 1; });
    this.destroyChart('zone');
    this.chartMap['zone'] = new Chart(this.q('zone-chart'), { type: 'doughnut', data: { labels: Object.keys(zm).map(z => `Zone ${z}`), datasets: [{ data: Object.values(zm), backgroundColor: ['#7c5cfc', '#4f8ef7', '#22c55e', '#f59e0b'], borderWidth: 0, hoverOffset: 8 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 11 }, padding: 8 } } } } });

    // Status
    const sc: Record<string, string> = { 'Delivered': '#22c55e', 'In Transit': '#4f8ef7', 'Picked Up': '#a855f7', 'Order Placed': '#06b6d4', 'Out for Delivery': '#7c5cfc', 'Undelivered': '#f59e0b', 'RTO In Transit': '#ef4444', 'RTO Delivered': '#ec4899', 'Cancelled': '#6b7280', 'Shipment Delayed': '#dc2626' };
    this.destroyChart('status');
    this.chartMap['status'] = new Chart(this.q('status-chart'), { type: 'bar', data: { labels: Object.keys(sm), datasets: [{ data: Object.values(sm), backgroundColor: Object.keys(sm).map(s => sc[s] || '#4f8ef7'), borderRadius: 6, borderSkipped: false }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor, font: { size: 9 } }, grid: { display: false } }, y: { ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    // NDR
    const nm: Record<string, number> = {};
    df.forEach(r => { if (r['Latest NDR Status'] && String(r['Latest NDR Status']).trim()) { const nn = r['Latest NDR Status']; nm[nn] = (nm[nn] || 0) + 1; } });
    this.destroyChart('ndr');
    this.chartMap['ndr'] = new Chart(this.q('ndr-chart'), { type: 'doughnut', data: { labels: Object.keys(nm), datasets: [{ data: Object.values(nm), backgroundColor: ['#f59e0b', '#ef4444', '#4f8ef7', '#7c5cfc', '#22c55e'], borderWidth: 0, hoverOffset: 8 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 10 }, padding: 8 } } } } });

    // Tags
    const tgm: Record<string, number> = {};
    df.forEach(r => { if (r['Tags'] && String(r['Tags']).trim() && r['Tags'] !== '=""') { String(r['Tags']).split(',').forEach((t: string) => { t = t.trim(); if (t) tgm[t] = (tgm[t] || 0) + 1; }); } });
    this.destroyChart('tag');
    this.chartMap['tag'] = new Chart(this.q('tag-chart'), { type: 'bar', data: { labels: Object.keys(tgm), datasets: [{ data: Object.values(tgm), backgroundColor: ['#ef4444', '#f59e0b', '#06b6d4', '#22c55e', '#7c5cfc'], borderRadius: 5 }] }, options: { responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor }, grid: { color: this.gridColor } }, y: { ticks: { color: '#9ca3af', font: { size: 10 } }, grid: { display: false } } } } });

    // TAT overview
    const tatm: Record<string, number> = {};
    this.tatRows.forEach(r => { const t = r['TAT Breach']; tatm[t] = (tatm[t] || 0) + 1; });
    const tatSorted = Object.entries(tatm).sort((a, b) => {
      const getN = (s: string) => parseInt((s.match(/\d+/) || ['0'])[0]);
      return getN(a[0]) - getN(b[0]);
    });
    this.destroyChart('tat-ov');
    this.chartMap['tat-ov'] = new Chart(this.q('tat-overview-chart'), { type: 'bar', data: { labels: tatSorted.map(x => x[0]), datasets: [{ data: tatSorted.map(x => x[1]), backgroundColor: ['#22c55e', '#f59e0b', '#ef4444', '#dc2626', '#991b1b', '#7f1d1d'].slice(0, tatSorted.length), borderRadius: 6 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });
  }

  // ═══════════════════════════════════════════════════════
  // TAT BREACH PAGE
  // ═══════════════════════════════════════════════════════

  private setupTatPage(): void {
    const tatSet: Record<string, number> = {};
    this.tatRows.forEach(r => { const t = r['TAT Breach']; tatSet[t] = (tatSet[t] || 0) + 1; });
    const sorted = Object.entries(tatSet).sort((a, b) => parseInt((a[0].match(/\d+/) || ['0'])[0]) - parseInt((b[0].match(/\d+/) || ['0'])[0]));

    const barItems: any[] = [
      { l: 'Total TAT Breached', v: this.tatRows.length, c: 'var(--red)' },
      ...sorted.map(([k, v]) => ({ l: k, v, c: v >= 10 ? 'var(--orange)' : 'var(--muted2)' })),
    ];
    this.q('tat-stat-bar').innerHTML = barItems.map(x => `<div class="stat-item"><div class="stat-val" style="color:${x.c}">${x.v}</div><div class="stat-lbl">${x.l}</div></div>`).join('');

    this.q('tat-tabs').innerHTML = `<button class="tat-tab active" data-tat="all" data-action="setTatTab" data-arg="all">All Breached <span class="cnt">${this.tatRows.length}</span></button>`
      + sorted.map(([k, v]) => `<button class="tat-tab" data-tat="${k}" data-action="setTatTab" data-arg="${k}">${k} <span class="cnt">${v}</span></button>`).join('');

    const zones = [...new Set(this.tatRows.map(r => r['Zone']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('tat-zone').innerHTML = '<option value="">All Zones</option>' + zones.map(z => `<option value="${z}">Zone ${z}</option>`).join('');

    this.activeTat = 'all';
    this.renderTatTable();
  }

  private setTatTab(val: string): void {
    this.activeTat = val;
    this.qa('.tat-tab').forEach(b => b.classList.toggle('active', b.getAttribute('data-tat') === val));
    this.renderTatTable();
  }

  private renderTatTable(): void {
    const search = this.q<HTMLInputElement>('tat-search').value.toLowerCase();
    const zone = this.q<HTMLSelectElement>('tat-zone').value;
    const pay = this.q<HTMLSelectElement>('tat-pay').value;

    const rows = this.tatRows.filter(r => {
      if (this.activeTat !== 'all' && r['TAT Breach'] !== this.activeTat) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });

    this.q('tat-count').textContent = `${rows.length} shipments`;
    this.q('tat-tbody').innerHTML = rows.map((r, i) => `<tr>
      <td style="color:var(--muted)">${i + 1}</td>
      <td style="font-family:monospace;font-size:11px;color:var(--accent)">${r['Shipment ID'] || '—'}</td>
      <td style="font-weight:600">${r['Customer'] || '—'}</td>
      <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
      <td>${this.pBadge(r['Payment Type'])}</td>
      <td>${this.sBadge(r['Shipment Status'])}</td>
      <td style="color:var(--muted);font-size:11px">${r['Pickup Date'] || '—'}</td>
      <td><span style="font-weight:800;color:var(--red);font-size:11px">${r['TAT Breach'] || '—'}</span></td>
      <td style="font-weight:700;color:var(--accent2)">Zone ${r['Zone'] || '?'}</td>
      <td>${this.fc(r['Total Amount'])}</td>
      <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
      <td style="text-align:center">${r['Attempts'] || '—'}</td>
      <td style="font-size:11px;color:var(--orange)">${r['Latest NDR Status'] || '—'}</td>
      <td style="font-size:11px;color:var(--muted);max-width:160px;overflow:hidden;text-overflow:ellipsis" title="${r['Latest Tracking Description'] || ''}">${String(r['Latest Tracking Description'] || '—').substring(0, 30)}${String(r['Latest Tracking Description'] || '').length > 30 ? '…' : ''}</td>
      <td>${this.tagChips(r['Tags'])}</td>
    </tr>`).join('');
  }

  private getCurrentTatRows(): Row[] {
    const search = this.q<HTMLInputElement>('tat-search').value.toLowerCase();
    const zone = this.q<HTMLSelectElement>('tat-zone').value;
    const pay = this.q<HTMLSelectElement>('tat-pay').value;
    return this.tatRows.filter(r => {
      if (this.activeTat !== 'all' && r['TAT Breach'] !== this.activeTat) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
  }

  private downloadTatAll(): void { this.excelDownload(this.tatRows, 'BANDB_TAT_Breach_All'); }
  private downloadTatFiltered(): void { this.excelDownload(this.getCurrentTatRows(), `BANDB_TAT_Breach_${this.activeTat.replace(/ /g, '_')}`); }

  // ═══════════════════════════════════════════════════════
  // DISPATCH BY DATE
  // ═══════════════════════════════════════════════════════

  private setupDispatchPage(): void {
    const dates = [...new Set(this.dispRows.map(r => r['Pickup Date']))].sort((a, b) => this.dateSortKey(b) - this.dateSortKey(a));
    this.q<HTMLSelectElement>('dispatch-date-sel').innerHTML = dates.map(d => `<option value="${d}">${d} (${this.dispRows.filter(r => r['Pickup Date'] === d).length} orders)</option>`).join('');
    const zones = [...new Set(this.dispRows.map(r => r['Zone']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('disp-zone').innerHTML = '<option value="">All Zones</option>' + zones.map(z => `<option value="${z}">Zone ${z}</option>`).join('');
    this.loadDispatchDate();
  }

  private loadDispatchDate(): void {
    const date = this.q<HTMLSelectElement>('dispatch-date-sel').value;
    this.dispDateRows = this.dispRows.filter(r => r['Pickup Date'] === date);
    const tot = this.dispDateRows.reduce((s, r) => s + (parseFloat(r['Total Amount']) || 0), 0);
    const cod = this.dispDateRows.reduce((s, r) => s + (parseFloat(r['COD Amount']) || 0), 0);
    const pre = this.dispDateRows.filter(r => r['Payment Type'] === 'Prepaid').length;
    const codC = this.dispDateRows.filter(r => r['Payment Type'] === 'COD').length;
    const bar = this.q('dispatch-stat-bar');
    bar.style.display = 'flex';
    const items: any[] = [
      { l: 'Total Orders', v: this.dispDateRows.length, c: 'var(--accent)' },
      { l: 'Prepaid', v: pre, c: 'var(--green)' },
      { l: 'COD', v: codC, c: 'var(--orange)' },
      { l: 'Total Value', v: this.fk(tot), c: 'var(--text)' },
      { l: 'COD Exposure', v: this.fk(cod), c: 'var(--orange)' },
    ];
    bar.innerHTML = items.map(x => `<div class="stat-item"><div class="stat-val" style="color:${x.c}">${x.v}</div><div class="stat-lbl">${x.l}</div></div>`).join('');
    this.filterDispatch();
  }

  private filterDispatch(): void {
    const search = this.q<HTMLInputElement>('disp-search').value.toLowerCase();
    const pay = this.q<HTMLSelectElement>('disp-pay').value;
    const zone = this.q<HTMLSelectElement>('disp-zone').value;
    const rows = this.dispDateRows.filter(r => {
      if (pay && r['Payment Type'] !== pay) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
    this.q('disp-count').textContent = `${rows.length} orders`;
    this.q('disp-tbody').innerHTML = rows.map((r, i) => `<tr>
      <td style="color:var(--muted)">${i + 1}</td>
      <td style="font-family:monospace;font-size:11px;color:var(--accent)">${r['Shipment ID'] || '—'}</td>
      <td style="font-family:monospace;font-size:11px">${r['AWB'] || '—'}</td>
      <td style="font-size:11px;color:var(--muted)">${r['Order Number'] || '—'}</td>
      <td style="font-weight:600">${r['Customer'] || '—'}</td>
      <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
      <td>${this.pBadge(r['Payment Type'])}</td>
      <td>${this.sBadge(r['Shipment Status'])}</td>
      <td style="font-weight:700;color:var(--accent2)">Zone ${r['Zone'] || '?'}</td>
      <td style="font-weight:600">${this.fc(r['Total Amount'])}</td>
      <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
      <td style="text-align:center">${r['Total Quantity'] || '—'}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;font-size:11px;color:var(--muted2)" title="${r['Product Name'] || ''}">${String(r['Product Name'] || '—').substring(0, 28)}${String(r['Product Name'] || '').length > 28 ? '…' : ''}</td>
      <td>${this.tagChips(r['Tags'])}</td>
    </tr>`).join('');
  }

  private downloadDispatch(): void {
    this.excelDownload(this.dispDateRows, `BANDB_Dispatch_${this.q<HTMLSelectElement>('dispatch-date-sel').value.replace(/ /g, '_')}`);
  }

  // ═══════════════════════════════════════════════════════
  // DELIVERED BY DATE
  // ═══════════════════════════════════════════════════════

  private setupDeliveryPage(): void {
    const dates = [...new Set(this.delRows.map(r => r['Delivery Date']))].sort((a, b) => this.dateSortKey(b) - this.dateSortKey(a));
    this.q<HTMLSelectElement>('delivery-date-sel').innerHTML = dates.map(d => `<option value="${d}">${d} (${this.delRows.filter(r => r['Delivery Date'] === d).length} delivered)</option>`).join('');
    const zones = [...new Set(this.delRows.map(r => r['Zone']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('del-zone').innerHTML = '<option value="">All Zones</option>' + zones.map(z => `<option value="${z}">Zone ${z}</option>`).join('');
    this.loadDeliveryDate();
  }

  private loadDeliveryDate(): void {
    const date = this.q<HTMLSelectElement>('delivery-date-sel').value;
    this.delDateRows = this.delRows.filter(r => r['Delivery Date'] === date);
    const tot = this.delDateRows.reduce((s, r) => s + (parseFloat(r['Total Amount']) || 0), 0);
    const cod = this.delDateRows.reduce((s, r) => s + (parseFloat(r['COD Amount']) || 0), 0);
    const pre = this.delDateRows.filter(r => r['Payment Type'] === 'Prepaid').length;
    const codC = this.delDateRows.filter(r => r['Payment Type'] === 'COD').length;

    const dtVals = this.delDateRows.map(r => this.daysTaken(r['Pickup Date'], r['Delivery Date'])).filter((v): v is number => v !== null);
    const avgDays = dtVals.length ? (dtVals.reduce((a, b) => a + b, 0) / dtVals.length).toFixed(1) : '—';

    const lateCount = this.delDateRows.filter(r => r['TAT Breach'] && String(r['TAT Breach']).trim()).length;
    const onTimeCount = this.delDateRows.length - lateCount;

    const zoneDays: Record<string, number[]> = {};
    this.delDateRows.forEach(r => {
      const z = r['Zone'] || '?';
      const dt = this.daysTaken(r['Pickup Date'], r['Delivery Date']);
      if (dt !== null) { if (!zoneDays[z]) zoneDays[z] = []; zoneDays[z].push(dt); }
    });
    const zoneAvgStr = Object.keys(zoneDays).sort().map(z => {
      const arr = zoneDays[z]; const avg = (arr.reduce((a, b) => a + b, 0) / arr.length).toFixed(1);
      return `Z${z}: ${avg}d`;
    }).join(' · ') || '—';

    const items: any[] = [
      { l: 'Delivered Orders', v: this.delDateRows.length, c: 'var(--green)' },
      { l: 'Avg Days Taken', v: avgDays, c: 'var(--accent)' },
      { l: 'On-time', v: onTimeCount, c: 'var(--green)' },
      { l: 'Late (TAT Breach)', v: lateCount, c: 'var(--red)' },
      { l: 'Prepaid / COD', v: `${pre} / ${codC}`, c: 'var(--text)' },
      { l: 'Total Value', v: this.fk(tot), c: 'var(--text)' },
      { l: 'COD Value', v: this.fk(cod), c: 'var(--orange)' },
    ];
    this.q('delivery-stat-bar').style.display = 'flex';
    this.q('delivery-stat-bar').innerHTML = items.map(x => `<div class="stat-item"><div class="stat-val" style="color:${x.c}">${x.v}</div><div class="stat-lbl">${x.l}</div></div>`).join('')
      + `<div class="stat-item" style="flex:2;min-width:220px"><div class="stat-val" style="font-size:13px;color:var(--accent2)">${zoneAvgStr}</div><div class="stat-lbl">Avg Delivery Days by Zone</div></div>`;

    this.filterDelivery();
  }

  private filterDelivery(): void {
    const search = this.q<HTMLInputElement>('del-search').value.toLowerCase();
    const pay = this.q<HTMLSelectElement>('del-pay').value;
    const zone = this.q<HTMLSelectElement>('del-zone').value;
    const tatF = this.q<HTMLSelectElement>('del-tat').value;
    const ndrF = this.q<HTMLSelectElement>('del-ndr').value;

    const rows = this.delDateRows.filter(r => {
      if (pay && r['Payment Type'] !== pay) return false;
      if (zone && r['Zone'] !== zone) return false;
      const isLate = !!(r['TAT Breach'] && String(r['TAT Breach']).trim());
      if (tatF === 'ontime' && isLate) return false;
      if (tatF === 'late' && !isLate) return false;
      const hadNdr = !!(r['Latest NDR Status'] && String(r['Latest NDR Status']).trim());
      if (ndrF === 'yes' && !hadNdr) return false;
      if (ndrF === 'no' && hadNdr) return false;
      if (search) { const h = (r['Customer'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
    this.q('del-count').textContent = `${rows.length} orders`;
    this.q('del-tbody').innerHTML = rows.map((r, i) => {
      const dt = this.daysTaken(r['Pickup Date'], r['Delivery Date']);
      const isLate = !!(r['TAT Breach'] && String(r['TAT Breach']).trim());
      const ndrText = r['Latest NDR Status'] && String(r['Latest NDR Status']).trim() ? r['Latest NDR Status'] : 'No NDR';
      const ndrBadgeClass = r['Latest NDR Status'] && String(r['Latest NDR Status']).trim() ? 'b-ndrflag' : 'b-noissue';
      return `<tr>
      <td style="color:var(--muted)">${i + 1}</td>
      <td style="font-family:monospace;font-size:11px">${r['AWB'] || '—'}</td>
      <td style="font-weight:600">${r['Customer'] || '—'}</td>
      <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
      <td>${this.pBadge(r['Payment Type'])}</td>
      <td style="font-size:11px;color:var(--muted)">${r['Pickup Date'] || '—'}</td>
      <td style="font-weight:700;color:var(--green)">${r['Delivery Date'] || '—'}</td>
      <td>${this.daysTakenCell(dt)}</td>
      <td style="text-align:center">${r['Attempts'] || '—'}</td>
      <td><span class="badge ${isLate ? 'b-late' : 'b-ontime'}">${isLate ? 'Late' : 'On-time'}</span></td>
      <td><span class="badge ${ndrBadgeClass}" title="${r['Latest NDR Status'] || ''}">${String(ndrText).length > 16 ? String(ndrText).substring(0, 16) + '…' : ndrText}</span></td>
      <td style="font-weight:700;color:var(--accent2)">Zone ${r['Zone'] || '?'}</td>
      <td style="font-weight:600">${this.fc(r['Total Amount'])}</td>
      <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
      <td style="text-align:center">${r['Total Quantity'] || '—'}</td>
      <td style="max-width:180px;overflow:hidden;text-overflow:ellipsis;font-size:11px;color:var(--muted2)" title="${r['Product Name'] || ''}">${String(r['Product Name'] || '—').substring(0, 28)}${String(r['Product Name'] || '').length > 28 ? '…' : ''}</td>
    </tr>`;
    }).join('');
  }

  private getCurrentDeliveryRows(): Row[] {
    const search = this.q<HTMLInputElement>('del-search').value.toLowerCase();
    const pay = this.q<HTMLSelectElement>('del-pay').value;
    const zone = this.q<HTMLSelectElement>('del-zone').value;
    const tatF = this.q<HTMLSelectElement>('del-tat').value;
    const ndrF = this.q<HTMLSelectElement>('del-ndr').value;
    return this.delDateRows.filter(r => {
      if (pay && r['Payment Type'] !== pay) return false;
      if (zone && r['Zone'] !== zone) return false;
      const isLate = !!(r['TAT Breach'] && String(r['TAT Breach']).trim());
      if (tatF === 'ontime' && isLate) return false;
      if (tatF === 'late' && !isLate) return false;
      const hadNdr = !!(r['Latest NDR Status'] && String(r['Latest NDR Status']).trim());
      if (ndrF === 'yes' && !hadNdr) return false;
      if (ndrF === 'no' && hadNdr) return false;
      if (search) { const h = (r['Customer'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
  }

  private downloadDelivery(): void {
    const rows = this.getCurrentDeliveryRows().map(r => ({
      'AWB': r['AWB'],
      'Customer': r['Customer'],
      'Phone #': r['Phone #'],
      'Payment Type': r['Payment Type'],
      'Pickup Date': r['Pickup Date'],
      'Delivery Date': r['Delivery Date'],
      'Days Taken': this.daysTaken(r['Pickup Date'], r['Delivery Date']),
      'Attempts': r['Attempts'],
      'On-time / Late': (r['TAT Breach'] && String(r['TAT Breach']).trim()) ? 'Late' : 'On-time',
      'NDR History': (r['Latest NDR Status'] && String(r['Latest NDR Status']).trim()) ? r['Latest NDR Status'] : 'No NDR',
      'Zone': r['Zone'],
      'Total Amount': r['Total Amount'],
      'COD Amount': r['COD Amount'],
      'Total Quantity': r['Total Quantity'],
      'Product Name': r['Product Name'],
    }));
    this.excelDownload(rows, `BANDB_Delivered_${this.q<HTMLSelectElement>('delivery-date-sel').value.replace(/ /g, '_')}`);
  }

  // ═══════════════════════════════════════════════════════
  // PENDING 5+ DAYS
  // ═══════════════════════════════════════════════════════

  private setupPendingPage(): void {
    const statuses = [...new Set(this.nd5Rows.map(r => r['Shipment Status']))].sort();
    this.q<HTMLSelectElement>('pend-status').innerHTML = '<option value="">All Statuses</option>' + statuses.map(s => `<option value="${s}">${s}</option>`).join('');
    const zones = [...new Set(this.nd5Rows.map(r => r['Zone']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('pend-zone').innerHTML = '<option value="">All Zones</option>' + zones.map(z => `<option value="${z}">Zone ${z}</option>`).join('');

    const cod = this.nd5Rows.filter(r => r['Payment Type'] === 'COD').length;
    const codAmt = this.nd5Rows.reduce((s, r) => s + (parseFloat(r['COD Amount']) || 0), 0);
    const ndr = this.nd5Rows.filter(r => r['Shipment Status'] === 'Undelivered').length;
    const items: any[] = [
      { l: 'Total Pending', v: this.nd5Rows.length, c: 'var(--red)' },
      { l: 'COD Orders', v: cod, c: 'var(--orange)' },
      { l: 'COD at Risk', v: this.fk(codAmt), c: 'var(--orange)' },
      { l: 'NDR Cases', v: ndr, c: 'var(--red)' },
      { l: 'TAT Breached', v: this.nd5Rows.filter(r => r['TAT Breach'] && String(r['TAT Breach']).trim()).length, c: 'var(--red)' },
    ];
    this.q('pending-stat-bar').innerHTML = items.map(x => `<div class="stat-item"><div class="stat-val" style="color:${x.c}">${x.v}</div><div class="stat-lbl">${x.l}</div></div>`).join('');
    this.filterPending();
  }

  private filterPending(): void {
    const search = this.q<HTMLInputElement>('pend-search').value.toLowerCase();
    const status = this.q<HTMLSelectElement>('pend-status').value;
    const zone = this.q<HTMLSelectElement>('pend-zone').value;
    const pay = this.q<HTMLSelectElement>('pend-pay').value;
    let rows = this.nd5Rows.filter(r => {
      if (status && r['Shipment Status'] !== status) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
    rows = [...rows].sort((a, b) => (this.daysBetween(b['Pickup Date']) || 0) - (this.daysBetween(a['Pickup Date']) || 0));
    this.q('pend-count').textContent = `${rows.length} shipments`;
    this.q('pend-tbody').innerHTML = rows.map((r, i) => {
      const days = this.daysBetween(r['Pickup Date']);
      return `<tr>
        <td style="color:var(--muted)">${i + 1}</td>
        <td style="font-family:monospace;font-size:11px;color:var(--red)">${r['Shipment ID'] || '—'}</td>
        <td style="font-family:monospace;font-size:11px">${r['AWB'] || '—'}</td>
        <td style="font-weight:600">${r['Customer'] || '—'}</td>
        <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
        <td>${this.pBadge(r['Payment Type'])}</td>
        <td>${this.sBadge(r['Shipment Status'])}</td>
        <td style="color:var(--muted);font-size:11px">${r['Pickup Date'] || '—'}</td>
        <td>${this.daysCell(days)}</td>
        <td style="font-weight:700;color:var(--accent2)">Zone ${r['Zone'] || '?'}</td>
        <td>${this.fc(r['Total Amount'])}</td>
        <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
        <td style="font-size:11px;color:var(--red)">${r['TAT Breach'] || '—'}</td>
        <td style="text-align:center">${r['Attempts'] || '—'}</td>
        <td style="font-size:11px;color:var(--orange)">${r['Latest NDR Status'] || '—'}</td>
        <td>${this.tagChips(r['Tags'])}</td>
        <td style="font-size:11px;color:var(--muted)">${String(r['Latest Status Date'] || '—').replace(', ', ' ')}</td>
      </tr>`;
    }).join('');
  }

  private downloadPending(): void {
    const search = this.q<HTMLInputElement>('pend-search').value.toLowerCase();
    const status = this.q<HTMLSelectElement>('pend-status').value;
    const zone = this.q<HTMLSelectElement>('pend-zone').value;
    const pay = this.q<HTMLSelectElement>('pend-pay').value;
    const rows = this.nd5Rows.filter(r => {
      if (status && r['Shipment Status'] !== status) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
    const out = rows.map(r => ({ ...r, 'Days Since Pickup': this.daysBetween(r['Pickup Date']) || '' }));
    this.excelDownload(out, 'BANDB_Pending_5Days');
  }

  // ═══════════════════════════════════════════════════════
  // 2+ ATTEMPTS (NOT DELIVERED) — MONITOR TAB
  // ═══════════════════════════════════════════════════════

  private setupAttemptsPage(): void {
    const statuses = [...new Set(this.attRows.map(r => r['Shipment Status']))].sort();
    this.q<HTMLSelectElement>('att-status').innerHTML = '<option value="">All Statuses</option>' + statuses.map(s => `<option value="${s}">${s}</option>`).join('');
    const zones = [...new Set(this.attRows.map(r => r['Zone']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('att-zone').innerHTML = '<option value="">All Zones</option>' + zones.map(z => `<option value="${z}">Zone ${z}</option>`).join('');

    const cod = this.attRows.filter(r => r['Payment Type'] === 'COD').length;
    const codAmt = this.attRows.reduce((s, r) => s + (parseFloat(r['COD Amount']) || 0), 0);
    const ndr = this.attRows.filter(r => r['Shipment Status'] === 'Undelivered').length;
    const tatBrCount = this.attRows.filter(r => r['TAT Breach'] && String(r['TAT Breach']).trim()).length;
    const avgAttempts = this.attRows.length ? (this.attRows.reduce((s, r) => s + (parseInt(r['Attempts']) || 0), 0) / this.attRows.length).toFixed(1) : '0';

    const items: any[] = [
      { l: 'Total (2+ Attempts)', v: this.attRows.length, c: 'var(--red)' },
      { l: 'Avg Attempts', v: avgAttempts, c: 'var(--orange)' },
      { l: 'COD Orders', v: cod, c: 'var(--orange)' },
      { l: 'COD at Risk', v: this.fk(codAmt), c: 'var(--orange)' },
      { l: 'NDR (Undelivered)', v: ndr, c: 'var(--red)' },
      { l: 'TAT Breached', v: tatBrCount, c: 'var(--red)' },
    ];
    this.q('attempts-stat-bar').innerHTML = items.map(x => `<div class="stat-item"><div class="stat-val" style="color:${x.c}">${x.v}</div><div class="stat-lbl">${x.l}</div></div>`).join('');
    this.filterAttempts();
  }

  private getCurrentAttemptsRows(): Row[] {
    const search = this.q<HTMLInputElement>('att-search').value.toLowerCase();
    const status = this.q<HTMLSelectElement>('att-status').value;
    const zone = this.q<HTMLSelectElement>('att-zone').value;
    const pay = this.q<HTMLSelectElement>('att-pay').value;
    return this.attRows.filter(r => {
      if (status && r['Shipment Status'] !== status) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
  }

  private filterAttempts(): void {
    const rows = this.getCurrentAttemptsRows();
    this.q('att-count').textContent = `${rows.length} shipments`;
    this.q('att-tbody').innerHTML = rows.map((r, i) => {
      const days = this.daysBetween(r['Pickup Date']);
      return `<tr>
        <td style="color:var(--muted)">${i + 1}</td>
        <td style="font-family:monospace;font-size:11px;color:var(--red)">${r['Shipment ID'] || '—'}</td>
        <td style="font-family:monospace;font-size:11px">${r['AWB'] || '—'}</td>
        <td style="font-weight:600">${r['Customer'] || '—'}</td>
        <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
        <td>${this.pBadge(r['Payment Type'])}</td>
        <td>${this.sBadge(r['Shipment Status'])}</td>
        <td style="color:var(--muted);font-size:11px">${r['Pickup Date'] || '—'}</td>
        <td>${this.daysCell(days)}</td>
        <td style="text-align:center;font-weight:900;color:var(--red)">${r['Attempts'] || '—'}</td>
        <td style="font-weight:700;color:var(--accent2)">Zone ${r['Zone'] || '?'}</td>
        <td>${this.fc(r['Total Amount'])}</td>
        <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
        <td style="font-size:11px;color:var(--red)">${r['TAT Breach'] || '—'}</td>
        <td style="font-size:11px;color:var(--orange)">${r['Latest NDR Status'] || '—'}</td>
        <td>${this.tagChips(r['Tags'])}</td>
        <td style="font-size:11px;color:var(--muted)">${String(r['Latest Status Date'] || '—').replace(', ', ' ')}</td>
      </tr>`;
    }).join('');
  }

  private downloadAttempts(): void {
    const rows = this.getCurrentAttemptsRows().map(r => ({ ...r, 'Days Since Pickup': this.daysBetween(r['Pickup Date']) || '' }));
    this.excelDownload(rows, 'BANDB_2plus_Attempts_Not_Delivered');
  }

  // ═══════════════════════════════════════════════════════
  // LAST MONTH
  // ═══════════════════════════════════════════════════════

  private getLastMonthRange(): { y: number; m: number } {
    let m = this.today.getMonth() - 1, y = this.today.getFullYear();
    if (m < 0) { m = 11; y--; }
    return { y, m };
  }

  private setupLastMonthPage(): void {
    const monthSet: Record<string, number> = {};
    this.ALL.forEach(r => {
      const d = this.parseDate(r['Pickup Date']);
      if (d) { const key = `${d.getFullYear()}-${String(d.getMonth()).padStart(2, '0')}`; monthSet[key] = (monthSet[key] || 0) + 1; }
    });
    const monthKeys = Object.keys(monthSet).sort((a, b) => b.localeCompare(a));

    if (!monthKeys.length) {
      this.q<HTMLSelectElement>('lm-month-sel').innerHTML = '<option value="">No dated shipments</option>';
      this.q('lm-label').textContent = 'No Pickup Date data available.';
      this.lmRows = [];
      this.renderLastMonthAll();
      return;
    }

    this.q<HTMLSelectElement>('lm-month-sel').innerHTML = monthKeys.map(key => {
      const [y, m] = key.split('-').map(Number);
      const label = new Date(y, m, 1).toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
      return `<option value="${key}">${label} (${monthSet[key]} orders)</option>`;
    }).join('');

    const { y: defY, m: defM } = this.getLastMonthRange();
    const defKey = `${defY}-${String(defM).padStart(2, '0')}`;
    this.q<HTMLSelectElement>('lm-month-sel').value = monthKeys.includes(defKey) ? defKey : monthKeys[0];

    this.loadLastMonthSelected();
  }

  private loadLastMonthSelected(): void {
    const val = this.q<HTMLSelectElement>('lm-month-sel').value;
    if (!val) { this.lmRows = []; this.renderLastMonthAll(); return; }
    const [y, m] = val.split('-').map(Number);
    this.lmYear = y; this.lmMonth = m;
    const monthName = new Date(y, m, 1).toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
    this.q('lm-label').textContent = `Showing: ${monthName} (based on Pickup Date)`;

    this.lmRows = this.ALL.filter(r => {
      const d = this.parseDate(r['Pickup Date']);
      return d && d.getFullYear() === y && d.getMonth() === m;
    });

    this.renderLastMonthAll();
  }

  private renderLastMonthAll(): void {
    const statuses = [...new Set(this.lmRows.map(r => r['Shipment Status']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('lm-status').innerHTML = '<option value="">All Statuses</option>' + statuses.map(s => `<option value="${s}">${s}</option>`).join('');
    const zones = [...new Set(this.lmRows.map(r => r['Zone']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('lm-zone').innerHTML = '<option value="">All Zones</option>' + zones.map(z => `<option value="${z}">Zone ${z}</option>`).join('');

    const sm: Record<string, number> = {};
    this.lmRows.forEach(r => { const s = r['Shipment Status'] || 'Unknown'; sm[s] = (sm[s] || 0) + 1; });
    const totalN = this.lmRows.length;
    const pctOfTotal = (nn: number) => totalN ? ((nn / totalN) * 100).toFixed(1) : '0.0';
    const dispatchedN = this.lmRows.filter(r => r['Pickup Date'] && r['Pickup Date'] !== 'Not Picked' && r['Pickup Date'] !== '--').length;
    const pre = this.lmRows.filter(r => r['Payment Type'] === 'Prepaid').length;
    const codC = this.lmRows.filter(r => r['Payment Type'] === 'COD').length;
    const rtoRows = this.lmRows.filter(r => r['Shipment Status'] === 'RTO In Transit');
    const rtoPre = rtoRows.filter(r => r['Payment Type'] === 'Prepaid').length;
    const rtoCod = rtoRows.filter(r => r['Payment Type'] === 'COD').length;
    const rtoPrePct = pctOfTotal(rtoPre);
    const rtoCodPct = pctOfTotal(rtoCod);
    const rtoPct = totalN ? ((rtoRows.length / totalN) * 100).toFixed(2) : '0.00';
    const statItems: any[] = [
      { l: 'Total Orders', v: `${totalN} (100.0%)`, c: 'var(--accent)' },
      { l: 'Dispatched', v: `${dispatchedN} (${pctOfTotal(dispatchedN)}%)`, c: 'var(--cyan)' },
      { l: 'Delivered', v: `${sm['Delivered'] || 0} (${pctOfTotal(sm['Delivered'] || 0)}%)`, c: 'var(--green)' },
      { l: 'In Transit', v: `${sm['In Transit'] || 0} (${pctOfTotal(sm['In Transit'] || 0)}%)`, c: 'var(--accent)' },
      { l: 'Out for Delivery', v: `${sm['Out for Delivery'] || 0} (${pctOfTotal(sm['Out for Delivery'] || 0)}%)`, c: 'var(--accent2)' },
      { l: 'Undelivered (NDR)', v: `${sm['Undelivered'] || 0} (${pctOfTotal(sm['Undelivered'] || 0)}%)`, c: 'var(--orange)' },
      { l: 'RTO Total', v: `${rtoRows.length} (${rtoPct}%)`, c: 'var(--red)' },
      { l: 'RTO — Prepaid', v: `${rtoPre} (${rtoPrePct}%)`, c: 'var(--accent)' },
      { l: 'RTO — COD', v: `${rtoCod} (${rtoCodPct}%)`, c: 'var(--orange)' },
      { l: 'Prepaid / COD', v: `${pre} (${pctOfTotal(pre)}%) / ${codC} (${pctOfTotal(codC)}%)`, c: 'var(--text)' },
    ];
    this.q('lm-stat-bar').innerHTML = statItems.map(x => `<div class="stat-item"><div class="stat-val" style="color:${x.c}">${x.v}</div><div class="stat-lbl">${x.l}</div></div>`).join('');

    const statusesAll = [...new Set([...this.STATUS_LIST, ...Object.keys(sm)])].filter(s => sm[s]);
    this.q('lm-summary-tbody').innerHTML = statusesAll.map(s => {
      const rows = this.lmRows.filter(r => r['Shipment Status'] === s);
      const preR = rows.filter(r => r['Payment Type'] === 'Prepaid');
      const codR = rows.filter(r => r['Payment Type'] === 'COD');
      const preVal = preR.reduce((a, r) => a + (parseFloat(r['Total Amount']) || 0), 0);
      const codVal = codR.reduce((a, r) => a + (parseFloat(r['Total Amount']) || 0), 0);
      const prePct = rows.length ? ((preR.length / rows.length) * 100).toFixed(1) : '0.0';
      const codPct = rows.length ? ((codR.length / rows.length) * 100).toFixed(1) : '0.0';
      return `<tr>
        <td>${this.sBadge(s)}</td>
        <td style="font-weight:700;color:var(--accent)">${preR.length}</td>
        <td style="font-weight:700;color:var(--orange)">${codR.length}</td>
        <td style="font-weight:900">${rows.length}</td>
        <td style="color:var(--accent)">${prePct}%</td>
        <td style="color:var(--orange)">${codPct}%</td>
        <td>${this.fc(preVal)}</td>
        <td style="color:var(--orange)">${this.fc(codVal)}</td>
        <td style="font-weight:700">${this.fc(preVal + codVal)}</td>
      </tr>`;
    }).join('')
      + `<tr style="border-top:2px solid var(--border2)">
        <td style="font-weight:900">TOTAL</td>
        <td style="font-weight:900;color:var(--accent)">${pre}</td>
        <td style="font-weight:900;color:var(--orange)">${codC}</td>
        <td style="font-weight:900">${this.lmRows.length}</td>
        <td style="font-weight:900;color:var(--accent)">${this.lmRows.length ? ((pre / this.lmRows.length) * 100).toFixed(1) : '0.0'}%</td>
        <td style="font-weight:900;color:var(--orange)">${this.lmRows.length ? ((codC / this.lmRows.length) * 100).toFixed(1) : '0.0'}%</td>
        <td style="font-weight:900" colspan="3">${this.fc(this.lmRows.reduce((a, r) => a + (parseFloat(r['Total Amount']) || 0), 0))}</td>
      </tr>`;

    this.destroyChart('lm-status');
    const sc: Record<string, string> = { 'Delivered': '#22c55e', 'In Transit': '#4f8ef7', 'Picked Up': '#a855f7', 'Order Placed': '#06b6d4', 'Out for Delivery': '#7c5cfc', 'Undelivered': '#f59e0b', 'RTO In Transit': '#ef4444', 'RTO Delivered': '#ec4899', 'Cancelled': '#6b7280', 'Shipment Delayed': '#dc2626' };
    this.chartMap['lm-status'] = new Chart(this.q('lm-status-chart'), { type: 'bar', data: { labels: Object.keys(sm), datasets: [{ data: Object.values(sm), backgroundColor: Object.keys(sm).map(s => sc[s] || '#4f8ef7'), borderRadius: 6 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor, font: { size: 9 } }, grid: { display: false } }, y: { ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    this.destroyChart('lm-pay');
    const labelsPay = statusesAll;
    this.chartMap['lm-pay'] = new Chart(this.q('lm-pay-chart'), { type: 'bar', data: { labels: labelsPay, datasets: [
      { label: 'Prepaid', data: labelsPay.map(s => this.lmRows.filter(r => r['Shipment Status'] === s && r['Payment Type'] === 'Prepaid').length), backgroundColor: '#4f8ef7', borderRadius: 4 },
      { label: 'COD', data: labelsPay.map(s => this.lmRows.filter(r => r['Shipment Status'] === s && r['Payment Type'] === 'COD').length), backgroundColor: '#f59e0b', borderRadius: 4 },
    ] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 10 } } } }, scales: { x: { stacked: true, ticks: { color: this.tickColor, font: { size: 9 } }, grid: { display: false } }, y: { stacked: true, ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    this.filterLastMonth();
  }

  private filterLastMonth(): void {
    const search = this.q<HTMLInputElement>('lm-search').value.toLowerCase();
    const status = this.q<HTMLSelectElement>('lm-status').value;
    const pay = this.q<HTMLSelectElement>('lm-pay').value;
    const zone = this.q<HTMLSelectElement>('lm-zone').value;
    const rows = this.lmRows.filter(r => {
      if (status && r['Shipment Status'] !== status) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
    this.q('lm-count').textContent = `${rows.length} orders`;
    this.q('lm-tbody').innerHTML = rows.map((r, i) => `<tr>
      <td style="color:var(--muted)">${i + 1}</td>
      <td style="font-family:monospace;font-size:11px;color:var(--accent)">${r['Shipment ID'] || '—'}</td>
      <td style="font-family:monospace;font-size:11px">${r['AWB'] || '—'}</td>
      <td style="font-weight:600">${r['Customer'] || '—'}</td>
      <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
      <td>${this.pBadge(r['Payment Type'])}</td>
      <td>${this.sBadge(r['Shipment Status'])}</td>
      <td style="color:var(--muted);font-size:11px">${r['Pickup Date'] || '—'}</td>
      <td style="color:var(--muted);font-size:11px">${r['Delivery Date'] || '—'}</td>
      <td style="font-weight:700;color:var(--accent2)">Zone ${r['Zone'] || '?'}</td>
      <td>${this.fc(r['Total Amount'])}</td>
      <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
      <td>${this.tagChips(r['Tags'])}</td>
    </tr>`).join('');
  }

  private getCurrentLastMonthRows(): Row[] {
    const search = this.q<HTMLInputElement>('lm-search').value.toLowerCase();
    const status = this.q<HTMLSelectElement>('lm-status').value;
    const pay = this.q<HTMLSelectElement>('lm-pay').value;
    const zone = this.q<HTMLSelectElement>('lm-zone').value;
    return this.lmRows.filter(r => {
      if (status && r['Shipment Status'] !== status) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (zone && r['Zone'] !== zone) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
  }

  private downloadLastMonth(): void {
    const monthName = new Date(this.lmYear || 0, this.lmMonth || 0, 1).toLocaleDateString('en-IN', { month: 'short', year: 'numeric' });
    this.excelDownload(this.getCurrentLastMonthRows(), `BANDB_LastMonth_${monthName.replace(' ', '_')}`);
  }

  private downloadLastMonthSummary(): void {
    const sm: Record<string, number> = {};
    this.lmRows.forEach(r => { const s = r['Shipment Status'] || 'Unknown'; sm[s] = (sm[s] || 0) + 1; });
    const statusesAll = [...new Set([...this.STATUS_LIST, ...Object.keys(sm)])].filter(s => sm[s]);
    const out = statusesAll.map(s => {
      const rows = this.lmRows.filter(r => r['Shipment Status'] === s);
      const preR = rows.filter(r => r['Payment Type'] === 'Prepaid');
      const codR = rows.filter(r => r['Payment Type'] === 'COD');
      return {
        'Status': s,
        'Prepaid Orders': preR.length,
        'COD Orders': codR.length,
        'Total Orders': rows.length,
        'Prepaid %': rows.length ? ((preR.length / rows.length) * 100).toFixed(1) : '0.0',
        'COD %': rows.length ? ((codR.length / rows.length) * 100).toFixed(1) : '0.0',
        'Prepaid Value': preR.reduce((a, r) => a + (parseFloat(r['Total Amount']) || 0), 0),
        'COD Value': codR.reduce((a, r) => a + (parseFloat(r['Total Amount']) || 0), 0),
      };
    });
    const monthName = new Date(this.lmYear || 0, this.lmMonth || 0, 1).toLocaleDateString('en-IN', { month: 'short', year: 'numeric' });
    this.excelDownload(out, `BANDB_LastMonth_Summary_${monthName.replace(' ', '_')}`);
  }

  // ═══════════════════════════════════════════════════════
  // ZONE / CITY
  // ═══════════════════════════════════════════════════════

  private getZoneField(): string {
    if (!this.ALL.length) return 'Zone';
    const keys = Object.keys(this.ALL[0]);
    const cityKey = keys.find(k => /city/i.test(k) && !/capacity/i.test(k));
    return cityKey || 'Zone';
  }

  private setupZonePage(): void {
    this.znField = this.getZoneField();
    const isCity = this.znField !== 'Zone';
    this.q('zone-summary-title').textContent = isCity ? `${this.znField}-wise Summary` : 'Zone-wise Summary (no City column found in data — grouped by Zone)';
    this.q('zone-col-hdr').textContent = isCity ? this.znField : 'Zone';
    this.q('zone-label').textContent = isCity ? `Grouped by ${this.znField}` : 'No city-level field found in this CSV — using Zone (A/B/C/D) as the geographic grouping. Upload a file with a "City" column to see true city-wise data.';

    this.znGroups = [...new Set(this.ALL.map(r => r[this.znField]).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('zn-zone').innerHTML = '<option value="">All</option>' + this.znGroups.map(z => `<option value="${z}">${isCity ? z : 'Zone ' + z}</option>`).join('');
    const statuses = [...new Set(this.ALL.map(r => r['Shipment Status']).filter(Boolean))].sort();
    this.q<HTMLSelectElement>('zn-status').innerHTML = '<option value="">All Statuses</option>' + statuses.map(s => `<option value="${s}">${s}</option>`).join('');

    const summary = this.znGroups.map(z => {
      const rows = this.ALL.filter(r => r[this.znField] === z);
      const pre = rows.filter(r => r['Payment Type'] === 'Prepaid').length;
      const cod = rows.filter(r => r['Payment Type'] === 'COD').length;
      const del = rows.filter(r => r['Shipment Status'] === 'Delivered').length;
      const it = rows.filter(r => r['Shipment Status'] === 'In Transit').length;
      const ofd = rows.filter(r => r['Shipment Status'] === 'Out for Delivery').length;
      const ndr = rows.filter(r => r['Shipment Status'] === 'Undelivered').length;
      const rto = rows.filter(r => r['Shipment Status'] === 'RTO In Transit').length;
      const rev = rows.reduce((a, r) => a + (parseFloat(r['Total Amount']) || 0), 0);
      const codVal = rows.reduce((a, r) => a + (parseFloat(r['COD Amount']) || 0), 0);
      const rate = rows.length ? ((del / rows.length) * 100).toFixed(1) : '0.0';
      return { z, total: rows.length, pre, cod, del, it, ofd, ndr, rto, rev, codVal, rate };
    }).sort((a, b) => b.total - a.total);

    this.q('zone-summary-tbody').innerHTML = summary.map(x => `<tr>
      <td style="font-weight:700;color:var(--accent2)">${isCity ? x.z : 'Zone ' + x.z}</td>
      <td style="font-weight:900">${x.total}</td>
      <td style="color:var(--accent)">${x.pre}</td>
      <td style="color:var(--orange)">${x.cod}</td>
      <td style="color:var(--green)">${x.del}</td>
      <td>${x.it}</td>
      <td>${x.ofd}</td>
      <td style="color:var(--orange)">${x.ndr}</td>
      <td style="color:var(--red)">${x.rto}</td>
      <td style="font-weight:700;color:${parseFloat(x.rate) >= 85 ? 'var(--green)' : 'var(--red)'}">${x.rate}%</td>
      <td>${this.fc(x.rev)}</td>
      <td style="color:var(--orange)">${this.fc(x.codVal)}</td>
    </tr>`).join('');

    this.destroyChart('zone-pay');
    this.chartMap['zone-pay'] = new Chart(this.q('zone-pay-chart'), { type: 'bar', data: { labels: summary.map(x => isCity ? x.z : 'Zone ' + x.z), datasets: [
      { label: 'Prepaid', data: summary.map(x => x.pre), backgroundColor: '#4f8ef7', borderRadius: 4 },
      { label: 'COD', data: summary.map(x => x.cod), backgroundColor: '#f59e0b', borderRadius: 4 },
    ] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 10 } } } }, scales: { x: { stacked: true, ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { stacked: true, ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    this.destroyChart('zone-rate');
    this.chartMap['zone-rate'] = new Chart(this.q('zone-rate-chart'), { type: 'bar', data: { labels: summary.map(x => isCity ? x.z : 'Zone ' + x.z), datasets: [{ data: summary.map(x => parseFloat(x.rate)), backgroundColor: summary.map(x => parseFloat(x.rate) >= 85 ? '#22c55e' : parseFloat(x.rate) >= 70 ? '#f59e0b' : '#ef4444'), borderRadius: 6 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { ticks: { color: this.tickColor, callback: (v: any) => v + '%' }, grid: { color: this.gridColor } } } } });

    this.filterZoneOrders();
  }

  private filterZoneOrders(): void {
    const search = this.q<HTMLInputElement>('zn-search').value.toLowerCase();
    const zone = this.q<HTMLSelectElement>('zn-zone').value;
    const status = this.q<HTMLSelectElement>('zn-status').value;
    const pay = this.q<HTMLSelectElement>('zn-pay').value;
    const rows = this.ALL.filter(r => {
      if (zone && r[this.znField] !== zone) return false;
      if (status && r['Shipment Status'] !== status) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
    this.q('zn-count').textContent = `${rows.length} orders`;
    this.q('zn-tbody').innerHTML = rows.map((r, i) => `<tr>
      <td style="color:var(--muted)">${i + 1}</td>
      <td style="font-family:monospace;font-size:11px;color:var(--accent)">${r['Shipment ID'] || '—'}</td>
      <td style="font-family:monospace;font-size:11px">${r['AWB'] || '—'}</td>
      <td style="font-weight:600">${r['Customer'] || '—'}</td>
      <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
      <td>${this.pBadge(r['Payment Type'])}</td>
      <td>${this.sBadge(r['Shipment Status'])}</td>
      <td style="font-weight:700;color:var(--accent2)">${r[this.znField] || '—'}</td>
      <td>${this.fc(r['Total Amount'])}</td>
      <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
      <td>${this.tagChips(r['Tags'])}</td>
    </tr>`).join('');
  }

  private getCurrentZoneRows(): Row[] {
    const search = this.q<HTMLInputElement>('zn-search').value.toLowerCase();
    const zone = this.q<HTMLSelectElement>('zn-zone').value;
    const status = this.q<HTMLSelectElement>('zn-status').value;
    const pay = this.q<HTMLSelectElement>('zn-pay').value;
    return this.ALL.filter(r => {
      if (zone && r[this.znField] !== zone) return false;
      if (status && r['Shipment Status'] !== status) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
  }

  private downloadZoneOrders(): void { this.excelDownload(this.getCurrentZoneRows(), 'BANDB_Zone_City_Orders'); }

  private downloadZoneSummary(): void {
    const summary = this.znGroups.map(z => {
      const rows = this.ALL.filter(r => r[this.znField] === z);
      const pre = rows.filter(r => r['Payment Type'] === 'Prepaid').length;
      const cod = rows.filter(r => r['Payment Type'] === 'COD').length;
      const del = rows.filter(r => r['Shipment Status'] === 'Delivered').length;
      return {
        [this.znField]: z, 'Total Orders': rows.length, 'Prepaid': pre, 'COD': cod,
        'Delivered': del, 'In Transit': rows.filter(r => r['Shipment Status'] === 'In Transit').length,
        'Out for Delivery': rows.filter(r => r['Shipment Status'] === 'Out for Delivery').length,
        'Undelivered (NDR)': rows.filter(r => r['Shipment Status'] === 'Undelivered').length,
        'RTO In Transit': rows.filter(r => r['Shipment Status'] === 'RTO In Transit').length,
        'Delivery Rate %': rows.length ? ((del / rows.length) * 100).toFixed(1) : '0.0',
        'Revenue': rows.reduce((a, r) => a + (parseFloat(r['Total Amount']) || 0), 0),
        'COD Value': rows.reduce((a, r) => a + (parseFloat(r['COD Amount']) || 0), 0),
      };
    });
    this.excelDownload(summary, 'BANDB_Zone_City_Summary');
  }

  // ═══════════════════════════════════════════════════════
  // RTO — MONTH-WISE (Prepaid vs COD)
  // ═══════════════════════════════════════════════════════

  private monthKeyOf(r: Row): string | null {
    const d = this.parseDate(r['Pickup Date']);
    if (!d) return null;
    return `${d.getFullYear()}-${String(d.getMonth()).padStart(2, '0')}`;
  }
  private monthLabelOf(key: string): string {
    const [y, m] = key.split('-').map(Number);
    return new Date(y, m, 1).toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
  }
  private buildRtoMonthMap(): Record<string, any> {
    const monthMap: Record<string, any> = {};
    this.ALL.forEach(r => {
      const key = this.monthKeyOf(r);
      if (!key) return;
      if (!monthMap[key]) monthMap[key] = { total: 0, rtoTransit: 0, rtoTransitPre: 0, rtoTransitCod: 0, rtoDel: 0, rtoDelPre: 0, rtoDelCod: 0 };
      monthMap[key].total++;
      const st = r['Shipment Status'];
      if (st === 'RTO In Transit') {
        monthMap[key].rtoTransit++;
        if (r['Payment Type'] === 'Prepaid') monthMap[key].rtoTransitPre++;
        else if (r['Payment Type'] === 'COD') monthMap[key].rtoTransitCod++;
      } else if (st === 'RTO Delivered') {
        monthMap[key].rtoDel++;
        if (r['Payment Type'] === 'Prepaid') monthMap[key].rtoDelPre++;
        else if (r['Payment Type'] === 'COD') monthMap[key].rtoDelCod++;
      }
    });
    return monthMap;
  }

  private setupRtoPage(): void {
    const monthMap = this.buildRtoMonthMap();
    this.rtoMonthKeys = Object.keys(monthMap).sort((a, b) => a.localeCompare(b));

    const summary = this.rtoMonthKeys.map(key => {
      const m = monthMap[key];
      const rto = m.rtoTransit + m.rtoDel;
      const rtoPre = m.rtoTransitPre + m.rtoDelPre;
      const rtoCod = m.rtoTransitCod + m.rtoDelCod;
      return {
        key, label: this.monthLabelOf(key), total: m.total,
        rtoTransit: m.rtoTransit, rtoDel: m.rtoDel, rto,
        rtoPct: m.total ? ((rto / m.total) * 100).toFixed(2) : '0.00',
        rtoPre, rtoPrePct: m.total ? ((rtoPre / m.total) * 100).toFixed(1) : '0.0',
        rtoCod, rtoCodPct: m.total ? ((rtoCod / m.total) * 100).toFixed(1) : '0.0',
        rtoDelPre: m.rtoDelPre, rtoDelCod: m.rtoDelCod,
      };
    });
    const summaryDesc = [...summary].sort((a, b) => b.key.localeCompare(a.key));

    this.q('rto-month-tbody').innerHTML = summaryDesc.map(x => `<tr>
      <td style="font-weight:700;color:var(--accent2)">${x.label}</td>
      <td style="font-weight:900">${x.total}</td>
      <td style="color:var(--red);font-weight:700">${x.rto}</td>
      <td style="font-weight:700;color:${parseFloat(x.rtoPct) >= 10 ? 'var(--red)' : 'var(--green)'}">${x.rtoPct}%</td>
      <td style="color:var(--red)">${x.rtoTransit}</td>
      <td style="color:var(--pink)">${x.rtoDel}</td>
      <td style="color:var(--accent)">${x.rtoPre}</td>
      <td>${x.rtoPrePct}%</td>
      <td style="color:var(--orange)">${x.rtoCod}</td>
      <td>${x.rtoCodPct}%</td>
    </tr>`).join('');

    this.destroyChart('rto-month');
    this.chartMap['rto-month'] = new Chart(this.q('rto-month-chart'), { type: 'bar', data: { labels: summary.map(x => x.label), datasets: [
      { label: 'RTO — Prepaid', data: summary.map(x => x.rtoPre), backgroundColor: '#4f8ef7', borderRadius: 4 },
      { label: 'RTO — COD', data: summary.map(x => x.rtoCod), backgroundColor: '#f59e0b', borderRadius: 4 },
    ] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 10 } } } }, scales: { x: { stacked: true, ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { stacked: true, ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    this.destroyChart('rto-rate');
    this.chartMap['rto-rate'] = new Chart(this.q('rto-rate-chart'), { type: 'line', data: { labels: summary.map(x => x.label), datasets: [{ data: summary.map(x => parseFloat(x.rtoPct)), borderColor: '#ef4444', backgroundColor: 'rgba(239,68,68,.15)', fill: true, tension: .3, pointRadius: 3 }] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } }, scales: { x: { ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { ticks: { color: this.tickColor, callback: (v: any) => v + '%' }, grid: { color: this.gridColor } } } } });

    this.destroyChart('rto-status');
    this.chartMap['rto-status'] = new Chart(this.q('rto-status-chart'), { type: 'bar', data: { labels: summary.map(x => x.label), datasets: [
      { label: 'RTO In Transit', data: summary.map(x => x.rtoTransit), backgroundColor: '#ef4444', borderRadius: 4 },
      { label: 'RTO Delivered', data: summary.map(x => x.rtoDel), backgroundColor: '#ec4899', borderRadius: 4 },
    ] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 10 } } } }, scales: { x: { stacked: true, ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { stacked: true, ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    this.destroyChart('rto-delivered');
    this.chartMap['rto-delivered'] = new Chart(this.q('rto-delivered-chart'), { type: 'bar', data: { labels: summary.map(x => x.label), datasets: [
      { label: 'RTO Delivered — Prepaid', data: summary.map(x => x.rtoDelPre), backgroundColor: '#4f8ef7', borderRadius: 4 },
      { label: 'RTO Delivered — COD', data: summary.map(x => x.rtoDelCod), backgroundColor: '#f59e0b', borderRadius: 4 },
    ] }, options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { color: '#9ca3af', font: { size: 10 } } } }, scales: { x: { stacked: true, ticks: { color: this.tickColor, font: { size: 10 } }, grid: { display: false } }, y: { stacked: true, ticks: { color: this.tickColor }, grid: { color: this.gridColor } } } } });

    this.q<HTMLSelectElement>('rto-month-sel-filter').innerHTML = '<option value="">All Months</option>' + summaryDesc.map(x => `<option value="${x.key}">${x.label}</option>`).join('');

    this.filterRtoOrders();
  }

  private rtoDaysStuck(r: Row): number | null {
    return this.daysBetween(r['Latest Status Date'] || r['Pickup Date']);
  }

  private getCurrentRtoRows(): Row[] {
    const search = this.q<HTMLInputElement>('rto-search').value.toLowerCase();
    const monthKey = this.q<HTMLSelectElement>('rto-month-sel-filter').value;
    const status = this.q<HTMLSelectElement>('rto-status-sel').value;
    const pay = this.q<HTMLSelectElement>('rto-pay').value;
    const minStuck = parseInt(this.q<HTMLSelectElement>('rto-stuck-sel').value) || 0;
    return this.ALL.filter(r => {
      if (r['Shipment Status'] !== 'RTO In Transit' && r['Shipment Status'] !== 'RTO Delivered') return false;
      if (status && r['Shipment Status'] !== status) return false;
      if (monthKey && this.monthKeyOf(r) !== monthKey) return false;
      if (pay && r['Payment Type'] !== pay) return false;
      if (minStuck > 0) { const d = this.rtoDaysStuck(r); if (d === null || d < minStuck) return false; }
      if (search) { const h = (r['Customer'] || '') + (r['AWB'] || '') + (r['Order Number'] || '') + (r['Phone #'] || ''); if (!h.toLowerCase().includes(search)) return false; }
      return true;
    });
  }

  private filterRtoOrders(): void {
    let rows = this.getCurrentRtoRows();
    rows = [...rows].sort((a, b) => {
      const da = this.rtoDaysStuck(a), db = this.rtoDaysStuck(b);
      return (db === null ? -1 : db) - (da === null ? -1 : da);
    });
    this.q('rto-count').textContent = `${rows.length} orders`;
    const stuck7 = this.ALL.filter(r => r['Shipment Status'] === 'RTO In Transit' && (this.rtoDaysStuck(r) || 0) >= 7).length;
    this.q('rto-stuck-count').textContent = `⚠ ${stuck7} RTO In Transit stuck 7+ days`;
    this.q('rto-tbody').innerHTML = rows.map((r, i) => `<tr>
      <td style="color:var(--muted)">${i + 1}</td>
      <td style="font-family:monospace;font-size:11px;color:var(--accent)">${r['Shipment ID'] || '—'}</td>
      <td style="font-family:monospace;font-size:11px">${r['AWB'] || '—'}</td>
      <td style="font-weight:600">${r['Customer'] || '—'}</td>
      <td><a href="tel:${r['Phone #']}" style="color:var(--cyan);text-decoration:none">${r['Phone #'] || '—'}</a></td>
      <td>${this.pBadge(r['Payment Type'])}</td>
      <td>${this.sBadge(r['Shipment Status'])}</td>
      <td>${this.daysCell(this.rtoDaysStuck(r))}</td>
      <td>${r['Pickup Date'] || '—'}</td>
      <td style="font-weight:700;color:var(--accent2)">${r['Zone'] || '—'}</td>
      <td>${this.fc(r['Total Amount'])}</td>
      <td style="color:${(parseFloat(r['COD Amount']) || 0) > 0 ? 'var(--orange)' : 'var(--muted)'}">${(parseFloat(r['COD Amount']) || 0) > 0 ? this.fc(r['COD Amount']) : '—'}</td>
    </tr>`).join('');
  }

  private downloadRtoOrders(): void {
    const rows = this.getCurrentRtoRows().map(r => ({
      'Shipment ID': r['Shipment ID'], 'AWB': r['AWB'], 'Customer': r['Customer'], 'Phone': r['Phone #'],
      'Payment Type': r['Payment Type'], 'RTO Status': r['Shipment Status'],
      'Days Stuck': this.rtoDaysStuck(r) === null ? '—' : this.rtoDaysStuck(r),
      'Pickup Date': r['Pickup Date'], 'Latest Status Date': r['Latest Status Date'],
      'Zone': r['Zone'], 'Total Amount': r['Total Amount'], 'COD Amount': r['COD Amount'],
    }));
    this.excelDownload(rows, 'BANDB_RTO_Orders');
  }

  private downloadRtoMonthly(): void {
    const monthMap = this.buildRtoMonthMap();
    const keys = Object.keys(monthMap).sort((a, b) => b.localeCompare(a));
    const out = keys.map(key => {
      const m = monthMap[key];
      const rto = m.rtoTransit + m.rtoDel;
      const rtoPre = m.rtoTransitPre + m.rtoDelPre;
      const rtoCod = m.rtoTransitCod + m.rtoDelCod;
      return {
        'Month': this.monthLabelOf(key),
        'Total Orders': m.total,
        'RTO Total': rto,
        'RTO %': m.total ? ((rto / m.total) * 100).toFixed(2) : '0.00',
        'RTO In Transit': m.rtoTransit,
        'RTO Delivered': m.rtoDel,
        'RTO - Prepaid': rtoPre,
        'RTO Prepaid %': m.total ? ((rtoPre / m.total) * 100).toFixed(1) : '0.0',
        'RTO - COD': rtoCod,
        'RTO COD %': m.total ? ((rtoCod / m.total) * 100).toFixed(1) : '0.0',
        'RTO Delivered - Prepaid': m.rtoDelPre,
        'RTO Delivered - COD': m.rtoDelCod,
      };
    });
    this.excelDownload(out, 'BANDB_RTO_MonthWise_Summary');
  }

  // ═══════════════════════════════════════════════════════
  // INSIGHTS
  // ═══════════════════════════════════════════════════════

  private renderInsights(df: Row[], sm: Record<string, number>, tot: number, cod: number, tatBr: number, nd5: number, n: number): void {
    const del = sm['Delivered'] || 0;
    const rate = ((del / n) * 100).toFixed(1);
    const codPct = ((cod / tot) * 100).toFixed(1);
    const tatPct = ((tatBr / n) * 100).toFixed(1);
    const nd5Pct = ((nd5 / n) * 100).toFixed(1);
    const addressIssue = df.filter(r => r['Tags'] && String(r['Tags']).includes('Address Issue')).length;
    const escalation = df.filter(r => r['Tags'] && String(r['Tags']).includes('Escalation Exists')).length;
    const ndrRows = df.filter(r => r['Latest NDR Status'] && String(r['Latest NDR Status']).trim());
    const reschedule = ndrRows.filter(r => r['Latest NDR Status'] === 'Delivery Rescheduled').length;
    const unavail = ndrRows.filter(r => r['Latest NDR Status'] === 'Customer Unavailable').length;
    const zone_d = df.filter(r => r['Zone'] === 'D').length;
    const addressIssueRows = df.filter(r => r['Tags'] && String(r['Tags']).includes('Address Issue'));
    const addressIssueDelivered = addressIssueRows.filter(r => r['Shipment Status'] === 'Delivered').length;
    const addressIssuePending = addressIssueRows.length - addressIssueDelivered;
    void escalation;

    const { y: lmy, m: lmm } = this.getLastMonthRange();
    const lmMonthRows = df.filter(r => { const d = this.parseDate(r['Pickup Date']); return d && d.getFullYear() === lmy && d.getMonth() === lmm; });
    const lmMonthName = new Date(lmy, lmm, 1).toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
    const lmDelivered = lmMonthRows.filter(r => r['Shipment Status'] === 'Delivered').length;
    const lmRate = lmMonthRows.length ? ((lmDelivered / lmMonthRows.length) * 100).toFixed(1) : '0.0';
    const lmCod = lmMonthRows.filter(r => r['Payment Type'] === 'COD').length;
    const lmPre = lmMonthRows.filter(r => r['Payment Type'] === 'Prepaid').length;

    const codRows = df.filter(r => r['Payment Type'] === 'COD');
    const preRows = df.filter(r => r['Payment Type'] === 'Prepaid');
    const codDelRate = codRows.length ? ((codRows.filter(r => r['Shipment Status'] === 'Delivered').length / codRows.length) * 100).toFixed(1) : '0.0';
    const preDelRate = preRows.length ? ((preRows.filter(r => r['Shipment Status'] === 'Delivered').length / preRows.length) * 100).toFixed(1) : '0.0';
    const codRtoRate = codRows.length ? ((codRows.filter(r => r['Shipment Status'] === 'RTO In Transit').length / codRows.length) * 100).toFixed(1) : '0.0';

    const zField = this.getZoneField();
    const zGroupsForInsight = [...new Set(df.map(r => r[zField]).filter(Boolean))];
    const zStats = zGroupsForInsight.map(z => {
      const rows = df.filter(r => r[zField] === z);
      const d = rows.filter(r => r['Shipment Status'] === 'Delivered').length;
      return { z, total: rows.length, rate: rows.length ? (d / rows.length) * 100 : 0 };
    }).filter(x => x.total >= 5);
    const worstZone = zStats.length ? zStats.reduce((a, b) => a.rate < b.rate ? a : b) : null;
    const bestZone = zStats.length ? zStats.reduce((a, b) => a.rate > b.rate ? a : b) : null;

    const items: any[] = [
      {
        icon: '🎯',
        title: `Delivery Rate is ${rate}% — ${parseFloat(rate) >= 85 ? 'On Track' : 'Needs Attention'}`,
        body: `Out of ${n.toLocaleString()} shipments, ${del.toLocaleString()} are delivered. ${parseFloat(rate) < 85 ? '<span class="thought-pill p-warn">Below 85% threshold</span> — review NDR and RTO patterns to improve first-attempt delivery.' : '<span class="thought-pill p-ok">Healthy rate</span> — maintain by reducing NDR and TAT breach cases.'}`,
      },
      {
        icon: '⚠️',
        title: `${nd5} Shipments Stuck 5+ Days (${nd5Pct}% of total)`,
        body: `These orders have been picked up but not delivered. <span class="thought-pill p-warn">${nd5} orders</span> need immediate follow-up. Priority: COD orders first (cash recovery risk). Use the "Pending 5+ Days" tab to filter by Zone D and COD to create a targeted calling list.
        <br><button class="dl-btn" style="margin-top:8px" data-action="downloadPending">⬇ Download These ${nd5} Orders (Excel)</button>`,
      },
      {
        icon: '⏱',
        title: `TAT Breach: ${tatBr} Shipments Delayed (${tatPct}%)`,
        body: `Most TAT breaches are 1–3 day delays (In Transit or OFD). <span class="thought-pill p-info">Action:</span> Escalate to EKART for shipments breached by 3+ days. For OFD breaches, push for priority re-assignment to delivery executive. Download the TAT tab filtered by "Delayed by 3 days" and above.`,
      },
      {
        icon: '📍',
        title: `${addressIssue} Address Issue Tags — Largest Risk Bucket`,
        body: `<span class="thought-pill p-warn">${addressIssue} shipments</span> (${((addressIssue / n) * 100).toFixed(1)}%) are tagged "Address Issue." Of these, <span class="thought-pill p-ok">${addressIssueDelivered} are already Delivered</span> and <span class="thought-pill p-warn">${addressIssuePending} are still pending</span> (not yet delivered). <span class="thought-pill p-info">Recommendation:</span> Pull the AWBs for the pending ones, call customers to confirm PIN/landmark, and update Unicommerce before they auto-RTO.
        <br><button class="dl-btn" style="margin-top:8px" data-action="downloadAddressIssue">⬇ Download Address Issue Orders — with Delivered/Not Status (Excel)</button>`,
      },
      {
        icon: '🔁',
        title: `NDR: ${reschedule} Rescheduled, ${unavail} Customer Unavailable`,
        body: `Rescheduled NDRs are recoverable — these customers have agreed to a new attempt. <span class="thought-pill p-warn">Customer Unavailable (${unavail} cases)</span> are risk of RTO. Call them proactively before the next delivery attempt window. High COD+Unavailable = double risk.`,
      },
      {
        icon: '🗺',
        title: `Zone D is ${((zone_d / n) * 100).toFixed(1)}% of All Shipments`,
        body: `Zone D dominates the mix (${zone_d.toLocaleString()} shipments). This is expected for a national brand, but any SLA issue with EKART's Zone D hubs (Bengaluru, Hyderabad, Chennai) ripples across most orders. Monitor Zone D TAT separately in the TAT Breach tab.`,
      },
      {
        icon: '💰',
        title: `COD Exposure: ₹${(cod / 1000).toFixed(1)}K — ${codPct}% of Revenue`,
        body: `COD orders carry remittance lag and RTO risk. <span class="thought-pill p-info">Tip:</span> For customers with 2+ NDR attempts on COD orders, consider calling before re-dispatch to confirm intent. This reduces failed delivery cost and speeds up COD collection cycles.`,
      },
      {
        icon: '🚀',
        title: `Dispatch Consistency: Gap on Weekends Visible`,
        body: `Pickup data shows clustering around specific dates with no Saturday/Sunday dispatch visible in this export. If weekend orders are being batched to Monday pickup, EDD commitments to customers may be missed — especially for Zone D (4–5 day SLA). Consider pushing Sunday pickup for at least premium/escalation orders.`,
      },
      {
        icon: '📅',
        title: `${lmMonthName}: ${lmMonthRows.length} Orders, ${lmRate}% Delivered`,
        body: `Last month had <span class="thought-pill p-info">${lmPre} Prepaid</span> and <span class="thought-pill p-info">${lmCod} COD</span> orders, with a delivery rate of ${lmRate}%. Check the "Last Month" tab for the full status × payment breakdown and a downloadable order list.
        <br><button class="dl-btn" style="margin-top:8px" data-action="switchTab" data-arg="lastmonth">📅 Open Last Month Tab</button>`,
      },
      {
        icon: '💳',
        title: `Prepaid Delivers ${preDelRate}% vs COD ${codDelRate}%`,
        body: `Prepaid orders are delivering ${parseFloat(preDelRate) >= parseFloat(codDelRate) ? 'better than' : 'worse than'} COD orders. COD also carries a <span class="thought-pill p-warn">${codRtoRate}% RTO rate</span>. <span class="thought-pill p-info">Action:</span> For repeat COD-RTO customers, consider nudging toward prepaid (small discount) to cut return logistics cost.`,
      },
      {
        icon: '🏆',
        title: worstZone ? `${zField === 'Zone' ? 'Zone' : zField} Gap: "${bestZone!.z}" ${bestZone!.rate.toFixed(1)}% vs "${worstZone.z}" ${worstZone.rate.toFixed(1)}%` : 'Not enough zone data for comparison',
        body: worstZone ? `Best-performing ${zField === 'Zone' ? 'zone' : zField.toLowerCase()} is <span class="thought-pill p-ok">${bestZone!.z} (${bestZone!.rate.toFixed(1)}% delivered)</span>, weakest is <span class="thought-pill p-warn">${worstZone.z} (${worstZone.rate.toFixed(1)}% delivered)</span>. See the "Zone / City" tab for the full breakdown and a downloadable summary.
        <br><button class="dl-btn" style="margin-top:8px" data-action="switchTab" data-arg="zone">🗺 Open Zone / City Tab</button>` : '',
      },
    ];

    this.q('thoughts-content').innerHTML = `<h3>💡 Data-Driven Insights — BANDB Shipment Analysis</h3>` +
      items.map((x, i) => `<div class="thought-item">
        <div class="thought-num">${i + 1}</div>
        <div class="thought-content">
          <h4>${x.icon} ${x.title}</h4>
          <p>${x.body}</p>
        </div>
      </div>`).join('');
  }

  private downloadAddressIssue(): void {
    const rows = this.ALL.filter(r => r['Tags'] && String(r['Tags']).includes('Address Issue')).map(r => ({
      'Shipment ID': r['Shipment ID'],
      'AWB': r['AWB'],
      'Order Number': r['Order Number'],
      'Customer': r['Customer'],
      'Phone #': r['Phone #'],
      'Payment Type': r['Payment Type'],
      'Delivered / Not Delivered': r['Shipment Status'] === 'Delivered' ? 'Delivered' : 'Not Delivered',
      'Shipment Status': r['Shipment Status'],
      'Pickup Date': r['Pickup Date'],
      'Delivery Date': r['Delivery Date'] || '',
      'Zone': r['Zone'],
      'Total Amount': r['Total Amount'],
      'COD Amount': r['COD Amount'],
      'Attempts': r['Attempts'],
      'Latest NDR Status': r['Latest NDR Status'],
      'Tags': r['Tags'],
    }));
    this.excelDownload(rows, 'BANDB_Address_Issue_Orders');
  }

  // ═══════════════════════════════════════════════════════
  // SHARED COLUMN-RESOLUTION HELPERS (BlueDart / India Post)
  // ═══════════════════════════════════════════════════════

  private resolveCol(headers: string[], aliases: string[]): string | null {
    const norm = (h: any) => String(h || '').toLowerCase().trim().replace(/[_\-\/]+/g, ' ').replace(/\s+/g, ' ').trim();
    for (const a of aliases) {
      const hit = headers.find(h => norm(h) === norm(a));
      if (hit) return hit;
    }
    for (const a of aliases) {
      const hit = headers.find(h => norm(h).includes(norm(a)));
      if (hit) return hit;
    }
    return null;
  }
  private pick(row: Row, col: string | null): string {
    return col ? (row[col] !== undefined && row[col] !== null ? String(row[col]).trim() : '') : '';
  }
  private topN(counts: Record<string, number>, n: number): [string, number][] {
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, n);
  }
  private rankedTbody(pairs: [string, number][], total: number): string {
    return pairs.map(([k, v], i) => `<tr><td><strong>#${i + 1}</strong></td><td>${k || '—'}</td><td><strong>${v}</strong></td><td>${total ? ((v / total) * 100).toFixed(2) : '0.00'}%</td></tr>`).join('');
  }
  private fillSelect(sel: HTMLSelectElement, values: any[]): void {
    const cur = sel.value;
    const firstOption = sel.querySelector('option');
    sel.innerHTML = '<option value="">' + (firstOption?.textContent || '') + '</option>' +
      [...new Set(values)].filter(Boolean).sort().map(v => `<option value="${v}">${v}</option>`).join('');
    if ([...sel.options].some(o => o.value === cur)) sel.value = cur;
  }

  // ═══════════════════════════════════════════════════════
  // BLUEDART MODULE
  // ═══════════════════════════════════════════════════════

  private switchTabBD(tab: string, btnEl?: HTMLElement): void {
    this.qa<HTMLElement>('#app-bluedart .tab-btn').forEach(b => b.classList.remove('active'));
    this.qa<HTMLElement>('#app-bluedart .page').forEach(p => p.classList.remove('active'));
    btnEl?.classList.add('active');
    this.q('bd-page-' + tab).classList.add('active');
  }

  private processBluedart(): void {
    const headers = Object.keys(this.BD_ALL[0] || {});
    this.BD_COLS = {
      awb: this.resolveCol(headers, ['waybill no', 'waybill number', 'way bill no', 'waybill', 'awb no', 'awb', 'airway bill', 'tracking number', 'docket number']),
      status: this.resolveCol(headers, ['status description', 'shipment status', 'delivery status', 'current status', 'status']),
      origin: this.resolveCol(headers, ['origin description', 'origin city', 'shipper city', 'from city', 'origin area', 'origin']),
      dest: this.resolveCol(headers, ['destination description', 'destination city', 'consignee city', 'to city', 'destination']),
      weight: this.resolveCol(headers, ['weight (kg)', 'weight', 'actual weight', 'chargeable weight']),
      sender: this.resolveCol(headers, ['sender name', 'shipper name', 'from name', 'sender']),
      receiver: this.resolveCol(headers, ['consignee name', 'receiver name', 'to name', 'customer name', 'receiver']),
      bookDate: this.resolveCol(headers, ['p u date', 'pu date', 'pickup date', 'booking date', 'origin date', 'dispatch date']),
      delDate: this.resolveCol(headers, ['exp delivery date', 'delivery date', 'delivered date', 'pod date']),
      product: this.resolveCol(headers, ['product type', 'service type', 'product', 'mode']),
    };
    this.BD_ROWS = this.BD_ALL.map(r => ({
      awb: this.pick(r, this.BD_COLS['awb']),
      status: this.pick(r, this.BD_COLS['status']) || 'Unknown',
      origin: this.pick(r, this.BD_COLS['origin']) || 'Unknown',
      dest: this.pick(r, this.BD_COLS['dest']) || 'Unknown',
      weight: parseFloat(this.pick(r, this.BD_COLS['weight'])) || 0,
      sender: this.pick(r, this.BD_COLS['sender']) || 'Unknown',
      receiver: this.pick(r, this.BD_COLS['receiver']),
      bookDate: this.pick(r, this.BD_COLS['bookDate']),
      delDate: this.pick(r, this.BD_COLS['delDate']),
      product: this.pick(r, this.BD_COLS['product']) || 'Standard',
    }));
    this.q('bd-rpt-date').textContent = 'Report generated ' + this.today.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
    this.renderBDKpis();
    this.renderBDCharts();
    this.renderBDRankedTables();
    this.renderBDTableFilters();
    this.renderBDTable();
    this.renderBDInsights();
  }

  private renderBDKpis(): void {
    const n = this.BD_ROWS.length;
    const delivered = this.BD_ROWS.filter(r => /delivered/i.test(r.status)).length;
    const rto = this.BD_ROWS.filter(r => /rto|return/i.test(r.status)).length;
    const pending = n - delivered - rto;
    const totalWeight = this.BD_ROWS.reduce((s, r) => s + r.weight, 0);
    const rate = n ? ((delivered / n) * 100).toFixed(1) : '0.0';
    const origins = new Set(this.BD_ROWS.map(r => r.origin)).size;
    const dests = new Set(this.BD_ROWS.map(r => r.dest)).size;
    const kpis: any[] = [
      { label: 'Total Shipments', val: n.toLocaleString(), icon: '📦', ka: 'var(--accent)' },
      { label: 'Delivered', val: delivered.toLocaleString(), sub: rate + '% delivery rate', icon: '✅', ka: 'var(--green)' },
      { label: 'RTO / Returned', val: rto.toLocaleString(), icon: '↩️', ka: 'var(--red)' },
      { label: 'In Transit / Pending', val: pending.toLocaleString(), icon: '🚚', ka: 'var(--orange)' },
      { label: 'Total Weight', val: totalWeight.toFixed(1) + ' kg', icon: '⚖️', ka: 'var(--cyan)' },
      { label: 'Unique Origins', val: origins, icon: '📍', ka: 'var(--accent2)' },
      { label: 'Unique Destinations', val: dests, icon: '🎯', ka: 'var(--pink)' },
    ];
    this.q('bd-kpi-grid').innerHTML = kpis.map(k => `<div class="kpi" style="--ka:${k.ka}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-val">${k.val}</div>
      ${k.sub ? `<div class="kpi-sub">${k.sub}</div>` : ''}
    </div>`).join('');
  }

  private renderBDCharts(): void {
    Object.values(this.bdChartMap).forEach((c: any) => c && c.destroy());
    const statusCounts: Record<string, number> = {};
    const weightBuckets: Record<string, number> = { 'Below 500g': 0, '500g - 1kg': 0, '1kg - 2kg': 0, '2kg - 5kg': 0, 'Above 5kg': 0 };
    const originCounts: Record<string, number> = {}, destCounts: Record<string, number> = {}, senderCounts: Record<string, number> = {}, productCounts: Record<string, number> = {};
    this.BD_ROWS.forEach(r => {
      statusCounts[r.status] = (statusCounts[r.status] || 0) + 1;
      const w = r.weight;
      if (w < 0.5) weightBuckets['Below 500g']++;
      else if (w < 1) weightBuckets['500g - 1kg']++;
      else if (w < 2) weightBuckets['1kg - 2kg']++;
      else if (w < 5) weightBuckets['2kg - 5kg']++;
      else if (w > 0) weightBuckets['Above 5kg']++;
      originCounts[r.origin] = (originCounts[r.origin] || 0) + 1;
      destCounts[r.dest] = (destCounts[r.dest] || 0) + 1;
      senderCounts[r.sender] = (senderCounts[r.sender] || 0) + 1;
      productCounts[r.product] = (productCounts[r.product] || 0) + 1;
    });
    const palette = ['#4f8ef7', '#7c5cfc', '#22c55e', '#f59e0b', '#ef4444', '#06b6d4', '#ec4899', '#a855f7'];
    const cfg = { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#9ca3af', font: { size: 11 } } } } };
    this.bdChartMap['status'] = new Chart(this.q('bd-status-chart'), { type: 'doughnut', data: { labels: Object.keys(statusCounts), datasets: [{ data: Object.values(statusCounts), backgroundColor: palette, borderColor: '#13161f', borderWidth: 3 }] }, options: cfg });
    this.bdChartMap['weight'] = new Chart(this.q('bd-weight-chart'), { type: 'bar', data: { labels: Object.keys(weightBuckets), datasets: [{ label: 'Shipments', data: Object.values(weightBuckets), backgroundColor: 'rgba(79,142,247,.8)', borderRadius: 6 }] }, options: { ...cfg, scales: { y: { beginAtZero: true, ticks: { color: '#9ca3af' } }, x: { ticks: { color: '#9ca3af' } } } } });
    this.bdChartMap['product'] = new Chart(this.q('bd-product-chart'), { type: 'pie', data: { labels: Object.keys(productCounts), datasets: [{ data: Object.values(productCounts), backgroundColor: palette, borderColor: '#13161f', borderWidth: 3 }] }, options: cfg });
    const topO = this.topN(originCounts, 10), topD = this.topN(destCounts, 10), topS = this.topN(senderCounts, 10);
    this.bdChartMap['origin'] = new Chart(this.q('bd-origin-chart'), { type: 'bar', data: { labels: topO.map(x => x[0]), datasets: [{ label: 'Shipments', data: topO.map(x => x[1]), backgroundColor: 'rgba(34,197,94,.8)', borderRadius: 6 }] }, options: { ...cfg, indexAxis: 'y', scales: { x: { beginAtZero: true, ticks: { color: '#9ca3af' } }, y: { ticks: { color: '#9ca3af' } } } } });
    this.bdChartMap['dest'] = new Chart(this.q('bd-dest-chart'), { type: 'bar', data: { labels: topD.map(x => x[0]), datasets: [{ label: 'Shipments', data: topD.map(x => x[1]), backgroundColor: 'rgba(245,158,11,.8)', borderRadius: 6 }] }, options: { ...cfg, indexAxis: 'y', scales: { x: { beginAtZero: true, ticks: { color: '#9ca3af' } }, y: { ticks: { color: '#9ca3af' } } } } });
    this.bdChartMap['sender'] = new Chart(this.q('bd-sender-chart'), { type: 'bar', data: { labels: topS.map(x => x[0]), datasets: [{ label: 'Shipments', data: topS.map(x => x[1]), backgroundColor: 'rgba(124,92,252,.8)', borderRadius: 6 }] }, options: { ...cfg, indexAxis: 'y', scales: { x: { beginAtZero: true, ticks: { color: '#9ca3af' } }, y: { ticks: { color: '#9ca3af' } } } } });
    this.BD_CACHE = { statusCounts, originCounts, destCounts, senderCounts, weightBuckets, productCounts };
  }

  private renderBDRankedTables(): void {
    const n = this.BD_ROWS.length;
    this.q('bd-origins-tbody').innerHTML = this.rankedTbody(this.topN(this.BD_CACHE.originCounts, 15), n);
    this.q('bd-dest-tbody').innerHTML = this.rankedTbody(this.topN(this.BD_CACHE.destCounts, 15), n);
  }

  private renderBDTableFilters(): void {
    this.fillSelect(this.q<HTMLSelectElement>('bd-status-filter'), this.BD_ROWS.map(r => r.status));
    this.fillSelect(this.q<HTMLSelectElement>('bd-origin-filter'), this.BD_ROWS.map(r => r.origin));
    this.fillSelect(this.q<HTMLSelectElement>('bd-dest-filter'), this.BD_ROWS.map(r => r.dest));
  }

  private renderBDTable(): void {
    const q = (this.q<HTMLInputElement>('bd-search').value || '').toLowerCase();
    const st = this.q<HTMLSelectElement>('bd-status-filter').value, or = this.q<HTMLSelectElement>('bd-origin-filter').value, de = this.q<HTMLSelectElement>('bd-dest-filter').value;
    const rows = this.BD_ROWS.filter(r => {
      if (st && r.status !== st) return false;
      if (or && r.origin !== or) return false;
      if (de && r.dest !== de) return false;
      if (q && !(`${r.awb} ${r.sender} ${r.receiver} ${r.origin} ${r.dest}`.toLowerCase().includes(q))) return false;
      return true;
    });
    this.q('bd-count').textContent = `${rows.length.toLocaleString()} of ${this.BD_ROWS.length.toLocaleString()} shipments`;
    this.q('bd-tbody').innerHTML = rows.slice(0, 1000).map((r, i) => `<tr>
      <td>${i + 1}</td><td>${r.awb || '—'}</td><td>${this.sBadgeBD(r.status)}</td><td>${r.origin}</td><td>${r.dest}</td>
      <td>${r.weight ? r.weight.toFixed(2) + ' kg' : '—'}</td><td>${r.sender}</td><td>${r.receiver || '—'}</td>
      <td>${r.bookDate || '—'}</td><td>${r.delDate || '—'}</td><td>${r.product}</td>
    </tr>`).join('');
    this.BD_FILTERED = rows;
  }

  private sBadgeBD(s: string): string {
    const cls = /delivered/i.test(s) ? 'b-del' : /rto|return/i.test(s) ? 'b-rto' : /transit/i.test(s) ? 'b-it' : /out for delivery/i.test(s) ? 'b-ofd' : 'b-it';
    return `<span class="badge ${cls}">${s || '—'}</span>`;
  }

  private downloadBD(): void {
    const rows = (this.BD_FILTERED.length ? this.BD_FILTERED : this.BD_ROWS).map(r => ({
      'AWB / Waybill': r.awb, 'Status': r.status, 'Origin': r.origin, 'Destination': r.dest,
      'Weight (kg)': r.weight, 'Sender': r.sender, 'Receiver': r.receiver,
      'Booking Date': r.bookDate, 'Delivery Date': r.delDate, 'Product': r.product,
    }));
    this.excelDownload(rows, 'BlueDart_Shipments');
  }

  private renderBDInsights(): void {
    const n = this.BD_ROWS.length;
    const delivered = this.BD_ROWS.filter(r => /delivered/i.test(r.status)).length;
    const rate = n ? ((delivered / n) * 100).toFixed(1) : '0.0';
    const rto = this.BD_ROWS.filter(r => /rto|return/i.test(r.status)).length;
    const topOrigin = this.topN(this.BD_CACHE.originCounts, 1)[0];
    const topDest = this.topN(this.BD_CACHE.destCounts, 1)[0];
    const topSender = this.topN(this.BD_CACHE.senderCounts, 1)[0];
    const avgWeight = n ? (this.BD_ROWS.reduce((s, r) => s + r.weight, 0) / n).toFixed(2) : '0.00';
    const items: any[] = [
      { icon: '🎯', title: `Delivery Rate is ${rate}% (${delivered.toLocaleString()} of ${n.toLocaleString()})`,
        body: `${parseFloat(rate) >= 85 ? '<span class="thought-pill p-ok">Healthy rate</span> — delivery performance is on track.' : '<span class="thought-pill p-warn">Below 85% threshold</span> — investigate undelivered and RTO shipments for recoverable orders.'}` },
      { icon: '↩️', title: `${rto.toLocaleString()} RTO / Returned Shipments`,
        body: `RTO shipments represent ${n ? ((rto / n) * 100).toFixed(1) : '0.0'}% of volume. <span class="thought-pill p-info">Action:</span> Review addresses and customer contact-ability for these shipments to reduce future returns.` },
      { icon: '📍', title: topOrigin ? `Top Origin: "${topOrigin[0]}" (${topOrigin[1]} shipments)` : 'No origin data available',
        body: topOrigin ? `This location accounts for ${((topOrigin[1] / n) * 100).toFixed(1)}% of all dispatches. Any operational issue here has outsized impact.` : '' },
      { icon: '🎯', title: topDest ? `Top Destination: "${topDest[0]}" (${topDest[1]} shipments)` : 'No destination data available',
        body: topDest ? `This destination receives ${((topDest[1] / n) * 100).toFixed(1)}% of shipments — consider dedicated last-mile follow-up here.` : '' },
      { icon: '👤', title: topSender ? `Highest-Volume Sender: "${topSender[0]}" (${topSender[1]} shipments)` : 'No sender data available',
        body: topSender ? `Concentration risk: keep this sender's shipments on a priority monitoring list.` : '' },
      { icon: '⚖️', title: `Average Shipment Weight: ${avgWeight} kg`,
        body: `Use this baseline to flag unusually heavy or light shipments for cost or damage review.` },
    ];
    this.q('bd-thoughts-content').innerHTML = `<h3>💡 Data-Driven Insights — BlueDart Analysis</h3>` +
      items.filter(x => x.body).map((x, i) => `<div class="thought-item">
        <div class="thought-num">${i + 1}</div>
        <div class="thought-content"><h4>${x.icon} ${x.title}</h4><p>${x.body}</p></div>
      </div>`).join('');
  }

  // ═══════════════════════════════════════════════════════
  // INDIA POST MODULE
  // ═══════════════════════════════════════════════════════

  private switchTabIP(tab: string, btnEl?: HTMLElement): void {
    this.qa<HTMLElement>('#app-indiapost .tab-btn').forEach(b => b.classList.remove('active'));
    this.qa<HTMLElement>('#app-indiapost .page').forEach(p => p.classList.remove('active'));
    btnEl?.classList.add('active');
    this.q('ip-page-' + tab).classList.add('active');
  }

  private processIndiaPost(): void {
    const headers = Object.keys(this.IP_ALL[0] || {});
    this.IP_COLS = {
      article: this.resolveCol(headers, ['article number', 'article no', 'consignment number', 'speed post number', 'tracking number', 'tracking id']),
      status: this.resolveCol(headers, ['event description', 'delivery status', 'current status', 'status']),
      bookOffice: this.resolveCol(headers, ['booking office name', 'booking office', 'origin office', 'from office']),
      destOffice: this.resolveCol(headers, ['destination office name', 'destination office', 'delivery office', 'to office']),
      weight: this.resolveCol(headers, ['weight (g)', 'weight (kg)', 'weight']),
      artType: this.resolveCol(headers, ['article type', 'item type', 'service type']),
      user: this.resolveCol(headers, ['user name', 'booked by', 'booking user', 'counter', 'employee', 'user']),
      bookDate: this.resolveCol(headers, ['booking date time', 'booking date', 'date of booking']),
      delDate: this.resolveCol(headers, ['delivery date', 'date of delivery', 'delivered date']),
      receiver: this.resolveCol(headers, ['receiver name', 'consignee name']),
      payMode: this.resolveCol(headers, ['payment mode description', 'payment mode']),
      amount: this.resolveCol(headers, ['base amount', 'tarrif', 'amount']),
    };
    // Detect weight unit: India Post exports are typically in grams (values like 1247) — auto-scale to kg
    const sampleWeights = this.IP_ALL.slice(0, 20).map(r => parseFloat(this.pick(r, this.IP_COLS['weight'])) || 0).filter(w => w > 0);
    const avgSample = sampleWeights.length ? sampleWeights.reduce((a, b) => a + b, 0) / sampleWeights.length : 0;
    const weightDivisor = avgSample > 50 ? 1000 : 1; // >50 strongly implies grams, not kg
    this.IP_ROWS = this.IP_ALL.map(r => ({
      article: this.pick(r, this.IP_COLS['article']),
      status: this.pick(r, this.IP_COLS['status']) || 'Unknown',
      bookOffice: this.pick(r, this.IP_COLS['bookOffice']) || 'Unknown',
      destOffice: this.pick(r, this.IP_COLS['destOffice']) || 'Unknown',
      weight: (parseFloat(this.pick(r, this.IP_COLS['weight'])) || 0) / weightDivisor,
      artType: this.pick(r, this.IP_COLS['artType']) || 'SP_INLAND_PARCEL',
      user: this.pick(r, this.IP_COLS['user']) || 'Unknown',
      bookDate: this.pick(r, this.IP_COLS['bookDate']),
      delDate: this.pick(r, this.IP_COLS['delDate']),
      receiver: this.pick(r, this.IP_COLS['receiver']),
      payMode: this.pick(r, this.IP_COLS['payMode']),
      amount: parseFloat(this.pick(r, this.IP_COLS['amount'])) || 0,
    }));
    this.q('ip-rpt-date').textContent = 'Report generated ' + this.today.toLocaleDateString('en-IN', { day: 'numeric', month: 'short', year: 'numeric' });
    this.renderIPKpis();
    this.renderIPCharts();
    this.renderIPRankedTables();
    this.renderIPTableFilters();
    this.renderIPTable();
    this.renderIPInsights();
  }

  private renderIPKpis(): void {
    const n = this.IP_ROWS.length;
    const delivered = this.IP_ROWS.filter(r => /delivered/i.test(r.status)).length;
    const dispatched = this.IP_ROWS.filter(r => /dispatch/i.test(r.status)).length;
    const pending = n - delivered;
    const rate = n ? ((delivered / n) * 100).toFixed(1) : '0.0';
    const offices = new Set(this.IP_ROWS.map(r => r.destOffice)).size;
    const users = new Set(this.IP_ROWS.map(r => r.user)).size;
    const totalWeight = this.IP_ROWS.reduce((s, r) => s + r.weight, 0);
    const totalValue = this.IP_ROWS.reduce((s, r) => s + r.amount, 0);
    const kpis: any[] = [
      { label: 'Total Articles', val: n.toLocaleString(), icon: '📮', ka: 'var(--accent)' },
      { label: 'Delivered', val: delivered.toLocaleString(), sub: rate + '% delivery rate', icon: '✅', ka: 'var(--green)' },
      { label: 'Dispatched', val: dispatched.toLocaleString(), icon: '🚚', ka: 'var(--orange)' },
      { label: 'Pending / In-Process', val: pending.toLocaleString(), icon: '⏳', ka: 'var(--red)' },
      { label: 'Destination Offices', val: offices, icon: '🏢', ka: 'var(--cyan)' },
      { label: 'Booking Users', val: users, icon: '👤', ka: 'var(--accent2)' },
      { label: 'Total Weight', val: totalWeight.toFixed(1) + ' kg', icon: '⚖️', ka: 'var(--pink)' },
      { label: 'Total Postage Value', val: '₹' + totalValue.toLocaleString('en-IN', { maximumFractionDigits: 0 }), icon: '💰', ka: 'var(--green)' },
    ];
    this.q('ip-kpi-grid').innerHTML = kpis.map(k => `<div class="kpi" style="--ka:${k.ka}">
      <div class="kpi-icon">${k.icon}</div>
      <div class="kpi-label">${k.label}</div>
      <div class="kpi-val">${k.val}</div>
      ${k.sub ? `<div class="kpi-sub">${k.sub}</div>` : ''}
    </div>`).join('');
  }

  private renderIPCharts(): void {
    Object.values(this.ipChartMap).forEach((c: any) => c && c.destroy());
    const statusCounts: Record<string, number> = {};
    const weightBuckets: Record<string, number> = { 'Below 500g': 0, '500g - 1kg': 0, '1kg - 2kg': 0, '2kg - 5kg': 0, 'Above 5kg': 0 };
    const officeCounts: Record<string, number> = {}, userCounts: Record<string, number> = {}, typeCounts: Record<string, number> = {};
    this.IP_ROWS.forEach(r => {
      statusCounts[r.status] = (statusCounts[r.status] || 0) + 1;
      const w = r.weight;
      if (w < 0.5) weightBuckets['Below 500g']++;
      else if (w < 1) weightBuckets['500g - 1kg']++;
      else if (w < 2) weightBuckets['1kg - 2kg']++;
      else if (w < 5) weightBuckets['2kg - 5kg']++;
      else if (w > 0) weightBuckets['Above 5kg']++;
      officeCounts[r.destOffice] = (officeCounts[r.destOffice] || 0) + 1;
      userCounts[r.user] = (userCounts[r.user] || 0) + 1;
      typeCounts[r.artType] = (typeCounts[r.artType] || 0) + 1;
    });
    const palette = ['#4f8ef7', '#7c5cfc', '#22c55e', '#f59e0b', '#ef4444', '#06b6d4', '#ec4899', '#a855f7'];
    const cfg = { responsive: true, maintainAspectRatio: false, plugins: { legend: { labels: { color: '#9ca3af', font: { size: 11 } } } } };
    this.ipChartMap['status'] = new Chart(this.q('ip-status-chart'), { type: 'doughnut', data: { labels: Object.keys(statusCounts), datasets: [{ data: Object.values(statusCounts), backgroundColor: palette, borderColor: '#13161f', borderWidth: 3 }] }, options: cfg });
    this.ipChartMap['weight'] = new Chart(this.q('ip-weight-chart'), { type: 'bar', data: { labels: Object.keys(weightBuckets), datasets: [{ label: 'Articles', data: Object.values(weightBuckets), backgroundColor: 'rgba(79,142,247,.8)', borderRadius: 6 }] }, options: { ...cfg, scales: { y: { beginAtZero: true, ticks: { color: '#9ca3af' } }, x: { ticks: { color: '#9ca3af' } } } } });
    this.ipChartMap['article'] = new Chart(this.q('ip-article-chart'), { type: 'pie', data: { labels: Object.keys(typeCounts), datasets: [{ data: Object.values(typeCounts), backgroundColor: palette, borderColor: '#13161f', borderWidth: 3 }] }, options: cfg });
    const topOff = this.topN(officeCounts, 15), topUsr = this.topN(userCounts, 10);
    this.ipChartMap['dest'] = new Chart(this.q('ip-dest-chart'), { type: 'bar', data: { labels: topOff.map(x => x[0]), datasets: [{ label: 'Articles', data: topOff.map(x => x[1]), backgroundColor: 'rgba(245,158,11,.8)', borderRadius: 6 }] }, options: { ...cfg, indexAxis: 'y', scales: { x: { beginAtZero: true, ticks: { color: '#9ca3af' } }, y: { ticks: { color: '#9ca3af' } } } } });
    this.ipChartMap['user'] = new Chart(this.q('ip-user-chart'), { type: 'bar', data: { labels: topUsr.map(x => x[0]), datasets: [{ label: 'Articles', data: topUsr.map(x => x[1]), backgroundColor: 'rgba(34,197,94,.8)', borderRadius: 6 }] }, options: { ...cfg, indexAxis: 'y', scales: { x: { beginAtZero: true, ticks: { color: '#9ca3af' } }, y: { ticks: { color: '#9ca3af' } } } } });
    this.IP_CACHE = { statusCounts, officeCounts, userCounts, typeCounts, weightBuckets };
  }

  private renderIPRankedTables(): void {
    const n = this.IP_ROWS.length;
    this.q('ip-dest-tbody').innerHTML = this.rankedTbody(this.topN(this.IP_CACHE.officeCounts, 15), n);
    this.q('ip-user-tbody').innerHTML = this.rankedTbody(this.topN(this.IP_CACHE.userCounts, 15), n);
  }

  private renderIPTableFilters(): void {
    this.fillSelect(this.q<HTMLSelectElement>('ip-status-filter'), this.IP_ROWS.map(r => r.status));
    this.fillSelect(this.q<HTMLSelectElement>('ip-office-filter'), this.IP_ROWS.map(r => r.destOffice));
  }

  private renderIPTable(): void {
    const q = (this.q<HTMLInputElement>('ip-search').value || '').toLowerCase();
    const st = this.q<HTMLSelectElement>('ip-status-filter').value, off = this.q<HTMLSelectElement>('ip-office-filter').value;
    const rows = this.IP_ROWS.filter(r => {
      if (st && r.status !== st) return false;
      if (off && r.destOffice !== off) return false;
      if (q && !(`${r.article} ${r.bookOffice} ${r.destOffice} ${r.user}`.toLowerCase().includes(q))) return false;
      return true;
    });
    this.q('ip-count').textContent = `${rows.length.toLocaleString()} of ${this.IP_ROWS.length.toLocaleString()} articles`;
    this.q('ip-tbody').innerHTML = rows.slice(0, 1000).map((r, i) => `<tr>
      <td>${i + 1}</td><td>${r.article || '—'}</td><td>${this.sBadgeIP(r.status)}</td><td>${r.bookOffice}</td><td>${r.destOffice}</td>
      <td>${r.weight ? r.weight.toFixed(2) + ' kg' : '—'}</td><td>${r.artType}</td><td>${r.user}</td>
      <td>${r.bookDate || '—'}</td><td>${r.delDate || '—'}</td>
    </tr>`).join('');
    this.IP_FILTERED = rows;
  }

  private sBadgeIP(s: string): string {
    const cls = /delivered/i.test(s) ? 'b-del' : /hold/i.test(s) ? 'b-ndr' : /dispatch|transit/i.test(s) ? 'b-it' : /out for delivery/i.test(s) ? 'b-ofd' : 'b-it';
    return `<span class="badge ${cls}">${s || '—'}</span>`;
  }

  private downloadIP(): void {
    const rows = (this.IP_FILTERED.length ? this.IP_FILTERED : this.IP_ROWS).map(r => ({
      'Article Number': r.article, 'Status': r.status, 'Booking Office': r.bookOffice, 'Destination Office': r.destOffice,
      'Weight (kg)': r.weight.toFixed(3), 'Article Type': r.artType, 'Booked By': r.user, 'Receiver': r.receiver,
      'Payment Mode': r.payMode, 'Amount': r.amount, 'Booking Date': r.bookDate, 'Delivery Date': r.delDate,
    }));
    this.excelDownload(rows, 'IndiaPost_Articles');
  }

  private renderIPInsights(): void {
    const n = this.IP_ROWS.length;
    const delivered = this.IP_ROWS.filter(r => /delivered/i.test(r.status)).length;
    const rate = n ? ((delivered / n) * 100).toFixed(1) : '0.0';
    const onHold = this.IP_ROWS.filter(r => /hold/i.test(r.status)).length;
    const topOffice = this.topN(this.IP_CACHE.officeCounts, 1)[0];
    const topUser = this.topN(this.IP_CACHE.userCounts, 1)[0];
    const avgWeight = n ? (this.IP_ROWS.reduce((s, r) => s + r.weight, 0) / n).toFixed(2) : '0.00';
    const items: any[] = [
      { icon: '🎯', title: `Delivery Rate is ${rate}% (${delivered.toLocaleString()} of ${n.toLocaleString()})`,
        body: `${parseFloat(rate) >= 85 ? '<span class="thought-pill p-ok">Healthy rate</span> — most articles are reaching addressees.' : '<span class="thought-pill p-warn">Below 85% threshold</span> — review dispatched and on-hold articles for delays.'}` },
      { icon: '⚠️', title: `${onHold.toLocaleString()} Articles On Hold`,
        body: onHold ? `<span class="thought-pill p-warn">On Hold</span> articles need manual follow-up with the destination office to release for delivery.` : 'No articles currently on hold — good operational health.' },
      { icon: '🏢', title: topOffice ? `Busiest Destination Office: "${topOffice[0]}" (${topOffice[1]} articles)` : 'No office data available',
        body: topOffice ? `This office handles ${((topOffice[1] / n) * 100).toFixed(1)}% of all deliveries — monitor its SLA closely.` : '' },
      { icon: '👤', title: topUser ? `Top Booking User: "${topUser[0]}" (${topUser[1]} articles)` : 'No user data available',
        body: topUser ? `This user processed the most bookings — useful for workload balancing and performance recognition.` : '' },
      { icon: '⚖️', title: `Average Article Weight: ${avgWeight} kg`,
        body: `Use this baseline to flag unusually heavy parcels for correct tariff and handling.` },
    ];
    this.q('ip-thoughts-content').innerHTML = `<h3>💡 Data-Driven Insights — India Post Analysis</h3>` +
      items.filter(x => x.body).map((x, i) => `<div class="thought-item">
        <div class="thought-num">${i + 1}</div>
        <div class="thought-content"><h4>${x.icon} ${x.title}</h4><p>${x.body}</p></div>
      </div>`).join('');
  }

  // ═══════════════════════════════════════════════════════
  // EXCEL EXPORT (shared by all three courier modules)
  // ═══════════════════════════════════════════════════════

  private excelDownload(rows: any[], filename: string): void {
    if (!rows || !rows.length) { alert('No data to export.'); return; }
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Data');
    const cols = Object.keys(rows[0] || {}).map(k => ({ wch: Math.max(k.length + 2, 14) }));
    (ws as any)['!cols'] = cols;
    const d = this.today.toISOString().split('T')[0];
    XLSX.writeFile(wb, `${filename}_${d}.xlsx`);
  }
}
