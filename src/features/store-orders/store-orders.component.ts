// src/features/store-orders/store-orders.component.ts
//
// Ported from the standalone "Uathayam & Ariser Unified Orders Portal" tool.
// See src/features/legacy-tool-base.ts for the porting pattern this follows.
//
// Library substitution: the source tool parsed CSVs with PapaParse
// (Papa.parse(file, {header:true, skipEmptyLines:true, complete: ...})). This
// app already depends on `xlsx`, so parseBrandCsv() below re-implements the
// same signature/return shape (array of row objects, each tagged with
// `_explicitBrand`) using XLSX.read/sheet_to_json instead of adding PapaParse
// as a new dependency.
import { AfterViewInit, Component, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import { LegacyToolBase } from '../legacy-tool-base';

interface OrderItem {
  name: string;
  sku: string;
  qty: number;
}

interface StoreOrder {
  Brand: string;
  OrderNumber: string;
  OrderDate: string;
  CustomerName: string;
  BillingCity: string;
  BillingProvince: string;
  BillingZip: string;
  PaymentType: string;
  ParsedTotal: number;
  Items: OrderItem[];
  [key: string]: any;
}

interface DateAgg {
  totalOrders: number;
  uathayam: number;
  ariser: number;
  prepaid: number;
  cod: number;
  fulfilled: number;
  unfulfilled: number;
  cancelled: number;
  revenue: number;
}

interface UnfulfilledDateAgg {
  totalUnfulfilled: number;
  prepaidUnfulfilled: number;
  codUnfulfilled: number;
  revenue: number;
}

interface GeoAgg {
  count: number;
  revenue: number;
}

type GeoType = 'province' | 'city' | 'zip';

@Component({
  selector: 'app-feature-store-orders',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './store-orders.component.html',
})
export class StoreOrdersComponent extends LegacyToolBase implements AfterViewInit {
  @ViewChild('root', { static: true }) rootRef!: ElementRef<HTMLElement>;

  // ── DOM refs (populated once in ngAfterViewInit, mirrors the original top-level consts) ──
  private uathayamFileInput!: HTMLInputElement;
  private ariserFileInput!: HTMLInputElement;
  private fileUploadSummaryEl!: HTMLElement;
  private alertSuccessEl!: HTMLElement;
  private alertErrorEl!: HTMLElement;
  private dashboardContentEl!: HTMLElement;
  private insightsGridEl!: HTMLElement;

  private brandFilterEl!: HTMLSelectElement;
  private fulfillmentFilterEl!: HTMLSelectElement;
  private financialFilterEl!: HTMLSelectElement;
  private startDateFilterEl!: HTMLInputElement;
  private endDateFilterEl!: HTMLInputElement;
  private searchFilterEl!: HTMLInputElement;

  private insTotalOrdersEl!: HTMLElement;
  private insUathayamShareEl!: HTMLElement;
  private insAriserShareEl!: HTMLElement;
  private insUathayamAOVEl!: HTMLElement;
  private insAriserAOVEl!: HTMLElement;
  private insPrepaidRateEl!: HTMLElement;
  private insCODRateEl!: HTMLElement;
  private insCancelRateEl!: HTMLElement;
  private insPrepaidCountEl!: HTMLElement;
  private insCODCountEl!: HTMLElement;
  private insFulfillmentRateEl!: HTMLElement;
  private insPendingCountEl!: HTMLElement;
  private insPendingValueEl!: HTMLElement;
  private insPrepaidPendingEl!: HTMLElement;
  private insCODPendingEl!: HTMLElement;
  private insTopSKUTableBodyEl!: HTMLElement;

  private tableBodyEl!: HTMLElement;
  private recordCountEl!: HTMLElement;
  private dateWiseTableBodyEl!: HTMLElement;
  private unfulfilledDateTableBodyEl!: HTMLElement;
  private unfulfilledSKUTableBodyEl!: HTMLElement;
  private unfulfilledOrderCountEl!: HTMLElement;
  private provinceTableBodyEl!: HTMLElement;
  private cityTableBodyEl!: HTMLElement;
  private zipTableBodyEl!: HTMLElement;

  // ── State (mirrors the original top-level `let` variables) ──
  private allOrders: StoreOrder[] = [];
  private filteredOrders: StoreOrder[] = [];

  ngAfterViewInit(): void {
    this.bindRoot(this.rootRef);
    this.cacheDomRefs();
    this.wireEvents();
  }

  // ── Setup ──────────────────────────────────────────────────────────────

  private cacheDomRefs(): void {
    this.uathayamFileInput = this.q('uathayamFile');
    this.ariserFileInput = this.q('ariserFile');
    this.fileUploadSummaryEl = this.q('fileUploadSummary');
    this.alertSuccessEl = this.q('alertSuccess');
    this.alertErrorEl = this.q('alertError');
    this.dashboardContentEl = this.q('dashboardContent');
    this.insightsGridEl = this.q('insightsGrid');

    this.brandFilterEl = this.q('brandFilter');
    this.fulfillmentFilterEl = this.q('fulfillmentFilter');
    this.financialFilterEl = this.q('financialFilter');
    this.startDateFilterEl = this.q('startDateFilter');
    this.endDateFilterEl = this.q('endDateFilter');
    this.searchFilterEl = this.q('searchFilter');

    this.insTotalOrdersEl = this.q('insTotalOrders');
    this.insUathayamShareEl = this.q('insUathayamShare');
    this.insAriserShareEl = this.q('insAriserShare');
    this.insUathayamAOVEl = this.q('insUathayamAOV');
    this.insAriserAOVEl = this.q('insAriserAOV');
    this.insPrepaidRateEl = this.q('insPrepaidRate');
    this.insCODRateEl = this.q('insCODRate');
    this.insCancelRateEl = this.q('insCancelRate');
    this.insPrepaidCountEl = this.q('insPrepaidCount');
    this.insCODCountEl = this.q('insCODCount');
    this.insFulfillmentRateEl = this.q('insFulfillmentRate');
    this.insPendingCountEl = this.q('insPendingCount');
    this.insPendingValueEl = this.q('insPendingValue');
    this.insPrepaidPendingEl = this.q('insPrepaidPending');
    this.insCODPendingEl = this.q('insCODPending');
    this.insTopSKUTableBodyEl = this.q('insTopSKUTableBody');

    this.tableBodyEl = this.q('tableBody');
    this.recordCountEl = this.q('recordCount');
    this.dateWiseTableBodyEl = this.q('dateWiseTableBody');
    this.unfulfilledDateTableBodyEl = this.q('unfulfilledDateTableBody');
    this.unfulfilledSKUTableBodyEl = this.q('unfulfilledSKUTableBody');
    this.unfulfilledOrderCountEl = this.q('unfulfilledOrderCount');
    this.provinceTableBodyEl = this.q('provinceTableBody');
    this.cityTableBodyEl = this.q('cityTableBody');
    this.zipTableBodyEl = this.q('zipTableBody');
  }

  private wireEvents(): void {
    // "⚡ Process & Merge Stores Data" — the only .btn-primary in the upload actions bar.
    this.qs<HTMLButtonElement>('.upload-actions .btn-primary')
      ?.addEventListener('click', () => void this.handleDualUpload());

    // Tab bar buttons, in DOM order — same order as the original switchTab(tabId, this) calls.
    const tabIds = ['insights-tab', 'orders-tab', 'datewise-tab', 'unfulfilled-tab', 'geo-tab'];
    this.qa<HTMLButtonElement>('.tab-bar .tab-btn').forEach((btn, i) => {
      btn.addEventListener('click', () => this.switchTab(tabIds[i], btn));
    });

    // Global filter bar.
    this.brandFilterEl.addEventListener('change', () => this.applyFilters());
    this.fulfillmentFilterEl.addEventListener('change', () => this.applyFilters());
    this.financialFilterEl.addEventListener('change', () => this.applyFilters());
    this.startDateFilterEl.addEventListener('change', () => this.applyFilters());
    this.endDateFilterEl.addEventListener('change', () => this.applyFilters());
    this.searchFilterEl.addEventListener('keyup', () => this.applyFilters());

    // "Reset All" — the only .btn-secondary in the filters bar.
    this.qs<HTMLButtonElement>('.filters-section .btn-secondary')
      ?.addEventListener('click', () => this.resetFilters());

    // Per-tab "Download Excel" buttons.
    this.qs<HTMLButtonElement>('#orders-tab .btn-success')
      ?.addEventListener('click', () => this.downloadOrdersExcel());
    this.qs<HTMLButtonElement>('#datewise-tab .btn-success')
      ?.addEventListener('click', () => this.downloadDateWiseExcel());

    const unfulfilledDownloadBtns = this.qa<HTMLButtonElement>('#unfulfilled-tab .btn-success');
    unfulfilledDownloadBtns[0]?.addEventListener('click', () => this.downloadUnfulfilledDateExcel());
    unfulfilledDownloadBtns[1]?.addEventListener('click', () => this.downloadUnfulfilledSKUExcel());

    // Geo Hub cards, in DOM order: state / city / pincode.
    const geoTypes: GeoType[] = ['province', 'city', 'zip'];
    this.qa<HTMLButtonElement>('#geo-tab .geo-card-header .btn-secondary').forEach((btn, i) => {
      btn.addEventListener('click', () => this.downloadGeoExcel(geoTypes[i]));
    });
  }

  // ── Tabs ───────────────────────────────────────────────────────────────

  private switchTab(tabId: string, btn: HTMLButtonElement): void {
    this.qa<HTMLElement>('.tab-content').forEach(el => el.classList.remove('active'));
    this.qa<HTMLElement>('.tab-btn').forEach(el => el.classList.remove('active'));
    this.q(tabId).classList.add('active');
    btn.classList.add('active');
  }

  // ── CSV parsing (PapaParse -> XLSX substitution) ──────────────────────

  private async parseBrandCsv(file: File | null, brandName: string): Promise<any[]> {
    if (!file) return [];
    const text = await file.text();
    const wb = XLSX.read(text, { type: 'string' });
    const rows = XLSX.utils.sheet_to_json<any>(wb.Sheets[wb.SheetNames[0]], { defval: '' });
    return rows
      .filter(r => Object.values(r).some(v => String(v).trim() !== ''))
      .map(row => ({ ...row, _explicitBrand: brandName }));
  }

  // ── Dual upload & order consolidation ─────────────────────────────────

  private async handleDualUpload(): Promise<void> {
    const uFile = this.uathayamFileInput.files?.[0] ?? null;
    const aFile = this.ariserFileInput.files?.[0] ?? null;

    if (!uFile && !aFile) {
      this.showAlert('Please choose at least one CSV file (Uathayam or Ariser).', 'error');
      return;
    }

    const [uRows, aRows] = await Promise.all([
      this.parseBrandCsv(uFile, 'Uathayam'),
      this.parseBrandCsv(aFile, 'Ariser'),
    ]);

    const combinedRows = [...uRows, ...aRows];
    if (combinedRows.length === 0) {
      this.showAlert('No order rows could be extracted from the uploaded files.', 'error');
      return;
    }

    // Consolidate line items by unique order number and brand.
    const ordersMap = new Map<string, StoreOrder>();

    combinedRows.forEach(row => {
      const rawOrderNumber = row['Name']?.trim();
      if (!rawOrderNumber) return;

      const brand = row._explicitBrand;
      const uniqueKey = `${brand}_${rawOrderNumber}`;

      if (!ordersMap.has(uniqueKey)) {
        const rawDate = row['Created at'] ? String(row['Created at']).split(' ')[0] : 'Unknown';
        const billingCity = (row['Billing City'] || row['Shipping City'] || 'Unspecified').trim();
        const billingProvince = (
          row['Billing Province'] || row['Billing Province Name'] || row['Shipping Province'] || 'Unspecified'
        ).trim();
        const billingZip = String(row['Billing Zip'] || row['Shipping Zip'] || 'Unspecified')
          .trim()
          .replace(/['"]+/g, '');

        const financialStatus = (row['Financial Status'] || '').toLowerCase();
        const paymentDetails = (row['Payment Method'] || row['Gateway'] || '').toLowerCase();
        let paymentType = 'Prepaid';

        if (paymentDetails.includes('cod') || paymentDetails.includes('cash')) {
          paymentType = 'COD';
        } else if (financialStatus === 'paid') {
          paymentType = 'Prepaid';
        }

        ordersMap.set(uniqueKey, {
          ...row,
          Brand: brand,
          OrderNumber: rawOrderNumber,
          OrderDate: rawDate,
          CustomerName: row['Billing Name'] || row['Shipping Name'] || 'Customer',
          BillingCity: billingCity,
          BillingProvince: billingProvince,
          BillingZip: billingZip,
          PaymentType: paymentType,
          ParsedTotal: parseFloat(row['Total'] || 0),
          Items: [],
        });
      }

      if (row['Lineitem name'] || row['Lineitem sku']) {
        const sku = String(row['Lineitem sku'] || 'NO-SKU').trim();
        const qty = parseInt(row['Lineitem quantity'] || 1, 10);
        ordersMap.get(uniqueKey)!.Items.push({
          name: row['Lineitem name'] || 'Unnamed Item',
          sku,
          qty: isNaN(qty) ? 1 : qty,
        });
      }
    });

    this.allOrders = Array.from(ordersMap.values());
    this.filteredOrders = [...this.allOrders];

    const summaryParts: string[] = [];
    if (uFile) summaryParts.push(`Uathayam: ${uFile.name}`);
    if (aFile) summaryParts.push(`Ariser: ${aFile.name}`);
    this.fileUploadSummaryEl.textContent = summaryParts.join(' | ');

    this.dashboardContentEl.style.display = 'block';
    this.showAlert(
      `Loaded ${this.allOrders.length} unique orders (${uRows.length > 0 ? 'Uathayam ✓' : ''} ${
        aRows.length > 0 ? 'Ariser ✓' : ''
      })!`,
      'success',
    );
    this.updateDashboard();
  }

  private isOrderCancelled(order: StoreOrder): boolean {
    return Boolean(order['Cancelled at'] || (order['Tags'] && String(order['Tags']).toLowerCase().includes('cancel')));
  }

  private updateDashboard(): void {
    this.updateKPIBar();
    this.applyFilters();
  }

  // ── KPI bar ────────────────────────────────────────────────────────────

  private updateKPIBar(): void {
    const uathayamOrders = this.allOrders.filter(o => o.Brand === 'Uathayam');
    const ariserOrders = this.allOrders.filter(o => o.Brand === 'Ariser');

    const fulfilled = this.allOrders.filter(
      o => !this.isOrderCancelled(o) && o['Fulfillment Status']?.toLowerCase() === 'fulfilled',
    ).length;
    const unfulfilled = this.allOrders.filter(
      o => !this.isOrderCancelled(o) && o['Fulfillment Status']?.toLowerCase() === 'unfulfilled',
    ).length;
    const cancelled = this.allOrders.filter(o => this.isOrderCancelled(o)).length;

    const totalRevenue = this.allOrders
      .filter(o => !this.isOrderCancelled(o))
      .reduce((sum, o) => sum + (o.ParsedTotal || 0), 0);

    const uathayamRev = uathayamOrders
      .filter(o => !this.isOrderCancelled(o))
      .reduce((sum, o) => sum + (o.ParsedTotal || 0), 0);

    const ariserRev = ariserOrders
      .filter(o => !this.isOrderCancelled(o))
      .reduce((sum, o) => sum + (o.ParsedTotal || 0), 0);

    const insights: { title: string; value: string | number; subtext: string; brand?: 'uathayam' | 'ariser' }[] = [
      { title: 'Total Combined Orders', value: this.allOrders.length, subtext: 'Uathayam + Ariser' },
      {
        title: 'Uathayam Orders',
        value: uathayamOrders.length,
        subtext: `Rev: ₹${uathayamRev.toLocaleString('en-IN', { maximumFractionDigits: 0 })}`,
        brand: 'uathayam',
      },
      {
        title: 'Ariser Orders',
        value: ariserOrders.length,
        subtext: `Rev: ₹${ariserRev.toLocaleString('en-IN', { maximumFractionDigits: 0 })}`,
        brand: 'ariser',
      },
      {
        title: 'Gross Revenue',
        value: '₹' + totalRevenue.toLocaleString('en-IN', { minimumFractionDigits: 2 }),
        subtext: 'Excl. cancellations',
      },
      { title: 'Fulfilled Orders', value: fulfilled, subtext: 'Delivered / Dispatched' },
      { title: 'Unfulfilled Orders', value: unfulfilled, subtext: 'Pending Pick / Pack' },
      { title: 'Cancelled', value: cancelled, subtext: 'Voided orders' },
    ];

    this.insightsGridEl.innerHTML = insights
      .map(
        item => `
                <div class="insight-card ${item.brand ? (item.brand === 'uathayam' ? 'uathayam-card' : 'ariser-card') : ''}">
                    <h3>${item.title}</h3>
                    <div class="value">${item.value}</div>
                    <div class="subtext">${item.subtext}</div>
                </div>
            `,
      )
      .join('');
  }

  // ── Filters ────────────────────────────────────────────────────────────

  private applyFilters(): void {
    const brandFilter = this.brandFilterEl.value;
    const fulfillmentFilter = this.fulfillmentFilterEl.value.toLowerCase();
    const financialFilter = this.financialFilterEl.value.toLowerCase();
    const searchFilter = this.searchFilterEl.value.toLowerCase();
    const startDate = this.startDateFilterEl.value;
    const endDate = this.endDateFilterEl.value;

    this.filteredOrders = this.allOrders.filter(order => {
      if (brandFilter && order.Brand !== brandFilter) return false;

      const cancelled = this.isOrderCancelled(order);
      let matchesFulfillment = true;

      if (fulfillmentFilter === 'cancel') {
        matchesFulfillment = cancelled;
      } else if (fulfillmentFilter) {
        matchesFulfillment = !cancelled && order['Fulfillment Status']?.toLowerCase() === fulfillmentFilter;
      }

      let matchesFinancial = true;
      if (financialFilter === 'paid') {
        matchesFinancial = order.PaymentType === 'Prepaid';
      } else if (financialFilter === 'cod') {
        matchesFinancial = order.PaymentType === 'COD';
      }

      let matchesDateRange = true;
      const orderDate = order.OrderDate;
      if (startDate && orderDate && orderDate < startDate) matchesDateRange = false;
      if (endDate && orderDate && orderDate > endDate) matchesDateRange = false;

      const skuString = order.Items.map(i => i.sku).join(' ').toLowerCase();
      const itemNames = order.Items.map(i => i.name).join(' ').toLowerCase();

      const matchesSearch =
        !searchFilter ||
        order.OrderNumber?.toLowerCase().includes(searchFilter) ||
        order['Email']?.toLowerCase().includes(searchFilter) ||
        order.CustomerName?.toLowerCase().includes(searchFilter) ||
        order.BillingCity?.toLowerCase().includes(searchFilter) ||
        order.BillingProvince?.toLowerCase().includes(searchFilter) ||
        order.BillingZip?.toLowerCase().includes(searchFilter) ||
        skuString.includes(searchFilter) ||
        itemNames.includes(searchFilter);

      return matchesFulfillment && matchesFinancial && matchesDateRange && matchesSearch;
    });

    this.renderInsightsTab();
    this.renderTable();
    this.renderDateWiseTable();
    this.renderUnfulfilledDateTables();
    this.renderGeoAnalytics();
  }

  private resetFilters(): void {
    this.brandFilterEl.value = '';
    this.fulfillmentFilterEl.value = '';
    this.financialFilterEl.value = '';
    this.startDateFilterEl.value = '';
    this.endDateFilterEl.value = '';
    this.searchFilterEl.value = '';
    this.applyFilters();
  }

  // ── Tab 1: Executive Insights ──────────────────────────────────────────

  private renderInsightsTab(): void {
    const total = this.filteredOrders.length;
    const cancelled = this.filteredOrders.filter(o => this.isOrderCancelled(o)).length;
    const activeOrders = this.filteredOrders.filter(o => !this.isOrderCancelled(o));
    const activeCount = activeOrders.length;

    const uathayamActive = activeOrders.filter(o => o.Brand === 'Uathayam');
    const ariserActive = activeOrders.filter(o => o.Brand === 'Ariser');

    const uathayamRev = uathayamActive.reduce((s, o) => s + (o.ParsedTotal || 0), 0);
    const ariserRev = ariserActive.reduce((s, o) => s + (o.ParsedTotal || 0), 0);

    const uathayamAOV = uathayamActive.length > 0 ? uathayamRev / uathayamActive.length : 0;
    const ariserAOV = ariserActive.length > 0 ? ariserRev / ariserActive.length : 0;

    const prepaidOrders = activeOrders.filter(o => o.PaymentType === 'Prepaid');
    const codOrders = activeOrders.filter(o => o.PaymentType === 'COD');

    const prepaidRate = activeCount > 0 ? ((prepaidOrders.length / activeCount) * 100).toFixed(1) : '0';
    const codRate = activeCount > 0 ? ((codOrders.length / activeCount) * 100).toFixed(1) : '0';
    const cancelRate = total > 0 ? ((cancelled / total) * 100).toFixed(1) : '0';

    const fulfilledOrders = activeOrders.filter(o => o['Fulfillment Status']?.toLowerCase() === 'fulfilled');
    const unfulfilledOrders = activeOrders.filter(o => o['Fulfillment Status']?.toLowerCase() !== 'fulfilled');
    const fulfillmentRate = activeCount > 0 ? ((fulfilledOrders.length / activeCount) * 100).toFixed(1) : '0';

    const pendingValue = unfulfilledOrders.reduce((s, o) => s + (o.ParsedTotal || 0), 0);
    const prepaidPending = unfulfilledOrders.filter(o => o.PaymentType === 'Prepaid').length;
    const codPending = unfulfilledOrders.filter(o => o.PaymentType === 'COD').length;

    // Card 1: Brand Comparison
    this.insTotalOrdersEl.textContent = activeCount.toLocaleString('en-IN');
    this.insUathayamShareEl.textContent =
      activeCount > 0 ? `${((uathayamActive.length / activeCount) * 100).toFixed(1)}% (${uathayamActive.length})` : '0%';
    this.insAriserShareEl.textContent =
      activeCount > 0 ? `${((ariserActive.length / activeCount) * 100).toFixed(1)}% (${ariserActive.length})` : '0%';
    this.insUathayamAOVEl.textContent = `₹${uathayamAOV.toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;
    this.insAriserAOVEl.textContent = `₹${ariserAOV.toLocaleString('en-IN', {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    })}`;

    // Card 2: Payment & Risk
    this.insPrepaidRateEl.textContent = `${prepaidRate}%`;
    this.insCODRateEl.textContent = `${codRate}%`;
    this.insCancelRateEl.textContent = `${cancelRate}%`;
    this.insPrepaidCountEl.textContent = `${prepaidOrders.length} orders`;
    this.insCODCountEl.textContent = `${codOrders.length} orders`;

    // Card 3: Fulfillment Health
    this.insFulfillmentRateEl.textContent = `${fulfillmentRate}%`;
    this.insPendingCountEl.textContent = `${unfulfilledOrders.length} orders`;
    this.insPendingValueEl.textContent = `₹${pendingValue.toLocaleString('en-IN', { minimumFractionDigits: 2 })}`;
    this.insPrepaidPendingEl.textContent = `${prepaidPending} orders`;
    this.insCODPendingEl.textContent = `${codPending} orders`;

    // Aggregate Top Unfulfilled SKUs to pick.
    const skuMap: Record<string, { brand: string; sku: string; name: string; qty: number }> = {};
    unfulfilledOrders.forEach(order => {
      order.Items.forEach(item => {
        const key = `${order.Brand}__${item.sku}`;
        if (!skuMap[key]) {
          skuMap[key] = { brand: order.Brand, sku: item.sku, name: item.name, qty: 0 };
        }
        skuMap[key].qty += item.qty;
      });
    });

    const sortedSKUs = Object.values(skuMap).sort((a, b) => b.qty - a.qty);
    const topSKUs = sortedSKUs.slice(0, 10);

    if (topSKUs.length === 0) {
      this.insTopSKUTableBodyEl.innerHTML =
        '<tr><td colspan="4" class="no-data">No pending SKUs. All current orders fulfilled!</td></tr>';
    } else {
      this.insTopSKUTableBodyEl.innerHTML = topSKUs
        .map(
          item => `
                    <tr>
                        <td><span class="brand-pill ${item.brand === 'Ariser' ? 'brand-ariser' : 'brand-uathayam'}">${item.brand}</span></td>
                        <td><span class="sku-badge">${item.sku}</span></td>
                        <td>${item.name}</td>
                        <td><strong>${item.qty} units</strong></td>
                    </tr>
                `,
        )
        .join('');
    }
  }

  // ── Tab 2: Consolidated Order Master ───────────────────────────────────

  private renderTable(): void {
    this.recordCountEl.textContent = `${this.filteredOrders.length} order${this.filteredOrders.length !== 1 ? 's' : ''}`;

    if (this.filteredOrders.length === 0) {
      this.tableBodyEl.innerHTML = '<tr><td colspan="11" class="no-data">No records match the active criteria.</td></tr>';
      return;
    }

    this.tableBodyEl.innerHTML = this.filteredOrders
      .map(order => {
        const cancelled = this.isOrderCancelled(order);
        const fulfillmentStatus = order['Fulfillment Status']?.toLowerCase() || 'unfulfilled';
        const orderStatus = cancelled
          ? 'Cancelled'
          : fulfillmentStatus.charAt(0).toUpperCase() + fulfillmentStatus.slice(1);
        const statusClass = cancelled
          ? 'status-cancel'
          : fulfillmentStatus === 'fulfilled'
            ? 'status-fulfilled'
            : 'status-unfulfilled';

        const paymentClass = order.PaymentType === 'COD' ? 'status-cod' : 'status-paid';
        const brandClass = order.Brand === 'Ariser' ? 'brand-ariser' : 'brand-uathayam';
        const totalQty = order.Items.length > 0 ? order.Items.reduce((sum, item) => sum + item.qty, 0) : 1;

        return `
                    <tr>
                        <td><span class="brand-pill ${brandClass}">${order.Brand}</span></td>
                        <td><strong>${order.OrderNumber}</strong></td>
                        <td>${order.OrderDate}</td>
                        <td>${order.CustomerName}</td>
                        <td>${order.BillingCity}</td>
                        <td>${order.BillingProvince}</td>
                        <td>${order.BillingZip}</td>
                        <td>${totalQty}</td>
                        <td>₹${order.ParsedTotal.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                        <td><span class="status-pill ${paymentClass}">${order.PaymentType}</span></td>
                        <td><span class="status-pill ${statusClass}">${orderStatus}</span></td>
                    </tr>
                `;
      })
      .join('');
  }

  // ── Tab 3: Date-Wise Overview ───────────────────────────────────────────

  private getDateWiseAggregations(): [string, DateAgg][] {
    const dateMap: Record<string, DateAgg> = {};

    this.filteredOrders.forEach(order => {
      const date = order.OrderDate || 'Unknown';
      if (!dateMap[date]) {
        dateMap[date] = {
          totalOrders: 0,
          uathayam: 0,
          ariser: 0,
          prepaid: 0,
          cod: 0,
          fulfilled: 0,
          unfulfilled: 0,
          cancelled: 0,
          revenue: 0,
        };
      }

      dateMap[date].totalOrders++;
      dateMap[date].revenue += order.ParsedTotal || 0;

      if (order.Brand === 'Ariser') dateMap[date].ariser++;
      else dateMap[date].uathayam++;

      if (this.isOrderCancelled(order)) {
        dateMap[date].cancelled++;
      } else {
        const fStatus = order['Fulfillment Status']?.toLowerCase();
        if (fStatus === 'fulfilled') dateMap[date].fulfilled++;
        else dateMap[date].unfulfilled++;

        if (order.PaymentType === 'COD') dateMap[date].cod++;
        else dateMap[date].prepaid++;
      }
    });

    return Object.entries(dateMap).sort((a, b) => b[0].localeCompare(a[0]));
  }

  private renderDateWiseTable(): void {
    const data = this.getDateWiseAggregations();

    if (data.length === 0) {
      this.dateWiseTableBodyEl.innerHTML =
        '<tr><td colspan="10" class="no-data">No data found for the current date selection.</td></tr>';
      return;
    }

    this.dateWiseTableBodyEl.innerHTML = data
      .map(
        ([date, d]) => `
                <tr>
                    <td><strong>${date}</strong></td>
                    <td><strong>${d.totalOrders}</strong></td>
                    <td><span class="brand-pill brand-uathayam">${d.uathayam}</span></td>
                    <td><span class="brand-pill brand-ariser">${d.ariser}</span></td>
                    <td><span class="tag-metric prepaid">${d.prepaid}</span></td>
                    <td><span class="tag-metric cod">${d.cod}</span></td>
                    <td><span class="tag-metric fulfilled">${d.fulfilled}</span></td>
                    <td><span class="tag-metric unfulfilled">${d.unfulfilled}</span></td>
                    <td><span class="tag-metric cancelled">${d.cancelled}</span></td>
                    <td>₹${d.revenue.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                </tr>
            `,
      )
      .join('');
  }

  // ── Tab 4: Unfulfilled — date summary + order-wise SKU pick list ────────

  private getUnfulfilledOrdersList(): StoreOrder[] {
    return this.filteredOrders
      .filter(order => !this.isOrderCancelled(order) && order['Fulfillment Status']?.toLowerCase() !== 'fulfilled')
      .sort((a, b) => b.OrderDate.localeCompare(a.OrderDate));
  }

  private renderUnfulfilledDateTables(): void {
    const unfulfilledOrders = this.getUnfulfilledOrdersList();
    const dateMap: Record<string, UnfulfilledDateAgg> = {};

    unfulfilledOrders.forEach(order => {
      const date = order.OrderDate || 'Unknown';
      if (!dateMap[date]) {
        dateMap[date] = { totalUnfulfilled: 0, prepaidUnfulfilled: 0, codUnfulfilled: 0, revenue: 0 };
      }
      dateMap[date].totalUnfulfilled++;
      dateMap[date].revenue += order.ParsedTotal || 0;
      if (order.PaymentType === 'COD') dateMap[date].codUnfulfilled++;
      else dateMap[date].prepaidUnfulfilled++;
    });

    const dateList = Object.entries(dateMap).sort((a, b) => b[0].localeCompare(a[0]));

    if (dateList.length === 0) {
      this.unfulfilledDateTableBodyEl.innerHTML =
        '<tr><td colspan="5" class="no-data">Zero unfulfilled orders found.</td></tr>';
    } else {
      this.unfulfilledDateTableBodyEl.innerHTML = dateList
        .map(
          ([date, d]) => `
                    <tr>
                        <td><strong>${date}</strong></td>
                        <td><strong>${d.totalUnfulfilled}</strong></td>
                        <td><span class="tag-metric prepaid">${d.prepaidUnfulfilled}</span></td>
                        <td><span class="tag-metric cod">${d.codUnfulfilled}</span></td>
                        <td>₹${d.revenue.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                    </tr>
                `,
        )
        .join('');
    }

    this.unfulfilledOrderCountEl.textContent = `${unfulfilledOrders.length} pending`;

    if (unfulfilledOrders.length === 0) {
      this.unfulfilledSKUTableBodyEl.innerHTML =
        '<tr><td colspan="9" class="no-data">No unfulfilled orders to fulfill.</td></tr>';
      return;
    }

    this.unfulfilledSKUTableBodyEl.innerHTML = unfulfilledOrders
      .map(order => {
        const totalQty = order.Items.reduce((sum, item) => sum + item.qty, 0);
        const paymentClass = order.PaymentType === 'COD' ? 'status-cod' : 'status-paid';
        const brandClass = order.Brand === 'Ariser' ? 'brand-ariser' : 'brand-uathayam';

        const skuBadges = order.Items.map(
          item => `
                    <span class="sku-badge" title="${item.name.replace(/"/g, '&quot;')}">
                        ${item.sku} <span class="sku-qty">x${item.qty}</span>
                    </span>
                `,
        ).join('');

        return `
                    <tr>
                        <td><span class="brand-pill ${brandClass}">${order.Brand}</span></td>
                        <td><strong>${order.OrderNumber}</strong></td>
                        <td>${order.OrderDate}</td>
                        <td>${order.CustomerName}</td>
                        <td>${order.BillingCity}</td>
                        <td><span class="status-pill ${paymentClass}">${order.PaymentType}</span></td>
                        <td><div class="sku-badge-container">${skuBadges}</div></td>
                        <td><strong>${totalQty}</strong></td>
                        <td>₹${order.ParsedTotal.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                    </tr>
                `;
      })
      .join('');
  }

  // ── Tab 5: Geographic Hub ───────────────────────────────────────────────

  private renderGeoAnalytics(): void {
    const provinceMap: Record<string, GeoAgg> = {};
    const cityMap: Record<string, GeoAgg> = {};
    const zipMap: Record<string, GeoAgg> = {};

    this.filteredOrders.forEach(order => {
      const p = order.BillingProvince || 'Unspecified';
      const c = order.BillingCity || 'Unspecified';
      const z = order.BillingZip || 'Unspecified';
      const rev = order.ParsedTotal || 0;

      if (!provinceMap[p]) provinceMap[p] = { count: 0, revenue: 0 };
      provinceMap[p].count++;
      provinceMap[p].revenue += rev;

      if (!cityMap[c]) cityMap[c] = { count: 0, revenue: 0 };
      cityMap[c].count++;
      cityMap[c].revenue += rev;

      if (!zipMap[z]) zipMap[z] = { count: 0, revenue: 0 };
      zipMap[z].count++;
      zipMap[z].revenue += rev;
    });

    this.renderGeoTable(this.provinceTableBodyEl, Object.entries(provinceMap).sort((a, b) => b[1].count - a[1].count));
    this.renderGeoTable(this.cityTableBodyEl, Object.entries(cityMap).sort((a, b) => b[1].count - a[1].count));
    this.renderGeoTable(this.zipTableBodyEl, Object.entries(zipMap).sort((a, b) => b[1].count - a[1].count));
  }

  private renderGeoTable(tbody: HTMLElement, dataset: [string, GeoAgg][]): void {
    if (dataset.length === 0) {
      tbody.innerHTML = '<tr><td colspan="3" class="no-data">No records found.</td></tr>';
      return;
    }

    tbody.innerHTML = dataset
      .slice(0, 15)
      .map(
        ([location, data]) => `
                <tr>
                    <td><strong>${location}</strong></td>
                    <td>${data.count}</td>
                    <td>₹${data.revenue.toLocaleString('en-IN', { minimumFractionDigits: 2 })}</td>
                </tr>
            `,
      )
      .join('');
  }

  // ── Excel exports ────────────────────────────────────────────────────────

  private downloadOrdersExcel(): void {
    if (!this.filteredOrders || this.filteredOrders.length === 0) {
      this.showAlert('No filtered orders available to download.', 'error');
      return;
    }

    const exportData = this.filteredOrders.map(order => {
      const cancelled = this.isOrderCancelled(order);
      const fulfillmentStatus = order['Fulfillment Status']?.toLowerCase() || 'unfulfilled';
      const orderStatus = cancelled
        ? 'Cancelled'
        : fulfillmentStatus.charAt(0).toUpperCase() + fulfillmentStatus.slice(1);
      const totalQty = order.Items.reduce((sum, item) => sum + item.qty, 0);
      const itemsSummary = order.Items.map(i => `${i.name} [SKU: ${i.sku}] (x${i.qty})`).join('; ');

      return {
        Brand: order.Brand,
        'Order #': order.OrderNumber,
        Date: order.OrderDate,
        'Customer Name': order.CustomerName,
        'Billing City': order.BillingCity,
        'Billing Province (State)': order.BillingProvince,
        'Billing Zip (PIN)': order.BillingZip,
        'Products & SKUs': itemsSummary,
        'Total Quantity': totalQty,
        'Total (₹)': order.ParsedTotal,
        'Payment Mode': order.PaymentType,
        'Fulfillment Status': orderStatus,
      };
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Combined_Orders');
    XLSX.writeFile(wb, `Unified_Orders_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  private downloadDateWiseExcel(): void {
    const data = this.getDateWiseAggregations();
    if (data.length === 0) {
      this.showAlert('No data to export.', 'error');
      return;
    }

    const rows = data.map(([date, d]) => ({
      Date: date,
      'Total Orders': d.totalOrders,
      'Uathayam Orders': d.uathayam,
      'Ariser Orders': d.ariser,
      'Prepaid Orders': d.prepaid,
      'COD Orders': d.cod,
      'Fulfilled Orders': d.fulfilled,
      'Unfulfilled Orders': d.unfulfilled,
      'Cancelled Orders': d.cancelled,
      'Total Revenue (₹)': d.revenue,
    }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Date_Wise_Summary');
    XLSX.writeFile(wb, `DateWise_Summary_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  private downloadUnfulfilledDateExcel(): void {
    const unfulfilledOrders = this.getUnfulfilledOrdersList();
    const dateMap: Record<string, { total: number; prepaid: number; cod: number; rev: number }> = {};

    unfulfilledOrders.forEach(order => {
      const date = order.OrderDate || 'Unknown';
      if (!dateMap[date]) dateMap[date] = { total: 0, prepaid: 0, cod: 0, rev: 0 };
      dateMap[date].total++;
      dateMap[date].rev += order.ParsedTotal || 0;
      if (order.PaymentType === 'COD') dateMap[date].cod++;
      else dateMap[date].prepaid++;
    });

    const rows = Object.entries(dateMap)
      .sort((a, b) => b[0].localeCompare(a[0]))
      .map(([date, d]) => ({
        Date: date,
        'Total Unfulfilled Orders': d.total,
        'Prepaid Unfulfilled': d.prepaid,
        'COD Unfulfilled': d.cod,
        'Pending Value (₹)': d.rev,
      }));

    if (rows.length === 0) {
      this.showAlert('No unfulfilled records to export.', 'error');
      return;
    }

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Unfulfilled_DateSummary');
    XLSX.writeFile(wb, `Unfulfilled_DateSummary_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  private downloadUnfulfilledSKUExcel(): void {
    const unfulfilledOrders = this.getUnfulfilledOrdersList();
    if (unfulfilledOrders.length === 0) {
      this.showAlert('No unfulfilled orders to export.', 'error');
      return;
    }

    const rows = unfulfilledOrders.map(order => {
      const skuSummary = order.Items.map(i => `${i.sku} (x${i.qty})`).join(', ');
      const totalQty = order.Items.reduce((sum, i) => sum + i.qty, 0);

      return {
        Brand: order.Brand,
        'Order #': order.OrderNumber,
        Date: order.OrderDate,
        'Customer Name': order.CustomerName,
        City: order.BillingCity,
        State: order.BillingProvince,
        'PIN Code': order.BillingZip,
        'Payment Mode': order.PaymentType,
        'SKUs (Code x Qty)': skuSummary,
        'Total Items Qty': totalQty,
        'Order Value (₹)': order.ParsedTotal,
      };
    });

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Unfulfilled_SKU_PickList');
    XLSX.writeFile(wb, `Pending_SKU_List_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  private downloadGeoExcel(type: GeoType): void {
    const dataset: Record<string, GeoAgg> = {};
    let label = '';

    this.filteredOrders.forEach(order => {
      let key = '';
      if (type === 'province') {
        key = order.BillingProvince || 'Unspecified';
        label = 'Billing State';
      } else if (type === 'city') {
        key = order.BillingCity || 'Unspecified';
        label = 'Billing City';
      } else if (type === 'zip') {
        key = order.BillingZip || 'Unspecified';
        label = 'Billing PIN';
      }

      if (!dataset[key]) dataset[key] = { count: 0, revenue: 0 };
      dataset[key].count++;
      dataset[key].revenue += order.ParsedTotal || 0;
    });

    const rows = Object.entries(dataset)
      .sort((a, b) => b[1].count - a[1].count)
      .map(([loc, d]) => ({
        [label]: loc,
        'Orders Placed': d.count,
        'Revenue (₹)': d.revenue,
      }));

    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `${type.toUpperCase()}_Stats`);
    XLSX.writeFile(wb, `Shopify_Geo_${type}_${new Date().toISOString().split('T')[0]}.xlsx`);
  }

  // ── Alerts ─────────────────────────────────────────────────────────────

  private showAlert(msg: string, type: 'success' | 'error'): void {
    if (type === 'success') {
      this.alertSuccessEl.textContent = '✓ ' + msg;
      this.alertSuccessEl.style.display = 'block';
      this.alertErrorEl.style.display = 'none';
      setTimeout(() => {
        this.alertSuccessEl.style.display = 'none';
      }, 4000);
    } else {
      this.alertErrorEl.textContent = '❌ ' + msg;
      this.alertErrorEl.style.display = 'block';
      this.alertSuccessEl.style.display = 'none';
      setTimeout(() => {
        this.alertErrorEl.style.display = 'none';
      }, 5000);
    }
  }
}
