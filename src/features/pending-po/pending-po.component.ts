// src/features/pending-po/pending-po.component.ts
//
// Ported from the standalone "All-in-One Operations Hub" tool — APP 3: PENDING PO section.
// See src/features/legacy-tool-base.ts for the porting pattern this follows.
import { AfterViewInit, Component, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { LegacyToolBase } from '../legacy-tool-base';

interface PoItem {
  name: string;
  sku: string;
  total: number;
  cod: number;
  prepaid: number;
  category: string;
}

interface CustomRule {
  keyword: string;
  category: string;
}

interface CategoryColor {
  hex: string;
  tint: string;
}

@Component({
  selector: 'app-feature-pending-po',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './pending-po.component.html',
})
export class PendingPoComponent extends LegacyToolBase implements AfterViewInit {
  @ViewChild('root', { static: true }) rootRef!: ElementRef<HTMLElement>;

  private static readonly CATEGORY_ORDER = [
    'SHIRTS', 'T-SHIRTS', 'DRESSES', 'TROUSERS', 'SHORTS', 'SKIRTS', 'JACKETS', 'SWEATERS', 'OTHERS',
  ];

  private static readonly CATEGORY_COLORS: Record<string, CategoryColor> = {
    SHIRTS: { hex: '1e3a5f', tint: 'e3f0ff' },
    'T-SHIRTS': { hex: '2d5016', tint: 'e8f5e0' },
    DRESSES: { hex: '6b21a8', tint: 'f3e8ff' },
    TROUSERS: { hex: '9333ea', tint: 'f3e8ff' },
    SHORTS: { hex: 'dc2626', tint: 'ffe8e8' },
    SKIRTS: { hex: 'e11d48', tint: 'ffe8f0' },
    JACKETS: { hex: '7c3aed', tint: 'f5e6ff' },
    SWEATERS: { hex: 'f59e0b', tint: 'fff7e0' },
    OTHERS: { hex: '6b7280', tint: 'f3f4f6' },
  };

  // ── DOM refs (populated once in ngAfterViewInit, mirrors the original top-level consts) ──
  private dropzonePO!: HTMLElement;
  private fileInputPO!: HTMLInputElement;
  private filenamePO!: HTMLElement;
  private statusPO!: HTMLElement;
  private downloadBtnPO!: HTMLButtonElement;
  private statSkus!: HTMLElement;
  private statTotal!: HTMLElement;
  private statCod!: HTMLElement;
  private statPrepaid!: HTMLElement;
  private catListContainer!: HTMLElement;
  private othersBox!: HTMLElement;
  private othersToggle!: HTMLElement;
  private othersListBody!: HTMLElement;
  private rulesTableBody!: HTMLElement;

  // ── State (mirrors the original top-level `let` variables) ──
  private aggregatedData: PoItem[] | null = null;
  private lastTotalRows = 0;
  private customRules: CustomRule[] = [];

  ngAfterViewInit(): void {
    this.bindRoot(this.rootRef);
    this.cacheDomRefs();
    this.wireEvents();
    this.q('todayDate').textContent = new Date().toLocaleDateString('en-GB', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
    });
  }

  /**
   * The source tool's escapeHtml() also escapes quotes/apostrophes (it's used both as text
   * content and inside HTML attribute values, e.g. `value="${escapeHtml(it.name)}"`).
   * LegacyToolBase.escapeHtml() only escapes &, < and > (sufficient for the other ported
   * tools, which never interpolate into an attribute value), so it's overridden here to
   * match the source's exact behavior and keep attribute-context interpolation safe.
   */
  protected override escapeHtml(s: unknown): string {
    return String(s ?? '').replace(/[&<>"']/g, (c) => (
      { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c] as string
    ));
  }

  // ── Setup ──────────────────────────────────────────────────────────────

  private cacheDomRefs(): void {
    this.dropzonePO = this.q('dropzonePO');
    this.fileInputPO = this.q('fileInputPO');
    this.filenamePO = this.q('filenamePO');
    this.statusPO = this.q('statusPO');
    this.downloadBtnPO = this.q('downloadBtnPO');
    this.statSkus = this.q('stat-skus');
    this.statTotal = this.q('stat-total');
    this.statCod = this.q('stat-cod');
    this.statPrepaid = this.q('stat-prepaid');
    this.catListContainer = this.q('catListContainer');
    this.othersBox = this.q('othersBox');
    this.othersToggle = this.q('othersToggle');
    this.othersListBody = this.q('othersListBody');
    this.rulesTableBody = this.q('rulesTableBody');
  }

  private wireEvents(): void {
    this.dropzonePO.addEventListener('click', () => this.fileInputPO.click());
    this.dropzonePO.addEventListener('dragover', (e) => {
      e.preventDefault();
      this.dropzonePO.classList.add('drag');
    });
    this.dropzonePO.addEventListener('dragleave', () => this.dropzonePO.classList.remove('drag'));
    this.dropzonePO.addEventListener('drop', (e) => {
      e.preventDefault();
      this.dropzonePO.classList.remove('drag');
      const dt = (e as DragEvent).dataTransfer;
      if (dt) this.handleFilesPO(dt.files);
    });
    this.fileInputPO.addEventListener('change', (e) => {
      const files = (e.target as HTMLInputElement).files;
      if (files) this.handleFilesPO(files);
    });

    this.othersToggle.addEventListener('click', () => {
      this.othersBox.classList.toggle('open');
    });

    this.downloadBtnPO.addEventListener('click', async () => {
      if (!this.aggregatedData) return;
      await this.generateXlsxPO(this.aggregatedData);
    });
  }

  // ── File import ────────────────────────────────────────────────────────

  private async handleFilesPO(fileList: FileList): Promise<void> {
    const files = Array.from(fileList).filter((f) => f.name.toLowerCase().match(/\.(xlsx?|csv)$/));
    if (files.length === 0) {
      this.showStatusPO('No valid files found', false);
      return;
    }

    const file = files[0];
    this.filenamePO.textContent = file.name;
    this.filenamePO.style.display = 'block';

    try {
      const buf = await file.arrayBuffer();
      let rows: unknown[][];
      if (file.name.toLowerCase().endsWith('.csv')) {
        const text = new TextDecoder().decode(buf);
        rows = this.parseCSVPO(text);
      } else {
        const wb = XLSX.read(buf, { type: 'array' });
        rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }) as unknown[][];
      }

      this.processDataPO(rows);
      this.showStatusPO('File loaded successfully', true);
    } catch (err) {
      console.error(err);
      this.showStatusPO('Error reading file: ' + (err as Error).message, false);
    }
  }

  private parseCSVPO(text: string): string[][] {
    const lines = text.split(/\r\n|\n/).filter((l) => l.trim());
    return lines.map((line) => {
      const result: string[] = [];
      let current = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const char = line[i];
        if (char === '"') inQuotes = !inQuotes;
        else if (char === ',' && !inQuotes) {
          result.push(current.trim());
          current = '';
        } else current += char;
      }
      result.push(current.trim());
      return result;
    });
  }

  private processDataPO(rows: unknown[][]): void {
    if (rows.length < 2) {
      this.showStatusPO('No data found', false);
      return;
    }

    const header = rows[0].map((h) => String(h).toLowerCase().trim());
    const nameIdx = header.findIndex((h) => h.includes('name') || h.includes('product'));
    const skuIdx = header.findIndex((h) => h.includes('sku') || h.includes('code'));
    const qtyIdx = header.findIndex((h) => h.includes('qty') || h.includes('quantity') || h.includes('units'));

    if (nameIdx === -1 || skuIdx === -1 || qtyIdx === -1) {
      this.showStatusPO('Required columns not found (Name, SKU, Qty)', false);
      return;
    }

    const items: PoItem[] = [];
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      if (!row[nameIdx] || !row[skuIdx] || !row[qtyIdx]) continue;

      const qty = parseInt(String(row[qtyIdx]), 10) || 0;
      items.push({
        name: String(row[nameIdx]).trim(),
        sku: String(row[skuIdx]).trim(),
        total: qty,
        cod: qty,
        prepaid: 0,
        category: 'OTHERS',
      });
    }

    if (items.length === 0) {
      this.showStatusPO('No valid data rows found', false);
      return;
    }

    this.lastTotalRows = items.length;
    this.aggregatedData = items;
    this.reclassifyAll();
    this.renderResults(this.aggregatedData, items.length);
  }

  // ── Classification ─────────────────────────────────────────────────────

  private reclassifyAll(): void {
    if (!this.aggregatedData) return;
    this.aggregatedData.forEach((item) => {
      item.category = this.classifyItem(item.name, item.sku);
    });
  }

  private classifyItem(name: string, sku: string): string {
    const n = (name + ' ' + sku).toUpperCase();
    for (const rule of this.customRules) {
      if (n.includes(rule.keyword.toUpperCase())) {
        return rule.category;
      }
    }
    if (n.includes('SHIRT') && !n.includes('T-')) return 'SHIRTS';
    if (n.includes('T-SHIRT') || n.includes('TSHIRT') || n.includes('TEE')) return 'T-SHIRTS';
    if (n.includes('DRESS')) return 'DRESSES';
    if (n.includes('TROUSER') || n.includes('PANTS')) return 'TROUSERS';
    if (n.includes('SHORT')) return 'SHORTS';
    if (n.includes('SKIRT')) return 'SKIRTS';
    if (n.includes('JACKET')) return 'JACKETS';
    if (n.includes('SWEATER') || n.includes('CARDIGAN') || n.includes('PULLOVER')) return 'SWEATERS';
    return 'OTHERS';
  }

  private getOrAssignColor(cat: string): CategoryColor {
    return PendingPoComponent.CATEGORY_COLORS[cat] || { hex: '6b7280', tint: 'f3f4f6' };
  }

  // ── Rendering ──────────────────────────────────────────────────────────

  private renderResults(items: PoItem[], totalRows: number): void {
    void totalRows; // unused in the source tool too — kept for signature fidelity
    const stats = {
      skus: items.length,
      total: items.reduce((s, i) => s + i.total, 0),
      cod: items.reduce((s, i) => s + i.cod, 0),
      prepaid: items.reduce((s, i) => s + i.prepaid, 0),
    };

    this.statSkus.textContent = String(stats.skus);
    this.statTotal.textContent = String(stats.total);
    this.statCod.textContent = String(stats.cod);
    this.statPrepaid.textContent = String(stats.prepaid);

    const cats = [...new Set(items.map((i) => i.category))].sort((a, b) => {
      const aIdx = PendingPoComponent.CATEGORY_ORDER.indexOf(a);
      const bIdx = PendingPoComponent.CATEGORY_ORDER.indexOf(b);
      return (aIdx === -1 ? 999 : aIdx) - (bIdx === -1 ? 999 : bIdx);
    });

    let html = '<div class="cat-row head"><div class="name">CATEGORY</div><span>SKUs</span><span>COD</span><span>TOTAL</span></div>';
    cats.filter((c) => c !== 'OTHERS').forEach((cat) => {
      const sub = items.filter((i) => i.category === cat);
      const color = this.getOrAssignColor(cat);
      html += `<div class="cat-row">
                <div class="name">
                    <div class="swatch" style="background:#${color.hex};"></div>
                    ${cat}
                </div>
                <span>${sub.length}</span>
                <span>${sub.reduce((s, i) => s + i.cod, 0)}</span>
                <span>${sub.reduce((s, i) => s + i.total, 0)}</span>
            </div>`;
    });
    this.catListContainer.innerHTML = html;

    const othersItems = items.filter((i) => i.category === 'OTHERS');
    if (othersItems.length > 0) {
      this.othersBox.style.display = 'block';
      this.othersListBody.innerHTML = othersItems
        .slice()
        .sort((a, b) => a.name.localeCompare(b.name))
        .map((it) => `
                    <div class="others-item">
                        <span class="oi-name">${this.escapeHtml(it.name)}</span>
                        <span class="oi-sku">${this.escapeHtml(it.sku)}</span>
                        <span class="oi-qty">${it.total} units</span>
                    </div>
                    <div class="oi-assign">
                        <input type="text" class="oi-keyword" value="${this.escapeHtml(it.name)}" title="Keyword">
                        <input type="text" class="oi-category" placeholder="Category" list="categoryDatalist">
                        <button class="oi-assign-btn" type="button" data-sku="${this.escapeHtml(it.sku)}">Assign</button>
                    </div>`)
        .join('');

      this.othersListBody.querySelectorAll<HTMLButtonElement>('.oi-assign-btn').forEach((btn) => {
        btn.addEventListener('click', () => {
          const wrap = btn.closest<HTMLElement>('.oi-assign');
          const kwInput = wrap?.querySelector<HTMLInputElement>('.oi-keyword');
          const catInput = wrap?.querySelector<HTMLInputElement>('.oi-category');
          if (!kwInput || !catInput) return;
          const ok = this.addCustomRule(kwInput.value, catInput.value);
          if (!ok) {
            catInput.style.borderColor = '#a8431f';
            return;
          }
          this.renderCustomRulesList();
          this.reclassifyAll();
          this.renderResults(this.aggregatedData!, this.lastTotalRows);
        });
      });
    } else {
      this.othersBox.style.display = 'none';
    }

    this.downloadBtnPO.disabled = false;
  }

  private addCustomRule(keyword: string, category: string): boolean {
    if (!keyword || !category) return false;
    if (!PendingPoComponent.CATEGORY_ORDER.includes(category)) return false;
    this.customRules.push({ keyword, category });
    return true;
  }

  private renderCustomRulesList(): void {
    this.rulesTableBody.innerHTML = this.customRules.map((r, i) => `
            <tr>
                <td>${this.escapeHtml(r.keyword)}</td>
                <td>${r.category}</td>
                <td><button class="oi-assign button" type="button" data-idx="${i}" style="padding:4px 8px;font-size:11px;background:#dc2626;">Remove</button></td>
            </tr>
        `).join('');

    this.rulesTableBody.querySelectorAll<HTMLButtonElement>('button').forEach((btn) => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset['idx'] || '0', 10);
        this.customRules.splice(idx, 1);
        this.renderCustomRulesList();
        this.reclassifyAll();
        this.renderResults(this.aggregatedData!, this.lastTotalRows);
      });
    });
  }

  private showStatusPO(msg: string, ok: boolean): void {
    this.statusPO.textContent = msg;
    this.statusPO.className = ok ? 'status-po ok' : 'status-po err';
  }

  // ── Export to Excel (styled multi-sheet workbook via ExcelJS) ─────────

  private async generateXlsxPO(items: PoItem[]): Promise<void> {
    this.downloadBtnPO.disabled = true;
    this.downloadBtnPO.textContent = 'Building workbook…';

    try {
      const cats = new Set(items.map((i) => i.category));
      const orderedCats = PendingPoComponent.CATEGORY_ORDER.filter((c) => cats.has(c));

      const wb = new ExcelJS.Workbook();
      wb.creator = 'Pending PO Generator';
      wb.created = new Date();

      const todayLabel = new Date().toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });

      const thin = { style: 'thin' as const, color: { argb: 'FF000000' } };
      const fullBorder = { top: thin, bottom: thin, left: thin, right: thin };
      const fontBase = { name: 'Calibri', size: 11 };

      // NB: the source's alignment objects use `vertical:'center'`, but ExcelJS's actual
      // (and TypeScript-typed) vertical alignment values are 'top'|'middle'|'bottom'|
      // 'distributed'|'justify' — 'center' isn't one of them. Ported as 'middle' below,
      // which is the equivalent value; every other alignment/fill/border value is verbatim.
      const fill = (hex: string) => ({ type: 'pattern' as const, pattern: 'solid' as const, fgColor: { argb: 'FF' + hex } });

      const ws = wb.addWorksheet('UATHAYAM', { views: [{ showGridLines: false }] });
      ws.columns = [
        { width: 55 }, { width: 24 }, { width: 9 }, { width: 10 }, { width: 9 }, { width: 10 },
      ];

      let r = 1;

      ws.mergeCells(r, 1, r, 6);
      const titleCell = ws.getCell(r, 1);
      titleCell.value = `PENDING PO — UATHAYAM`;
      titleCell.font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      titleCell.fill = fill('1C2521');
      ws.getRow(r).height = 26;
      r++;

      ws.mergeCells(r, 1, r, 6);
      const subCell = ws.getCell(r, 1);
      subCell.value = `Generated ${todayLabel}  •  ${items.length} SKUs  •  ${items.reduce((s, i) => s + i.total, 0)} units pending`;
      subCell.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF5A5448' } };
      subCell.alignment = { horizontal: 'center', vertical: 'middle' };
      ws.getRow(r).height = 18;
      r += 2;

      const headerLabels = ['NAME', 'SKU', 'COD', 'PREPAID', 'TOTAL', 'ASRS'];

      orderedCats.forEach((cat, idx) => {
        const sub = items.filter((i) => i.category === cat).sort((a, b) => a.name.localeCompare(b.name));
        const color = this.getOrAssignColor(cat);

        ws.mergeCells(r, 1, r, 6);
        const catCell = ws.getCell(r, 1);
        catCell.value = cat === 'OTHERS' ? 'OTHERS  —  uncategorized, please review' : cat;
        catCell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        catCell.alignment = { horizontal: 'center', vertical: 'middle' };
        catCell.fill = fill(color.hex);
        for (let c = 1; c <= 6; c++) {
          ws.getCell(r, c).border = { bottom: { style: 'medium', color: { argb: 'FF' + color.hex } } };
        }
        ws.getRow(r).height = 20;
        r++;

        headerLabels.forEach((h, c) => {
          const cell = ws.getCell(r, c + 1);
          cell.value = h;
          cell.font = { ...fontBase, bold: true, color: { argb: 'FF' + color.hex } };
          cell.fill = fill(color.tint);
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = fullBorder;
        });
        r++;

        sub.forEach((it, i) => {
          const band = i % 2 === 1;
          const rowVals: (string | number | null)[] = [it.name, it.sku, it.cod, it.prepaid, it.total, null];
          rowVals.forEach((v, c) => {
            const cell = ws.getCell(r, c + 1);
            cell.value = v;
            cell.font = { ...fontBase, bold: true, color: { argb: 'FF1C2521' } };
            cell.alignment = { horizontal: c === 0 ? 'left' : 'center', vertical: 'middle' };
            cell.border = fullBorder;
            if (band) cell.fill = fill('F4F2EC');
          });
          r++;
        });

        const codSum = sub.reduce((s, i) => s + i.cod, 0);
        const prepaidSum = sub.reduce((s, i) => s + i.prepaid, 0);
        const totalSum = sub.reduce((s, i) => s + i.total, 0);
        const totVals: (string | number | null)[] = ['TOTAL', null, codSum, prepaidSum, totalSum, null];
        totVals.forEach((v, c) => {
          const cell = ws.getCell(r, c + 1);
          cell.value = v;
          cell.font = { ...fontBase, bold: true, color: { argb: 'FFFFFFFF' } };
          cell.fill = fill(color.hex);
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
          cell.border = fullBorder;
        });
        r++;

        if (idx !== orderedCats.length - 1) r++;
      });

      ws.getColumn(1).alignment = { horizontal: 'left', vertical: 'middle' };
      ws.views = [{ state: 'frozen', ySplit: 3, showGridLines: false }];

      const ov = wb.addWorksheet('OVERVIEW', { views: [{ showGridLines: false }] });
      ov.columns = [{ width: 28 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 12 }];

      ov.mergeCells(1, 1, 1, 5);
      const ovTitle = ov.getCell(1, 1);
      ovTitle.value = 'PENDING PO — OVERVIEW';
      ovTitle.font = { name: 'Calibri', size: 16, bold: true, color: { argb: 'FFFFFFFF' } };
      ovTitle.fill = fill('1C2521');
      ovTitle.alignment = { horizontal: 'center', vertical: 'middle' };
      ov.getRow(1).height = 26;

      ov.mergeCells(2, 1, 2, 5);
      const ovSub = ov.getCell(2, 1);
      ovSub.value = `Generated ${todayLabel}`;
      ovSub.font = { name: 'Calibri', size: 10, italic: true, color: { argb: 'FF5A5448' } };
      ovSub.alignment = { horizontal: 'center', vertical: 'middle' };

      let orow = 4;
      ['CATEGORY', 'SKUs', 'COD', 'PREPAID', 'TOTAL'].forEach((h, c) => {
        const cell = ov.getCell(orow, c + 1);
        cell.value = h;
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = fill('1C2521');
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.border = fullBorder;
      });
      orow++;

      let grandSku = 0;
      let grandCod = 0;
      let grandPrepaid = 0;
      let grandTotal = 0;
      orderedCats.forEach((cat) => {
        const sub = items.filter((i) => i.category === cat);
        const codSum = sub.reduce((s, i) => s + i.cod, 0);
        const prepaidSum = sub.reduce((s, i) => s + i.prepaid, 0);
        const totalSum = sub.reduce((s, i) => s + i.total, 0);
        grandSku += sub.length;
        grandCod += codSum;
        grandPrepaid += prepaidSum;
        grandTotal += totalSum;
        const color = this.getOrAssignColor(cat);

        const vals: (string | number)[] = [cat === 'OTHERS' ? 'OTHERS (review)' : cat, sub.length, codSum, prepaidSum, totalSum];
        vals.forEach((v, c) => {
          const cell = ov.getCell(orow, c + 1);
          cell.value = v;
          cell.font = { name: 'Calibri', size: 11, bold: c === 0, color: { argb: c === 0 ? 'FFFFFFFF' : 'FF1C2521' } };
          cell.alignment = { horizontal: c === 0 ? 'left' : 'center', vertical: 'middle', indent: c === 0 ? 1 : 0 };
          cell.border = fullBorder;
          if (c === 0) cell.fill = fill(color.hex);
          else cell.fill = fill(color.tint);
        });
        orow++;
      });

      const grandVals: (string | number)[] = ['GRAND TOTAL', grandSku, grandCod, grandPrepaid, grandTotal];
      grandVals.forEach((v, c) => {
        const cell = ov.getCell(orow, c + 1);
        cell.value = v;
        cell.font = { name: 'Calibri', size: 12, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.fill = fill('1C2521');
        cell.alignment = { horizontal: c === 0 ? 'left' : 'center', vertical: 'middle', indent: c === 0 ? 1 : 0 };
        cell.border = fullBorder;
      });

      ov.autoFilter = { from: { row: 4, column: 1 }, to: { row: orow - 1, column: 5 } };

      const buf = await wb.xlsx.writeBuffer();
      const blob = new Blob([buf], { type: 'application/octet-stream' });
      const dateStr = new Date().toISOString().slice(0, 10).split('-').reverse().join('_');
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = `PENDING_PO_FOR_${dateStr}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    } catch (err) {
      this.showStatusPO('Could not build the workbook: ' + (err as Error).message, false);
    } finally {
      this.downloadBtnPO.disabled = false;
      this.downloadBtnPO.textContent = 'Download Pending PO (.xlsx)';
    }
  }
}
