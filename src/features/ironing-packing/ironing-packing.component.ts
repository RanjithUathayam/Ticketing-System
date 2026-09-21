// src/features/ironing-packing/ironing-packing.component.ts
//
// Ported from the standalone "Enterprise Ironing & Packing Operations Matrix" tool.
// See src/features/legacy-tool-base.ts for the porting pattern this follows.
import { AfterViewInit, Component, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import { LegacyToolBase } from '../legacy-tool-base';

interface ProcessedEntry {
  sku: string;
  desc: string;
  brand: string;
  garClass: string;
  qty: number;
  size: string;
  sizeTagQty: number;
  collarTagQty: number;
  backSupportQty: number;
  chestBandQty: number;
  mrpTagQty: number;
  folderCardQty: number;
  clipQty: number;
}

interface SummaryRow {
  brand: string;
  garClass: string;
  size: string;
  qty: number;
  sizeTag: number;
  collarTag: number;
  back: number;
  chest: number;
  mrp: number;
  folder: number;
  clip: number;
}

interface Totals {
  qty: number;
  sizeTag: number;
  collarTag: number;
  back: number;
  chest: number;
  mrp: number;
  folder: number;
  clip: number;
}

@Component({
  selector: 'app-feature-ironing-packing',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './ironing-packing.component.html',
})
export class IroningPackingComponent extends LegacyToolBase implements AfterViewInit {
  @ViewChild('root', { static: true }) rootRef!: ElementRef<HTMLElement>;

  // ── DOM refs (populated once in ngAfterViewInit, mirrors the original top-level consts) ──
  private ironingUploadCard!: HTMLElement;
  private datatableUploadCard!: HTMLElement;
  private ironingFileInput!: HTMLInputElement;
  private datatableFileInput!: HTMLInputElement;
  private ironingLabel!: HTMLElement;
  private datatableLabel!: HTMLElement;
  private processBtn!: HTMLButtonElement;
  private downloadBtn!: HTMLButtonElement;
  private statusMessageEl!: HTMLElement;
  private kpiVol!: HTMLElement;
  private kpiSizeTags!: HTMLElement;
  private kpiCollarTags!: HTMLElement;
  private kpiBack!: HTMLElement;
  private kpiChest!: HTMLElement;
  private kpiMrp!: HTMLElement;
  private kpiFolder!: HTMLElement;
  private kpiClips!: HTMLElement;
  private summaryTableBody!: HTMLElement;
  private resultTableBody!: HTMLElement;

  // ── State (mirrors the original top-level `let` variables) ──
  private processedEntries: ProcessedEntry[] = [];

  ngAfterViewInit(): void {
    this.bindRoot(this.rootRef);
    this.cacheDomRefs();
    this.wireEvents();
  }

  // ── Setup ──────────────────────────────────────────────────────────────

  private cacheDomRefs(): void {
    this.ironingUploadCard = this.q('ironingUploadCard');
    this.datatableUploadCard = this.q('datatableUploadCard');
    this.ironingFileInput = this.q('ironingFile');
    this.datatableFileInput = this.q('datatableFile');
    this.ironingLabel = this.q('ironingLabel');
    this.datatableLabel = this.q('datatableLabel');
    this.processBtn = this.q('processBtn');
    this.downloadBtn = this.q('downloadBtn');
    this.statusMessageEl = this.q('statusMessage');
    this.kpiVol = this.q('kpi-vol');
    this.kpiSizeTags = this.q('kpi-sizetags');
    this.kpiCollarTags = this.q('kpi-collartags');
    this.kpiBack = this.q('kpi-back');
    this.kpiChest = this.q('kpi-chest');
    this.kpiMrp = this.q('kpi-mrp');
    this.kpiFolder = this.q('kpi-folder');
    this.kpiClips = this.q('kpi-clips');
    this.summaryTableBody = this.q('summaryTable').querySelector('tbody')!;
    this.resultTableBody = this.q('resultTable').querySelector('tbody')!;
  }

  private wireEvents(): void {
    this.ironingUploadCard.addEventListener('click', () => this.ironingFileInput.click());
    this.datatableUploadCard.addEventListener('click', () => this.datatableFileInput.click());
    this.ironingFileInput.addEventListener('click', (e) => e.stopPropagation());
    this.datatableFileInput.addEventListener('click', (e) => e.stopPropagation());
    this.ironingFileInput.addEventListener('change', () => this.updateFileLabel(this.ironingFileInput, this.ironingLabel));
    this.datatableFileInput.addEventListener('change', () => this.updateFileLabel(this.datatableFileInput, this.datatableLabel));
    this.processBtn.addEventListener('click', () => this.processDataFiles());
    this.downloadBtn.addEventListener('click', () => this.downloadStructuredCSV());
  }

  // ── File label ─────────────────────────────────────────────────────────

  private updateFileLabel(input: HTMLInputElement, label: HTMLElement): void {
    const file = input.files?.[0];
    if (file) {
      label.textContent = `✓ ${file.name} successfully loaded`;
      label.classList.add('file-loaded');
    }
  }

  private showStatus(msg: string): void {
    this.statusMessageEl.style.display = 'block';
    this.statusMessageEl.textContent = msg;
  }

  // ── CSV parsing (mirrors the original tool's custom parser) ───────────

  private parseCSV(text: string): string[][] {
    const lines = text.split(/\r\n|\n/);
    return lines
      .map((line) => {
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
      })
      .filter((row) => row.length > 1 || row[0] !== '');
  }

  private extractSize(text: unknown): string {
    const match = String(text || '').match(/\b(36|38|40|42|44|46)\b/);
    return match ? match[0] : 'N/A';
  }

  private readDataStream(file: File): Promise<any[][]> {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
        reader.onload = (e) => {
          const data = new Uint8Array(e.target?.result as ArrayBuffer);
          const wb = XLSX.read(data, { type: 'array' });
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }) as any[][];
          resolve(rows);
        };
        reader.readAsArrayBuffer(file);
      } else {
        reader.onload = () => resolve(this.parseCSV(String(reader.result)));
        reader.readAsText(file);
      }
      reader.onerror = reject;
    });
  }

  // ── Main processing ───────────────────────────────────────────────────

  private async processDataFiles(): Promise<void> {
    const ironFile = this.ironingFileInput.files?.[0];
    const dtFile = this.datatableFileInput.files?.[0];

    if (!ironFile || !dtFile) {
      alert('Operational Fault: Missing operational inputs. Please upload both structural master logs.');
      return;
    }

    this.showStatus('Analyzing and matching records across SKU matrix layers...');

    try {
      const ironRows = await this.readDataStream(ironFile);
      const dtRows = await this.readDataStream(dtFile);

      if (!ironRows.length || !dtRows.length) {
        this.showStatus('Error: One or both uploaded files appear to be empty.');
        return;
      }

      // --- 1. Locate Datatable Header & Build Brand Map ---
      let dtHeaderRowIdx = dtRows.findIndex((row) => row && row.some((cell) => /sku|item/i.test(String(cell))));
      if (dtHeaderRowIdx === -1) dtHeaderRowIdx = 0;

      const dtHeader = dtRows[dtHeaderRowIdx].map((h) => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));

      let dtSkuIdx = dtHeader.findIndex((h) => h.includes('sku') || h.includes('itemcode') || h.includes('item'));
      if (dtSkuIdx === -1) dtSkuIdx = 1;

      let dtBrandIdx = dtHeader.findIndex((h) => h.includes('brand'));
      if (dtBrandIdx === -1) dtBrandIdx = 4;

      const brandLookup: Record<string, string> = {};
      for (let i = dtHeaderRowIdx + 1; i < dtRows.length; i++) {
        const r = dtRows[i];
        if (r && r[dtSkuIdx]) {
          const skuKey = String(r[dtSkuIdx]).toUpperCase().trim();
          brandLookup[skuKey] = String(r[dtBrandIdx] || 'UATHAYAM').toUpperCase().trim();
        }
      }

      // --- 2. Locate Ironing Log Header ---
      let ironHeaderRowIdx = ironRows.findIndex((row) => row && row.some((cell) => /item\s*code|sku/i.test(String(cell))));
      if (ironHeaderRowIdx === -1) ironHeaderRowIdx = 0;

      const ironHeader = ironRows[ironHeaderRowIdx].map((h) => String(h || '').toLowerCase().replace(/[^a-z0-9]/g, ''));

      // Dynamic header detection with positional fallbacks
      let ironSkuIdx = ironHeader.findIndex((h) => h.includes('itemcode') || h.includes('sku') || h.includes('itemno'));
      let ironDescIdx = ironHeader.findIndex((h) => h.includes('description') || h.includes('desc') || h.includes('itemname'));
      let ironQtyIdx = ironHeader.findIndex(
        (h) => h.includes('serviceorderqty') || h.includes('orderqty') || h.includes('qty') || h.includes('quantity'),
      );

      if (ironSkuIdx === -1) ironSkuIdx = 3;
      if (ironDescIdx === -1) ironDescIdx = 4;
      if (ironQtyIdx === -1) ironQtyIdx = 10;

      this.processedEntries = [];
      const summaryGrid: Record<string, SummaryRow> = {};
      const totals: Totals = { qty: 0, sizeTag: 0, collarTag: 0, back: 0, chest: 0, mrp: 0, folder: 0, clip: 0 };

      this.resultTableBody.innerHTML = '';

      // --- 3. Process Ironing Log Rows ---
      for (let i = ironHeaderRowIdx + 1; i < ironRows.length; i++) {
        const r = ironRows[i];
        if (!r || !r[ironSkuIdx]) continue;

        const sku = String(r[ironSkuIdx]).toUpperCase().trim();
        if (!sku || sku === 'UNDEFINED' || sku === 'NULL') continue;

        const desc = String(r[ironDescIdx] || '').trim();

        // Sanitize quantity string (removes commas, currency, spaces, units)
        const rawQty = String(r[ironQtyIdx] || '').replace(/[^0-9.-]+/g, '');
        const qty = parseInt(rawQty, 10) || 0;
        if (qty <= 0) continue;

        const comb = (sku + ' ' + desc).toUpperCase();
        const isTShirt = comb.includes('TSHIRT') || comb.includes('T-SHIRT') || comb.includes('TEE') || comb.includes('POLO');
        const garClass = isTShirt ? 'T-SHIRT' : 'SHIRT';

        const brand = brandLookup[sku] || 'UATHAYAM';
        const size = this.extractSize(sku) !== 'N/A' ? this.extractSize(sku) : this.extractSize(desc);

        // Configured constraints matrix logic
        const sizeTagQty = isTShirt ? 0 : qty;
        const collarTagQty = isTShirt ? 0 : qty;
        const clipQty = isTShirt ? 0 : qty * 2;
        const backSupportQty = qty;
        const chestBandQty = isTShirt ? 0 : qty;
        const mrpTagQty = qty;
        const folderCardQty = isTShirt ? qty : 0;

        this.processedEntries.push({
          sku,
          desc,
          brand,
          garClass,
          qty,
          size,
          sizeTagQty,
          collarTagQty,
          backSupportQty,
          chestBandQty,
          mrpTagQty,
          folderCardQty,
          clipQty,
        });

        const key = `${brand}_${garClass}_${size}`;
        if (!summaryGrid[key]) {
          summaryGrid[key] = { brand, garClass, size, qty: 0, sizeTag: 0, collarTag: 0, back: 0, chest: 0, mrp: 0, folder: 0, clip: 0 };
        }

        summaryGrid[key].qty += qty;
        summaryGrid[key].sizeTag += sizeTagQty;
        summaryGrid[key].collarTag += collarTagQty;
        summaryGrid[key].back += backSupportQty;
        summaryGrid[key].chest += chestBandQty;
        summaryGrid[key].mrp += mrpTagQty;
        summaryGrid[key].folder += folderCardQty;
        summaryGrid[key].clip += clipQty;

        totals.qty += qty;
        totals.sizeTag += sizeTagQty;
        totals.collarTag += collarTagQty;
        totals.back += backSupportQty;
        totals.chest += chestBandQty;
        totals.mrp += mrpTagQty;
        totals.folder += folderCardQty;
        totals.clip += clipQty;

        const tr = document.createElement('tr');
        tr.innerHTML = `
                    <td>${this.escapeHtml(sku)}</td>
                    <td>${this.escapeHtml(desc)}</td>
                    <td class="brand-${brand.toLowerCase()}">${this.escapeHtml(brand)}</td>
                    <td><span class="badge ${isTShirt ? 'badge-tshirt' : 'badge-shirt'}">${this.escapeHtml(garClass)}</span></td>
                    <td><strong>${qty}</strong></td>
                    <td>${sizeTagQty}</td>
                    <td>${collarTagQty}</td>
                    <td>${backSupportQty}</td>
                    <td>${chestBandQty}</td>
                    <td>${mrpTagQty}</td>
                    <td>${folderCardQty}</td>
                    <td>${clipQty}</td>
                `;
        this.resultTableBody.appendChild(tr);
      }

      // Populate Interactive Dash counters
      this.kpiVol.textContent = String(totals.qty);
      this.kpiSizeTags.textContent = String(totals.sizeTag);
      this.kpiCollarTags.textContent = String(totals.collarTag);
      this.kpiBack.textContent = String(totals.back);
      this.kpiChest.textContent = String(totals.chest);
      this.kpiMrp.textContent = String(totals.mrp);
      this.kpiFolder.textContent = String(totals.folder);
      this.kpiClips.textContent = String(totals.clip);

      // Render Summary Grid
      this.summaryTableBody.innerHTML = '';
      Object.keys(summaryGrid)
        .sort()
        .forEach((k) => {
          const s = summaryGrid[k];
          const tr = document.createElement('tr');
          if (s.brand === 'ARISER') tr.classList.add('row-ariser');
          tr.innerHTML = `
                    <td class="brand-${s.brand.toLowerCase()}">${this.escapeHtml(s.brand)}</td>
                    <td><span class="badge ${s.garClass === 'T-SHIRT' ? 'badge-tshirt' : 'badge-shirt'}">${this.escapeHtml(s.garClass)}</span></td>
                    <td><strong>Size ${this.escapeHtml(s.size)}</strong></td>
                    <td><strong>${s.qty}</strong></td>
                    <td>${s.sizeTag}</td>
                    <td>${s.collarTag}</td>
                    <td>${s.back}</td>
                    <td>${s.chest}</td>
                    <td>${s.mrp}</td>
                    <td>${s.folder}</td>
                    <td>${s.clip}</td>
                `;
          this.summaryTableBody.appendChild(tr);
        });

      if (this.processedEntries.length === 0) {
        this.showStatus(
          'Warning: 0 items processed. Please check if your sheet headers match (Item Code / SKU / Qty) and that quantities are greater than zero.',
        );
        this.downloadBtn.disabled = true;
      } else {
        this.showStatus(`Processing run completed: Computed material constraints for ${this.processedEntries.length} items.`);
        this.downloadBtn.disabled = false;
      }
    } catch (e) {
      console.error(e);
      alert('Processing error: ' + (e as Error).message);
    }
  }

  // ── Export to CSV ──────────────────────────────────────────────────────

  private downloadStructuredCSV(): void {
    if (this.processedEntries.length === 0) return;
    let c = 'data:text/csv;charset=utf-8,SKU Code,Description,Brand,Class,Qty,Size Tag,Collar Tag,Back Support,Chest Band,MRP Tag,Folder Card,Clips\n';
    this.processedEntries.forEach((r) => {
      c += `${r.sku},"${r.desc.replace(/"/g, '""')}",${r.brand},${r.garClass},${r.qty},${r.sizeTagQty},${r.collarTagQty},${r.backSupportQty},${r.chestBandQty},${r.mrpTagQty},${r.folderCardQty},${r.clipQty}\n`;
    });
    const l = document.createElement('a');
    l.setAttribute('href', encodeURI(c));
    l.setAttribute('download', 'Production_Packaging_Matrix_Report.csv');
    document.body.appendChild(l);
    l.click();
    document.body.removeChild(l);
  }
}
