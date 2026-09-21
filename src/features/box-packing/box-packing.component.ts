// src/features/box-packing/box-packing.component.ts
//
// Ported from the standalone "Re-Iorning Details — Box Packing Entry" tool.
// See src/features/legacy-tool-base.ts for the porting pattern this follows.
import { AfterViewInit, Component, ElementRef, OnDestroy, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import { LegacyToolBase } from '../legacy-tool-base';

interface CatalogItem {
  name: string;
  sku: string;
}

interface BoxLineItem {
  itemName: string;
  sku: string;
  qty: number;
}

interface ClosedBox {
  boxNumber: string;
  items: BoxLineItem[];
}

@Component({
  selector: 'app-feature-box-packing',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './box-packing.component.html',
})
export class BoxPackingComponent extends LegacyToolBase implements AfterViewInit, OnDestroy {
  @ViewChild('root', { static: true }) rootRef!: ElementRef<HTMLElement>;

  // ── DOM refs (populated once in ngAfterViewInit, mirrors the original top-level consts) ──
  private csvFileInput!: HTMLInputElement;
  private importCsvBtn!: HTMLButtonElement;
  private csvStatusEl!: HTMLElement;
  private itemCountDisplay!: HTMLElement;
  private boxNumberInput!: HTMLInputElement;
  private openBoxBtn!: HTMLButtonElement;
  private boxStatusEl!: HTMLElement;
  private itemEntryPanel!: HTMLElement;
  private currentBoxPanel!: HTMLElement;
  private closedSummaryPanel!: HTMLElement;
  private itemSearch!: HTMLInputElement;
  private itemDropdown!: HTMLElement;
  private qtyInput!: HTMLInputElement;
  private addItemBtn!: HTMLButtonElement;
  private selectedChip!: HTMLElement;
  private selectedChipText!: HTMLElement;
  private clearSelectedBtn!: HTMLElement;
  private currentBoxBody!: HTMLElement;
  private currentLineCount!: HTMLElement;
  private currentQtyTotal!: HTMLElement;
  private closeBoxBtn!: HTMLButtonElement;
  private clearCurrentBoxBtn!: HTMLButtonElement;
  private closedBoxesList!: HTMLElement;
  private grandTotalLine!: HTMLElement;
  private exportBtn!: HTMLButtonElement;
  private resetAllBtn!: HTMLButtonElement;

  // ── State (mirrors the original top-level `let` variables) ──
  private itemCatalog: CatalogItem[] = [];
  private currentBoxItems: BoxLineItem[] = [];
  private closedBoxes: ClosedBox[] = [];
  private currentBoxNumber = '';
  private boxOpen = false;
  private pendingSelection: { itemName: string; sku: string } | null = null;
  private filteredList: CatalogItem[] = [];
  private activeIndex = -1;

  private readonly onDocumentClick = (e: MouseEvent): void => {
    const target = e.target as HTMLElement | null;
    if (!target || !target.closest('.field')) {
      this.itemDropdown.classList.remove('open');
    }
  };

  ngAfterViewInit(): void {
    this.bindRoot(this.rootRef);
    this.cacheDomRefs();
    this.wireEvents();
    this.renderCurrentBox();
    this.renderClosedBoxes();
    this.updateGrandTotal();
    this.updateAddButtonState();
  }

  ngOnDestroy(): void {
    document.removeEventListener('click', this.onDocumentClick);
  }

  // ── Setup ──────────────────────────────────────────────────────────────

  private cacheDomRefs(): void {
    this.csvFileInput = this.q('csvFileInput');
    this.importCsvBtn = this.q('importCsvBtn');
    this.csvStatusEl = this.q('csvStatus');
    this.itemCountDisplay = this.q('itemCountDisplay');
    this.boxNumberInput = this.q('boxNumberInput');
    this.openBoxBtn = this.q('openBoxBtn');
    this.boxStatusEl = this.q('boxStatus');
    this.itemEntryPanel = this.q('itemEntryPanel');
    this.currentBoxPanel = this.q('currentBoxPanel');
    this.closedSummaryPanel = this.q('closedSummaryPanel');
    this.itemSearch = this.q('itemSearch');
    this.itemDropdown = this.q('itemDropdown');
    this.qtyInput = this.q('qtyInput');
    this.addItemBtn = this.q('addItemBtn');
    this.selectedChip = this.q('selectedChip');
    this.selectedChipText = this.q('selectedChipText');
    this.clearSelectedBtn = this.q('clearSelected');
    this.currentBoxBody = this.q('currentBoxBody');
    this.currentLineCount = this.q('currentLineCount');
    this.currentQtyTotal = this.q('currentQtyTotal');
    this.closeBoxBtn = this.q('closeBoxBtn');
    this.clearCurrentBoxBtn = this.q('clearCurrentBoxBtn');
    this.closedBoxesList = this.q('closedBoxesList');
    this.grandTotalLine = this.q('grandTotalLine');
    this.exportBtn = this.q('exportBtn');
    this.resetAllBtn = this.q('resetAllBtn');
  }

  private wireEvents(): void {
    this.importCsvBtn.addEventListener('click', () => this.handleImportCsv());
    this.openBoxBtn.addEventListener('click', () => this.handleOpenBox());
    this.closeBoxBtn.addEventListener('click', () => this.handleCloseBox());
    this.clearCurrentBoxBtn.addEventListener('click', () => this.handleClearCurrentBox());
    this.itemSearch.addEventListener('input', (e) => this.handleItemSearchInput(e));
    this.itemSearch.addEventListener('keydown', (e) => this.handleItemSearchKeydown(e as KeyboardEvent));
    this.clearSelectedBtn.addEventListener('click', () => this.clearPendingSelection(true));
    this.qtyInput.addEventListener('input', () => this.updateAddButtonState());
    this.addItemBtn.addEventListener('click', () => this.handleAddItem());
    this.exportBtn.addEventListener('click', () => this.handleExport());
    this.resetAllBtn.addEventListener('click', () => this.handleResetAll());
    document.addEventListener('click', this.onDocumentClick);
  }

  // ── CSV import ─────────────────────────────────────────────────────────

  private handleImportCsv(): void {
    const file = this.csvFileInput.files?.[0];
    if (!file) {
      this.showCsvStatus('Please select a CSV file', 'error');
      return;
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const csv = String(e.target?.result ?? '');
        const lines = csv.split('\n').map(l => l.trim()).filter(l => l.length > 0);
        if (lines.length < 2) {
          this.showCsvStatus('CSV must have at least header + 1 row', 'error');
          return;
        }

        const headers = this.parseCsvLine(lines[0]).map(h => h.toLowerCase().trim());

        const itemIdx = headers.findIndex(h => h.includes('item') || h.includes('product') || h.includes('name'));
        const skuIdx = headers.findIndex(h => h.includes('sku'));

        if (itemIdx === -1 || skuIdx === -1) {
          this.showCsvStatus('CSV must have "Item Name" and "SKU Code" columns', 'error');
          return;
        }

        this.itemCatalog = [];
        let rowCount = 0;
        for (let i = 1; i < lines.length; i++) {
          const cells = this.parseCsvLine(lines[i]);
          if (cells.length <= Math.max(itemIdx, skuIdx)) continue;

          const itemName = cells[itemIdx].trim();
          const sku = cells[skuIdx].trim();
          if (!itemName || !sku) continue;

          this.itemCatalog.push({ name: itemName, sku });
          rowCount++;
        }

        if (rowCount === 0) {
          this.showCsvStatus('No valid data rows found', 'error');
          return;
        }

        this.showCsvStatus(`✓ Loaded ${rowCount} products`, 'success');
        this.itemCountDisplay.textContent = `${rowCount} products loaded`;
        this.openBoxBtn.disabled = false;
        this.itemSearch.focus();
      } catch (err) {
        this.showCsvStatus('Error parsing CSV: ' + (err as Error).message, 'error');
      }
    };
    reader.readAsText(file);
  }

  private parseCsvLine(line: string): string[] {
    const result: string[] = [];
    let current = '';
    let insideQuotes = false;

    for (let i = 0; i < line.length; i++) {
      const char = line[i];
      if (char === '"') {
        insideQuotes = !insideQuotes;
      } else if (char === ',' && !insideQuotes) {
        result.push(current.trim());
        current = '';
      } else {
        current += char;
      }
    }
    result.push(current.trim());
    return result;
  }

  private showCsvStatus(msg: string, type: 'success' | 'error' | 'warning' = 'success'): void {
    this.csvStatusEl.textContent = msg;
    this.csvStatusEl.style.display = 'block';
    this.csvStatusEl.className = 'csv-status';
    if (type === 'error') this.csvStatusEl.classList.add('error');
    if (type === 'warning') this.csvStatusEl.classList.add('warning');
  }

  // ── Box management ─────────────────────────────────────────────────────

  private handleOpenBox(): void {
    if (!this.boxNumberInput.value.trim()) {
      this.showToast('Please enter a box number');
      return;
    }
    this.currentBoxNumber = this.boxNumberInput.value.trim();
    this.boxOpen = true;
    this.boxStatusEl.textContent = 'Open — accepting items';
    this.boxStatusEl.classList.remove('closed');
    this.boxNumberInput.disabled = true;
    this.openBoxBtn.disabled = true;
    this.closeBoxBtn.disabled = false;
    this.itemEntryPanel.style.display = 'block';
    this.currentBoxPanel.style.display = 'block';
    this.itemSearch.focus();
    this.updateAddButtonState();
  }

  private handleCloseBox(): void {
    if (this.currentBoxItems.length === 0) {
      this.showToast('Add at least one item before closing the box.');
      return;
    }
    this.closedBoxes.push({ boxNumber: this.currentBoxNumber, items: this.currentBoxItems.slice() });
    this.boxOpen = false;
    this.currentBoxItems = [];
    this.boxStatusEl.textContent = 'Closed';
    this.boxStatusEl.classList.add('closed');
    this.closeBoxBtn.disabled = true;
    this.boxNumberInput.disabled = false;
    this.openBoxBtn.disabled = false;
    this.itemEntryPanel.style.display = 'none';
    this.boxNumberInput.value = '';
    this.clearPendingSelection(true);
    this.qtyInput.value = '';
    this.renderCurrentBox();
    this.renderClosedBoxes();
    this.updateGrandTotal();
    this.showToast('Box closed: ' + this.currentBoxNumber);
  }

  private handleClearCurrentBox(): void {
    if (!confirm('Clear all items from this box?')) return;
    this.currentBoxItems = [];
    this.renderCurrentBox();
    this.updateGrandTotal();
    this.showToast('Box cleared.');
  }

  // ── Item search & dropdown ─────────────────────────────────────────────

  private handleItemSearchInput(e: Event): void {
    const val = (e.target as HTMLInputElement).value.toLowerCase().trim();
    if (!val) {
      this.itemDropdown.classList.remove('open');
      return;
    }

    this.filteredList = this.itemCatalog
      .filter(item => item.name.toLowerCase().includes(val) || item.sku.toLowerCase().includes(val))
      .slice(0, 50);

    this.renderDropdown();
  }

  private handleItemSearchKeydown(e: KeyboardEvent): void {
    const opts = this.qa<HTMLElement>('.dropdown .opt');
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      this.activeIndex = Math.min(this.activeIndex + 1, opts.length - 1);
      this.highlightActive(opts);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      this.activeIndex = Math.max(this.activeIndex - 1, 0);
      this.highlightActive(opts);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (this.activeIndex >= 0 && opts[this.activeIndex]) {
        opts[this.activeIndex].click();
      } else if (opts.length === 1) {
        opts[0].click();
      }
    } else if (e.key === 'Escape') {
      this.itemDropdown.classList.remove('open');
    }
  }

  private highlightActive(opts: HTMLElement[]): void {
    opts.forEach(o => o.classList.remove('active'));
    if (opts[this.activeIndex]) {
      opts[this.activeIndex].classList.add('active');
      opts[this.activeIndex].scrollIntoView({ block: 'nearest' });
    }
  }

  private renderDropdown(): void {
    this.activeIndex = -1;
    if (this.filteredList.length === 0) {
      this.itemDropdown.innerHTML = '<div class="empty">No matching products found</div>';
      this.itemDropdown.classList.add('open');
      return;
    }
    this.itemDropdown.innerHTML = this.filteredList.map((item, i) => `
      <div class="opt" data-idx="${i}">
        <div class="opt-content">
          <span class="opt-name">${this.escapeHtml(item.name)}</span>
          <span class="opt-sku">${this.escapeHtml(item.sku)}</span>
        </div>
      </div>
    `).join('');
    this.itemDropdown.classList.add('open');
    this.qa<HTMLElement>('.dropdown .opt').forEach(el => {
      el.addEventListener('click', () => {
        const idx = parseInt(el.getAttribute('data-idx') || '0', 10);
        this.selectItem(this.filteredList[idx]);
      });
    });
  }

  private selectItem(item: CatalogItem): void {
    this.itemSearch.value = item.name;
    this.itemDropdown.classList.remove('open');
    this.pendingSelection = { itemName: item.name, sku: item.sku };
    this.showSelectedChip();
    this.updateAddButtonState();
  }

  private showSelectedChip(): void {
    if (!this.pendingSelection) {
      this.selectedChip.classList.remove('show');
      return;
    }
    this.selectedChipText.textContent = `✓ ${this.pendingSelection.itemName} — SKU: ${this.pendingSelection.sku}`;
    this.selectedChip.classList.add('show');
  }

  private clearPendingSelection(clearText: boolean): void {
    this.pendingSelection = null;
    this.selectedChip.classList.remove('show');
    if (clearText) {
      this.itemSearch.value = '';
      this.itemSearch.focus();
    }
    this.updateAddButtonState();
  }

  // ── Qty + add ──────────────────────────────────────────────────────────

  private updateAddButtonState(): void {
    const qtyVal = parseInt(this.qtyInput.value, 10);
    this.addItemBtn.disabled = !(this.boxOpen && this.pendingSelection && qtyVal > 0);
  }

  private handleAddItem(): void {
    const qtyVal = parseInt(this.qtyInput.value, 10);
    if (!this.boxOpen || !this.pendingSelection || !(qtyVal > 0)) return;
    this.currentBoxItems.push({
      itemName: this.pendingSelection.itemName,
      sku: this.pendingSelection.sku,
      qty: qtyVal,
    });
    this.renderCurrentBox();
    this.updateGrandTotal();
    const addedName = this.pendingSelection.itemName;
    this.itemSearch.value = '';
    this.qtyInput.value = '';
    this.clearPendingSelection(false);
    this.itemSearch.focus();
    this.showToast('Added ' + qtyVal + ' × ' + addedName);
  }

  // ── Render current box ─────────────────────────────────────────────────

  private renderCurrentBox(): void {
    if (this.currentBoxItems.length === 0) {
      this.currentBoxBody.innerHTML = '<tr class="empty-row"><td colspan="5">No items added to this box yet</td></tr>';
    } else {
      this.currentBoxBody.innerHTML = this.currentBoxItems.map((it, i) => `
        <tr>
          <td>${i + 1}</td>
          <td>${this.escapeHtml(it.itemName)}</td>
          <td class="sku-cell">${this.escapeHtml(it.sku)}</td>
          <td class="qty-cell">${it.qty}</td>
          <td class="rm"><button class="rm-btn" data-i="${i}" type="button">Remove</button></td>
        </tr>
      `).join('');
      this.qa<HTMLButtonElement>('.rm-btn').forEach(btn => {
        btn.addEventListener('click', () => {
          const i = parseInt(btn.getAttribute('data-i') || '0', 10);
          this.currentBoxItems.splice(i, 1);
          this.renderCurrentBox();
          this.updateGrandTotal();
        });
      });
    }
    this.currentLineCount.textContent = String(this.currentBoxItems.length);
    this.currentQtyTotal.textContent = String(this.currentBoxItems.reduce((a, b) => a + b.qty, 0));
  }

  // ── Render closed boxes ────────────────────────────────────────────────

  private renderClosedBoxes(): void {
    if (this.closedBoxes.length === 0) {
      this.closedSummaryPanel.style.display = 'none';
      return;
    }
    this.closedSummaryPanel.style.display = 'block';
    this.closedBoxesList.innerHTML = this.closedBoxes.map(b => {
      const qty = b.items.reduce((a, x) => a + x.qty, 0);
      return `<span class="box-pill">${this.escapeHtml(b.boxNumber)} — ${b.items.length} lines / ${qty} qty</span>`;
    }).join('');
  }

  // ── Grand total ────────────────────────────────────────────────────────

  private updateGrandTotal(): void {
    const totalBoxes = this.closedBoxes.length + (this.boxOpen ? 1 : 0);
    const totalLines = this.closedBoxes.reduce((a, b) => a + b.items.length, 0) + this.currentBoxItems.length;
    const totalQty =
      this.closedBoxes.reduce((a, b) => a + b.items.reduce((x, y) => x + y.qty, 0), 0) +
      this.currentBoxItems.reduce((a, b) => a + b.qty, 0);
    this.grandTotalLine.textContent =
      `${totalBoxes} box${totalBoxes !== 1 ? 'es' : ''} · ${totalLines} line${totalLines !== 1 ? 's' : ''} · ${totalQty} qty`;
  }

  // ── Reset all ──────────────────────────────────────────────────────────

  private handleResetAll(): void {
    if (!confirm('This clears all boxes and items entered so far. This cannot be undone. Continue?')) return;
    this.currentBoxItems = [];
    this.closedBoxes = [];
    this.currentBoxNumber = '';
    this.boxOpen = false;
    this.boxNumberInput.value = '';
    this.boxNumberInput.disabled = false;
    this.boxStatusEl.textContent = 'Not started';
    this.boxStatusEl.classList.remove('closed');
    this.closeBoxBtn.disabled = true;
    this.openBoxBtn.disabled = this.itemCatalog.length === 0;
    this.itemEntryPanel.style.display = 'none';
    this.currentBoxPanel.style.display = 'none';
    this.clearPendingSelection(true);
    this.qtyInput.value = '';
    this.renderCurrentBox();
    this.renderClosedBoxes();
    this.updateGrandTotal();
    this.showToast('Everything reset.');
  }

  // ── Export to Excel ────────────────────────────────────────────────────

  private handleExport(): void {
    const allBoxes = this.closedBoxes.slice();
    if (this.boxOpen && this.currentBoxItems.length > 0) {
      allBoxes.push({
        boxNumber: this.currentBoxNumber || ('BOX-' + (this.closedBoxes.length + 1)) + ' (open)',
        items: this.currentBoxItems.slice(),
      });
    }
    if (allBoxes.length === 0) {
      this.showToast('Nothing to export yet — add at least one item.');
      return;
    }

    const wb = XLSX.utils.book_new();

    const headerRow = ['Box Number', 'Product Name', 'SKU Code', 'Qty'];
    const dataRows: (string | number)[][] = [];
    allBoxes.forEach(box => {
      box.items.forEach(it => {
        dataRows.push([box.boxNumber, it.itemName, it.sku, it.qty]);
      });
    });
    const aoa = [headerRow, ...dataRows];
    const ws = XLSX.utils.aoa_to_sheet(aoa);

    ws['!cols'] = [{ wch: 18 }, { wch: 42 }, { wch: 24 }, { wch: 10 }];

    const headerStyle = {
      font: { bold: true, color: { rgb: 'FFFFFF' } },
      fill: { fgColor: { rgb: '1C2422' } },
      alignment: { horizontal: 'center', vertical: 'center' },
    };
    for (let c = 0; c < headerRow.length; c++) {
      const cellRef = XLSX.utils.encode_cell({ r: 0, c });
      if (ws[cellRef]) (ws[cellRef] as any).s = headerStyle;
    }

    const skuHighlight = {
      fill: { fgColor: { rgb: 'FFE27A' } },
      font: { bold: true },
      alignment: { horizontal: 'left' },
    };
    for (let r = 1; r <= dataRows.length; r++) {
      const cellRef = XLSX.utils.encode_cell({ r, c: 2 });
      if (ws[cellRef]) (ws[cellRef] as any).s = skuHighlight;
      const qtyRef = XLSX.utils.encode_cell({ r, c: 3 });
      if (ws[qtyRef]) (ws[qtyRef] as any).s = { alignment: { horizontal: 'center' } };
    }

    XLSX.utils.book_append_sheet(wb, ws, 'All Entries');

    const summaryHeader = ['Box Number', 'Total Lines', 'Total Qty'];
    const summaryRows = allBoxes.map(b => [b.boxNumber, b.items.length, b.items.reduce((a, x) => a + x.qty, 0)]);
    const grandQty = summaryRows.reduce((a, r) => a + Number(r[2]), 0);
    const grandLines = summaryRows.reduce((a, r) => a + Number(r[1]), 0);
    const summaryAoa: (string | number)[][] = [summaryHeader, ...summaryRows, ['TOTAL', grandLines, grandQty]];
    const ws2 = XLSX.utils.aoa_to_sheet(summaryAoa);
    ws2['!cols'] = [{ wch: 22 }, { wch: 14 }, { wch: 14 }];
    for (let c = 0; c < 3; c++) {
      const ref = XLSX.utils.encode_cell({ r: 0, c });
      if (ws2[ref]) (ws2[ref] as any).s = headerStyle;
      const lastRef = XLSX.utils.encode_cell({ r: summaryAoa.length - 1, c });
      if (ws2[lastRef]) (ws2[lastRef] as any).s = { font: { bold: true }, fill: { fgColor: { rgb: 'F1EEE2' } } };
    }
    XLSX.utils.book_append_sheet(wb, ws2, 'Box Summary');

    const stamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Re-Iorning_Details_${stamp}.xlsx`);
    this.showToast('Excel file downloaded.');
  }
}
