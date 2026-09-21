// src/features/manifest-extractor/manifest-extractor.component.ts
//
// Ported from the standalone "Manifest → Order Sheet" extractor tool
// (APP 2 of the combined Operations Hub HTML file). See
// src/features/legacy-tool-base.ts for the porting pattern this follows.
import { AfterViewInit, Component, ElementRef, ViewChild } from '@angular/core';
import { CommonModule } from '@angular/common';
import * as XLSX from 'xlsx';
import * as pdfjsLib from 'pdfjs-dist';
import type { PDFDocumentProxy } from 'pdfjs-dist';
import { LegacyToolBase } from '../legacy-tool-base';

// Same import/worker-path convention already used by LabelPdfService — the
// app ships pdf.worker.min.mjs as a static asset (see angular.json).
pdfjsLib.GlobalWorkerOptions.workerSrc = '/pdf.worker.min.mjs';

interface ManifestRow {
  order: string;
  article: string;
  weight: string;
}

@Component({
  selector: 'app-feature-manifest-extractor',
  standalone: true,
  imports: [CommonModule],
  templateUrl: './manifest-extractor.component.html',
})
export class ManifestExtractorComponent extends LegacyToolBase implements AfterViewInit {
  @ViewChild('root', { static: true }) rootRef!: ElementRef<HTMLElement>;

  // ── DOM refs (populated once in ngAfterViewInit, mirrors the original top-level consts) ──
  private dropzone!: HTMLElement;
  private fileInput!: HTMLInputElement;
  private fileListEl!: HTMLElement;
  private tableHolder!: HTMLElement;
  private emptyMsg!: HTMLElement;
  private countBadge!: HTMLElement;
  private downloadBtn2!: HTMLButtonElement;
  private startNoInput!: HTMLInputElement;
  private dateField!: HTMLInputElement;
  private fromField!: HTMLTextAreaElement;
  private toField!: HTMLTextAreaElement;
  private mobileField!: HTMLInputElement;
  private addRowBtn!: HTMLButtonElement;
  private renumberBtn!: HTMLButtonElement;
  private clearBtn!: HTMLButtonElement;

  // ── State (mirrors the original top-level `let rows = []`) ──
  private rows: ManifestRow[] = [];

  ngAfterViewInit(): void {
    this.bindRoot(this.rootRef);
    this.cacheDomRefs();
    this.initDateField();
    this.wireEvents();
    this.renderTable();
  }

  // ── Setup ──────────────────────────────────────────────────────────────

  private cacheDomRefs(): void {
    this.dropzone = this.q('dropzone');
    this.fileInput = this.q('fileInput');
    this.fileListEl = this.q('fileList');
    this.tableHolder = this.q('tableHolder');
    this.emptyMsg = this.q('emptyMsg');
    this.countBadge = this.q('countBadge');
    this.downloadBtn2 = this.q('downloadBtn2');
    this.startNoInput = this.q('startNo');
    this.dateField = this.q('dateField');
    this.fromField = this.q<HTMLTextAreaElement>('fromField');
    this.toField = this.q<HTMLTextAreaElement>('toField');
    this.mobileField = this.q('mobileField');
    this.addRowBtn = this.q('addRowBtn');
    this.renumberBtn = this.q('renumberBtn');
    this.clearBtn = this.q('clearBtn');
  }

  private initDateField(): void {
    // Set today's date, same as the original hub script.
    const today = new Date();
    this.dateField.value = today.toLocaleDateString('en-GB', { day: '2-digit', month: '2-digit', year: 'numeric' });
  }

  private wireEvents(): void {
    this.dropzone.addEventListener('click', () => this.fileInput.click());
    this.dropzone.addEventListener('dragover', (e) => {
      e.preventDefault();
      this.dropzone.classList.add('drag');
    });
    this.dropzone.addEventListener('dragleave', () => this.dropzone.classList.remove('drag'));
    this.dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      this.dropzone.classList.remove('drag');
      const dt = (e as DragEvent).dataTransfer;
      if (dt) this.handleFiles(dt.files);
    });
    this.fileInput.addEventListener('change', (e) => this.handleFiles((e.target as HTMLInputElement).files));

    this.renumberBtn.addEventListener('click', () => this.renderTable());
    this.startNoInput.addEventListener('change', () => this.renderTable());

    this.addRowBtn.addEventListener('click', () => {
      this.rows.push({ order: '', article: '', weight: '' });
      this.renderTable();
    });

    this.clearBtn.addEventListener('click', () => {
      this.rows = [];
      this.fileListEl.innerHTML = '';
      this.renderTable();
    });

    this.downloadBtn2.addEventListener('click', () => this.handleDownload());
  }

  // ── File handling ──────────────────────────────────────────────────────

  private async handleFiles(fileList: FileList | null): Promise<void> {
    if (!fileList) return;
    const files = Array.from(fileList).filter(
      (f) => f.type === 'application/pdf' || f.name.toLowerCase().endsWith('.pdf')
    );
    for (const file of files) {
      const rowEl = document.createElement('div');
      rowEl.className = 'f';
      rowEl.innerHTML = `<span class="name">${this.escapeHtml(file.name)}</span><span class="status">reading…</span>`;
      this.fileListEl.appendChild(rowEl);
      const statusEl = rowEl.querySelector('.status') as HTMLElement;
      try {
        const extracted = await this.extractFromPdf(file);
        if (extracted.length === 0) {
          statusEl.textContent = 'no matching rows found';
          statusEl.className = 'status err';
        } else {
          this.rows = this.rows.concat(extracted);
          statusEl.textContent = `${extracted.length} row${extracted.length === 1 ? '' : 's'} added`;
          statusEl.className = 'status ok';
        }
        this.renderTable();
      } catch (err) {
        console.error(err);
        statusEl.textContent = 'could not read this file';
        statusEl.className = 'status err';
      }
    }
    this.fileInput.value = '';
  }

  // ── PDF extraction — ported verbatim from the original tool ─────────────
  // (airwaybill / order-number / weight regexes and the x/y-gap based line
  // reconstruction heuristic are the core value of this tool; do not "clean up".)

  private async extractFromPdf(file: File): Promise<ManifestRow[]> {
    const buf = await file.arrayBuffer();
    const pdf: PDFDocumentProxy = await pdfjsLib.getDocument({ data: buf }).promise;
    const found: ManifestRow[] = [];

    const airwaybillRe = /([A-Z]{2}\d{9}IN)/g;
    const orderNearRe = /#?\s?(\d{4,7})/;
    const weightRe = /(\d{1,6}\.000)\b/;

    for (let p = 1; p <= pdf.numPages; p++) {
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      // pdf.js text items don't need strict typing here — we only read the
      // handful of fields the original script used (transform/width/str).
      const items = content.items as any[];

      let text = '';
      let prevX: number | null = null;
      let prevY: number | null = null;
      let prevEndX: number | null = null;
      for (const it of items) {
        const x = it.transform[4];
        const y = it.transform[5];
        const w = it.width || 0;
        const h = Math.abs(it.transform[3]) || Math.abs(it.transform[0]) || 10;

        if (prevY !== null) {
          const sameLine = Math.abs(y - prevY) <= Math.max(2, h * 0.4);
          if (!sameLine) {
            text += '\n';
          } else {
            const gap = x - (prevEndX === null ? x : prevEndX);
            if (gap > h * 0.25) {
              text += ' ';
            }
          }
        }
        text += it.str;
        prevX = x;
        prevY = y;
        prevEndX = x + w;
      }

      const awMatches = [...text.matchAll(airwaybillRe)];
      for (let i = 0; i < awMatches.length; i++) {
        const aw = awMatches[i];
        const awEnd = (aw.index ?? 0) + aw[0].length;
        const nextAwStart = i + 1 < awMatches.length ? (awMatches[i + 1].index ?? text.length) : text.length;
        const segment = text.slice(awEnd, Math.min(nextAwStart, awEnd + 400));

        const ordMatch = segment.match(orderNearRe);
        if (!ordMatch) continue;

        const wtMatch = segment.match(weightRe);
        found.push({
          order: '#' + ordMatch[1],
          article: aw[1],
          weight: wtMatch ? String(parseInt(wtMatch[1], 10)) : '',
        });
      }
    }

    const seen = new Set<string>();
    return found.filter((r) => {
      if (seen.has(r.article)) return false;
      seen.add(r.article);
      return true;
    });
  }

  // ── escapeHtml — overrides LegacyToolBase's version. The original app2
  //    script also escapes quotes because values are interpolated into
  //    `value="..."` attributes in renderTable(); the base class helper only
  //    escapes &/</> which would be insufficient here. ────────────────────
  protected override escapeHtml(s: unknown): string {
    const str = String(s ?? '');
    const map: Record<string, string> = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };
    return str.replace(/[&<>"']/g, (c) => map[c]);
  }

  // ── Render extracted-data table ──────────────────────────────────────────

  private renderTable(): void {
    this.countBadge.innerHTML = `<b>${this.rows.length}</b> row${this.rows.length === 1 ? '' : 's'} extracted`;
    this.downloadBtn2.disabled = this.rows.length === 0;

    if (this.rows.length === 0) {
      this.tableHolder.innerHTML = '';
      this.tableHolder.appendChild(this.emptyMsg);
      return;
    }

    const start = parseInt(this.startNoInput.value, 10) || 1;
    let html = `<table><thead><tr>
            <th>S.NO</th><th>ORDER NO</th><th>ARTICLE NO</th><th>WEIGHT</th><th></th>
        </tr></thead><tbody>`;
    this.rows.forEach((r, idx) => {
      const flagged = !r.weight;
      html += `<tr data-idx="${idx}" class="${flagged ? 'flag' : ''}">
                <td class="rowno">${start + idx}</td>
                <td><input type="text" value="${this.escapeHtml(r.order)}" data-field="order"></td>
                <td><input type="text" value="${this.escapeHtml(r.article)}" data-field="article"></td>
                <td class="num"><input type="text" value="${this.escapeHtml(r.weight)}" data-field="weight" style="text-align:right"></td>
                <td class="del" data-remove="${idx}">✕</td>
            </tr>`;
    });
    html += '</tbody></table>';
    this.tableHolder.innerHTML = html;

    this.tableHolder.querySelectorAll('input').forEach((inp) => {
      inp.addEventListener('input', (e) => {
        const target = e.target as HTMLInputElement;
        const tr = target.closest('tr') as HTMLElement;
        const idx = parseInt(tr.dataset['idx'] || '0', 10);
        const field = target.dataset['field'] as keyof ManifestRow;
        this.rows[idx][field] = target.value;
      });
    });
    this.tableHolder.querySelectorAll('[data-remove]').forEach((el) => {
      el.addEventListener('click', () => {
        const idx = parseInt((el as HTMLElement).dataset['remove'] || '0', 10);
        this.rows.splice(idx, 1);
        this.renderTable();
      });
    });
  }

  // ── Download styled, merged-cell order-sheet XLSX — ported verbatim ─────

  private handleDownload(): void {
    const start = parseInt(this.startNoInput.value, 10) || 1;
    const date = this.dateField.value;
    const fromLines = this.fromField.value.split('\n');
    const toLines = this.toField.value.split('\n');
    const mobile = this.mobileField.value;
    const addrLines = Math.max(fromLines.length, toLines.length);

    const aoa: (string | number)[][] = [];
    aoa.push([date, '', '', '']);
    aoa.push(['From', '', 'To:', '']);
    for (let i = 0; i < addrLines; i++) {
      aoa.push([fromLines[i] || '', '', toLines[i] || '', '']);
    }
    aoa.push([mobile, '', ' ', '']);
    aoa.push(['', '', '', '']);
    aoa.push(['S.NO', 'ORDER NO ', 'ARTICLE NO', 'WEIGHT ']);
    this.rows.forEach((r, idx) => {
      aoa.push([start + idx, r.order, r.article, r.weight ? parseInt(r.weight, 10) : '']);
    });

    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws['!cols'] = [{ wch: 13 }, { wch: 13 }, { wch: 13 }, { wch: 13 }];

    const merges: any[] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 3 } },
      { s: { r: 1, c: 0 }, e: { r: 1, c: 1 } },
      { s: { r: 1, c: 2 }, e: { r: 1, c: 3 } },
    ];
    for (let i = 0; i < addrLines; i++) {
      const r = 2 + i;
      merges.push({ s: { r, c: 0 }, e: { r, c: 1 } });
      merges.push({ s: { r, c: 2 }, e: { r, c: 3 } });
    }
    const mobileRow = 2 + addrLines;
    merges.push({ s: { r: mobileRow, c: 0 }, e: { r: mobileRow, c: 1 } });
    merges.push({ s: { r: mobileRow, c: 2 }, e: { r: mobileRow, c: 3 } });
    const blankRow = mobileRow + 1;
    merges.push({ s: { r: blankRow, c: 0 }, e: { r: blankRow, c: 3 } });
    ws['!merges'] = merges;

    const thinBorder = {
      top: { style: 'thin' },
      bottom: { style: 'thin' },
      left: { style: 'thin' },
      right: { style: 'thin' },
    };

    const range = XLSX.utils.decode_range(ws['!ref'] as string);
    for (let R = range.s.r; R <= range.e.r; R++) {
      for (let C = range.s.c; C <= range.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        const cell = ws[addr] || { t: 's', v: '' };
        let align = 'center';
        let sz = 14;
        if (R === 0) {
          align = 'right';
          sz = 12;
        } else if (R >= 1 && R <= mobileRow) {
          align = C === 0 || C === 2 ? 'left' : 'left';
        }
        (cell as any).s = {
          font: { bold: true, sz },
          alignment: { horizontal: align, vertical: 'center', wrapText: true },
          border: thinBorder,
        };
        ws[addr] = cell;
      }
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, 'order_sheet_extracted.xlsx', { cellStyles: true } as any);
  }
}
