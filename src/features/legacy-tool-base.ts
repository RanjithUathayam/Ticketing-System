// src/features/legacy-tool-base.ts
//
// Every ported operations tool started life as a self-contained HTML file that
// manipulated the DOM directly (document.getElementById, innerHTML templating,
// addEventListener) instead of using Angular bindings. Rewriting six large,
// intricate tools into fully idiomatic Angular templates would risk subtly
// breaking their business logic, so instead each tool is ported as a real
// Angular component whose template is the tool's original markup (scoped by
// Angular's emulated view encapsulation) and whose class keeps the original
// imperative logic almost verbatim, translated from free functions into
// methods on `this`.
//
// This base class supplies the handful of DOM helpers every port needs, scoped
// to the component's own root element so IDs from one tool can never collide
// with another tool (or with the surrounding app shell) even though only one
// tool is mounted at a time.
import { ElementRef } from '@angular/core';

const TOAST_STYLE_ID = 'legacy-tool-toast-style';

/** Injects the one global rule the toast helper needs, exactly once, the first time any tool shows a toast. */
function ensureToastStyle(): void {
  if (document.getElementById(TOAST_STYLE_ID)) return;
  const style = document.createElement('style');
  style.id = TOAST_STYLE_ID;
  style.textContent = `
    .legacy-tool-toast {
      position: fixed;
      bottom: 20px;
      left: 50%;
      transform: translateX(-50%);
      background: #1c2422;
      color: #fff;
      padding: 12px 20px;
      border-radius: 8px;
      font-size: 14px;
      font-family: 'Inter', -apple-system, Segoe UI, Roboto, sans-serif;
      z-index: 10000;
      box-shadow: 0 8px 24px rgba(0,0,0,0.2);
      animation: legacyToolToastSlideUp 0.3s ease;
    }
    @keyframes legacyToolToastSlideUp {
      from { transform: translateX(-50%) translateY(20px); opacity: 0; }
      to   { transform: translateX(-50%) translateY(0); opacity: 1; }
    }
  `;
  document.head.appendChild(style);
}

export abstract class LegacyToolBase {
  /** Set once the host element is available (see ngAfterViewInit in each feature component). */
  protected hostEl!: HTMLElement;

  protected bindRoot(ref: ElementRef<HTMLElement>): void {
    this.hostEl = ref.nativeElement;
  }

  /** Scoped replacement for document.getElementById('id'). */
  protected q<T extends HTMLElement = HTMLElement>(id: string): T {
    const el = this.hostEl.querySelector<T>('#' + CSS.escape(id));
    if (!el) throw new Error(`[legacy-tool] element #${id} not found in this tool's template`);
    return el;
  }

  /** Like q(), but returns null instead of throwing when the element is optional/conditional. */
  protected qOpt<T extends HTMLElement = HTMLElement>(id: string): T | null {
    return this.hostEl.querySelector<T>('#' + CSS.escape(id));
  }

  /** Scoped replacement for document.querySelectorAll(selector). */
  protected qa<T extends HTMLElement = HTMLElement>(selector: string): T[] {
    return Array.from(this.hostEl.querySelectorAll<T>(selector));
  }

  /** Scoped replacement for document.querySelector(selector). */
  protected qs<T extends HTMLElement = HTMLElement>(selector: string): T | null {
    return this.hostEl.querySelector<T>(selector);
  }

  protected escapeHtml(s: unknown): string {
    return String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;');
  }

  /** Matches the toast pattern used by the original box-packing / manifest tools. */
  protected showToast(msg: string, durationMs = 3000): void {
    ensureToastStyle();
    const toast = document.createElement('div');
    toast.className = 'legacy-tool-toast';
    toast.textContent = msg;
    document.body.appendChild(toast);
    setTimeout(() => toast.remove(), durationMs);
  }
}
