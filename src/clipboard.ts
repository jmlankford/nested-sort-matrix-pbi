/*
 * clipboard.ts
 * ------------
 * Clipboard write for a Power BI custom visual, which runs in a sandboxed iframe.
 *
 * Findings / strategy:
 *   1. `document.execCommand("copy")` via a temporary <textarea>, called
 *      SYNCHRONOUSLY inside the user-gesture handler (keydown / menu click), is
 *      the most reliable path in the Power BI sandbox — it is gesture-bound and
 *      needs no Permissions-Policy grant, so it works even when the host iframe
 *      does not advertise `clipboard-write`. We try it first.
 *   2. `navigator.clipboard.writeText()` is the modern API but is async and
 *      requires the `clipboard-write` permission, which the sandbox may withhold
 *      (rejects with NotAllowedError). We use it as a secondary attempt.
 *   3. If both are unavailable, we surface a small dismissible panel with the
 *      text pre-selected so the user can Ctrl+C manually.
 *
 * copyText returns "copied" when a programmatic path succeeded synchronously,
 * or "manual" when the fallback panel was shown. The async modern API, when it
 * is the one that runs, resolves out of band; on rejection it triggers the
 * manual panel.
 */

export type CopyResult = "copied" | "manual";

function execCommandCopy(text: string): boolean {
    try {
        const ta = document.createElement("textarea");
        ta.value = text;
        // Keep it off-screen but focusable/selectable.
        ta.setAttribute("readonly", "");
        ta.style.position = "fixed";
        ta.style.top = "0";
        ta.style.left = "-9999px";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        ta.setSelectionRange(0, text.length);
        const ok = document.execCommand("copy");
        document.body.removeChild(ta);
        return ok;
    } catch {
        return false;
    }
}

/**
 * Copy `text` to the clipboard. Must be called from within a user-gesture
 * handler. `showManual` renders the last-resort manual-copy panel.
 */
export function copyText(text: string, showManual: (text: string) => void): CopyResult {
    // Primary: synchronous execCommand within the gesture.
    if (execCommandCopy(text)) {
        return "copied";
    }
    // Secondary: modern async API (may resolve or reject out of band).
    const nav = navigator as Navigator & { clipboard?: { writeText?: (t: string) => Promise<void> } };
    if (nav.clipboard && typeof nav.clipboard.writeText === "function") {
        try {
            nav.clipboard.writeText(text).then(
                () => undefined,
                () => showManual(text)
            );
            return "copied";
        } catch {
            /* fall through to manual */
        }
    }
    showManual(text);
    return "manual";
}

/**
 * Minimal, dismissible manual-copy panel: a textarea with the content selected
 * and a hint. Escape or the close button removes it. Used only when no
 * programmatic clipboard path is available.
 */
export class ManualCopyPanel {
    private readonly host: HTMLElement;
    private overlay: HTMLElement | null = null;

    constructor(host: HTMLElement) {
        this.host = host;
    }

    public show(text: string): void {
        this.close();
        const overlay = document.createElement("div");
        overlay.className = "nsm-manualcopy-overlay";

        const panel = document.createElement("div");
        panel.className = "nsm-manualcopy-panel";

        const hint = document.createElement("div");
        hint.className = "nsm-manualcopy-hint";
        hint.textContent = "Press Ctrl+C to copy, then Esc to close.";

        const area = document.createElement("textarea");
        area.className = "nsm-manualcopy-text";
        area.value = text;
        area.readOnly = true;

        const close = document.createElement("button");
        close.className = "nsm-manualcopy-close";
        close.textContent = "Close";
        close.addEventListener("click", () => this.close());

        panel.appendChild(hint);
        panel.appendChild(area);
        panel.appendChild(close);
        overlay.appendChild(panel);
        overlay.addEventListener("click", (e) => {
            if (e.target === overlay) {
                this.close();
            }
        });
        this.host.appendChild(overlay);
        this.overlay = overlay;

        // Pre-select for an immediate Ctrl+C.
        area.focus();
        area.select();
        area.setSelectionRange(0, text.length);

        const onKey = (e: KeyboardEvent): void => {
            if (e.key === "Escape") {
                this.close();
                document.removeEventListener("keydown", onKey, true);
            }
        };
        document.addEventListener("keydown", onKey, true);
    }

    public close(): void {
        if (this.overlay && this.overlay.parentNode) {
            this.overlay.parentNode.removeChild(this.overlay);
        }
        this.overlay = null;
    }
}
