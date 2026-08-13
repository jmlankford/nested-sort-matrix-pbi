/*
 * contextMenu.ts
 * --------------
 * A minimal, framework-free popup menu for the visual's own right-click actions
 * (Copy selection / Copy column). Positioned at a viewport point, dismissed on
 * outside click, Escape, scroll, or after an item is chosen.
 */

export interface ContextMenuItem {
    label: string;
    action: () => void;
}

export class ContextMenu {
    private readonly host: HTMLElement;
    private menu: HTMLElement | null = null;
    private readonly onOutside: (e: MouseEvent) => void;
    private readonly onKey: (e: KeyboardEvent) => void;

    constructor(host: HTMLElement) {
        this.host = host;
        this.onOutside = (e: MouseEvent) => {
            if (this.menu && !this.menu.contains(e.target as Node)) {
                this.close();
            }
        };
        this.onKey = (e: KeyboardEvent) => {
            if (e.key === "Escape") {
                this.close();
            }
        };
    }

    public open(items: ContextMenuItem[], x: number, y: number): void {
        this.close();
        if (items.length === 0) {
            return;
        }
        const menu = document.createElement("div");
        menu.className = "nsm-ctxmenu";
        items.forEach((it) => {
            const el = document.createElement("div");
            el.className = "nsm-ctxmenu-item";
            el.textContent = it.label;
            el.addEventListener("click", (e) => {
                e.stopPropagation();
                this.close();
                it.action();
            });
            menu.appendChild(el);
        });
        // Position within the host; clamp so it stays on screen.
        const hostRect = this.host.getBoundingClientRect();
        let left = x - hostRect.left;
        let top = y - hostRect.top;
        menu.style.left = Math.max(0, left) + "px";
        menu.style.top = Math.max(0, top) + "px";
        menu.style.visibility = "hidden";
        this.host.appendChild(menu);
        // Clamp after measuring.
        const mRect = menu.getBoundingClientRect();
        if (mRect.right > hostRect.right) {
            left = Math.max(0, hostRect.width - mRect.width - 2);
            menu.style.left = left + "px";
        }
        if (mRect.bottom > hostRect.bottom) {
            top = Math.max(0, hostRect.height - mRect.height - 2);
            menu.style.top = top + "px";
        }
        menu.style.visibility = "visible";
        this.menu = menu;

        // Defer listener attach so the opening right-click doesn't immediately close it.
        setTimeout(() => {
            document.addEventListener("mousedown", this.onOutside, true);
            document.addEventListener("keydown", this.onKey, true);
        }, 0);
    }

    public isOpen(): boolean {
        return this.menu !== null;
    }

    public close(): void {
        if (this.menu && this.menu.parentNode) {
            this.menu.parentNode.removeChild(this.menu);
        }
        this.menu = null;
        document.removeEventListener("mousedown", this.onOutside, true);
        document.removeEventListener("keydown", this.onKey, true);
    }
}
