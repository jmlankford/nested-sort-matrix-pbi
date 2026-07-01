/*
 * virtualScroller.ts
 * ------------------
 * Viewport-aware row virtualization with DOM node recycling. Only rows in the
 * current viewport plus a configurable buffer (default 20 above/below) are kept
 * in the DOM. Row elements are REUSED across scroll events (a fixed pool that we
 * reposition and rebind) rather than created/destroyed.
 *
 * Layout strategy: the scroll container is `position: relative; overflow: auto`.
 * A zero-content "sizer" element establishes the full scroll height (rowCount *
 * rowHeight) and full content width (so the horizontal scrollbar appears). Row
 * elements are absolutely positioned siblings; because they live inside the
 * scrolling content box they scroll naturally on both axes. Sticky headers are
 * rendered by the renderer outside this container.
 *
 * Uniform row height is assumed (auto -> a fixed default, or the configured
 * fixed px). This keeps index<->offset math O(1) and the virtualization exact.
 */

export type RenderRowCallback<T> = (element: HTMLElement, item: T, index: number) => void;

export class VirtualScroller<T> {
    private readonly container: HTMLElement;
    private readonly sizer: HTMLElement;
    private rowHeight = 28;
    private buffer = 20;
    private items: T[] = [];
    private contentWidth = 0;
    /** Height reserved at the top for the sticky header region. */
    private topOffset = 0;
    private renderRow: RenderRowCallback<T> = () => undefined;

    /** Recyclable row element pool (always attached; hidden when unused). */
    private pool: HTMLElement[] = [];
    private rafHandle: number | null = null;
    private readonly onScrollBound: () => void;
    /** Index of the topmost visible item (for repeat-header decoration). */
    private firstVisibleIndex = 0;

    constructor(container: HTMLElement) {
        this.container = container;
        this.container.classList.add("nsm-vscroll");
        this.container.style.position = "relative";
        this.container.style.overflow = "auto";

        this.sizer = document.createElement("div");
        this.sizer.className = "nsm-vscroll-sizer";
        this.sizer.style.position = "absolute";
        this.sizer.style.top = "0";
        this.sizer.style.left = "0";
        this.sizer.style.width = "1px";
        this.sizer.style.height = "0px";
        this.sizer.style.pointerEvents = "none";
        this.container.appendChild(this.sizer);

        this.onScrollBound = () => this.scheduleUpdate();
        this.container.addEventListener("scroll", this.onScrollBound, { passive: true });
    }

    public configure(rowHeight: number, buffer: number): void {
        this.rowHeight = Math.max(8, Math.floor(rowHeight));
        this.buffer = Math.max(0, Math.floor(buffer));
    }

    public setRenderRow(cb: RenderRowCallback<T>): void {
        this.renderRow = cb;
    }

    public setContentWidth(px: number): void {
        this.contentWidth = Math.max(0, Math.floor(px));
        this.sizer.style.width = this.contentWidth + "px";
    }

    /** Reserve vertical space at the top for a sticky header region. */
    public setTopOffset(px: number): void {
        this.topOffset = Math.max(0, Math.floor(px));
        this.updateSizerHeight();
    }

    private updateSizerHeight(): void {
        this.sizer.style.height = this.topOffset + this.items.length * this.rowHeight + "px";
    }

    public getRowHeight(): number {
        return this.rowHeight;
    }

    /** Index of the topmost currently-visible item. */
    public getFirstVisibleIndex(): number {
        return this.firstVisibleIndex;
    }

    public setItems(items: T[]): void {
        this.items = items || [];
        this.updateSizerHeight();
        this.update();
    }

    public scrollToTop(): void {
        this.container.scrollTop = 0;
        this.update();
    }

    /** Re-render the currently visible window (e.g. after a selection change). */
    public refresh(): void {
        this.update();
    }

    public destroy(): void {
        this.container.removeEventListener("scroll", this.onScrollBound);
        if (this.rafHandle !== null) {
            cancelAnimationFrame(this.rafHandle);
            this.rafHandle = null;
        }
        this.pool.forEach((el) => {
            if (el.parentNode) {
                el.parentNode.removeChild(el);
            }
        });
        this.pool = [];
        if (this.sizer.parentNode) {
            this.sizer.parentNode.removeChild(this.sizer);
        }
    }

    private scheduleUpdate(): void {
        if (this.rafHandle !== null) {
            return;
        }
        this.rafHandle = requestAnimationFrame(() => {
            this.rafHandle = null;
            this.update();
        });
    }

    private ensurePool(size: number): void {
        while (this.pool.length < size) {
            const el = document.createElement("div");
            el.className = "nsm-row";
            el.style.position = "absolute";
            el.style.left = "0";
            el.style.right = "0";
            el.style.height = this.rowHeight + "px";
            el.style.display = "none";
            this.container.appendChild(el);
            this.pool.push(el);
        }
    }

    private update(): void {
        const total = this.items.length;
        const viewportH = this.container.clientHeight || 0;
        const scrollTop = this.container.scrollTop || 0;

        if (total === 0) {
            this.firstVisibleIndex = 0;
            // Hide all pool rows.
            for (let i = 0; i < this.pool.length; i++) {
                this.pool[i].style.display = "none";
            }
            return;
        }

        const effectiveScroll = Math.max(0, scrollTop - this.topOffset);
        const firstVisible = Math.floor(effectiveScroll / this.rowHeight);
        // Record the topmost visible item (clamped) before binding rows.
        this.firstVisibleIndex = Math.max(0, Math.min(total - 1, firstVisible));
        const start = Math.max(0, firstVisible - this.buffer);
        const visibleCount = Math.ceil(viewportH / this.rowHeight) + this.buffer * 2;
        const end = Math.min(total - 1, start + visibleCount - 1);
        const needed = end - start + 1;

        this.ensurePool(needed);

        // Bind/position the rows we need, reusing pool elements in order.
        for (let i = 0; i < needed; i++) {
            const index = start + i;
            const el = this.pool[i];
            el.style.display = "block";
            el.style.height = this.rowHeight + "px";
            el.style.top = this.topOffset + index * this.rowHeight + "px";
            if (this.contentWidth > 0) {
                el.style.minWidth = this.contentWidth + "px";
            }
            this.renderRow(el, this.items[index], index);
        }

        // Hide unused pool elements and return them to the pool. Wipe their cell
        // content ON RETURN (not on next checkout) so stale rows from a previous
        // assignment can never ghost through while a node is reused — e.g. after
        // a group collapse shrinks the row set.
        for (let i = needed; i < this.pool.length; i++) {
            const el = this.pool[i];
            if (el.style.display !== "none") {
                el.textContent = "";
                el.style.display = "none";
            }
        }
    }
}
