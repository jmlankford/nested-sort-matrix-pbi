/*
 * configPanel.ts
 * --------------
 * The runtime configuration overlay. Slides up from the bottom and covers ~65%
 * of the visual height. Three independent tabs (Rows / Values / Column Fields),
 * each listing the active (populated) slots with: a drag handle, the field name,
 * an eye visibility toggle, and a session-rename text input.
 *
 * Reordering uses NATIVE HTML5 drag-and-drop (no external libraries). Reordering
 * within each tab is independent. Buttons: Reset (restore designer defaults) and
 * Done (apply + close).
 *
 * The panel edits a working clone; nothing is applied until Done is pressed.
 * Persistence (order + visibility) is performed by the visual via persistProperties;
 * renames are session-only and are intentionally not persisted.
 */

export type ConfigRole = "row" | "value" | "col";

export interface SlotEntry {
    role: ConfigRole;
    /** Original role slot index — the stable identity across reorders. */
    slotIndex: number;
    /** Original display name from the DataView (placeholder for rename input). */
    originalName: string;
    visible: boolean;
    /** Session-only rename; null means use originalName. */
    rename: string | null;
}

export interface ConfigModel {
    rows: SlotEntry[];
    values: SlotEntry[];
    cols: SlotEntry[];
}

export interface ConfigPanelCallbacks {
    /** Apply the edited model (Done). */
    onApply: (model: ConfigModel) => void;
    /** Produce a fresh default model (Reset). The panel re-renders with it. */
    onReset: () => ConfigModel;
    /** Panel closed without further action hook. */
    onClose: () => void;
}

type TabKey = "rows" | "values" | "cols";

export function cloneModel(model: ConfigModel): ConfigModel {
    return {
        rows: model.rows.map((e) => ({ ...e })),
        values: model.values.map((e) => ({ ...e })),
        cols: model.cols.map((e) => ({ ...e }))
    };
}

export class ConfigPanel {
    private readonly overlay: HTMLElement;
    private readonly panel: HTMLElement;
    private readonly tabBar: HTMLElement;
    private readonly listHost: HTMLElement;
    private working: ConfigModel = { rows: [], values: [], cols: [] };
    private activeTab: TabKey = "rows";
    private open_ = false;
    private dragIndex = -1;

    constructor(container: HTMLElement, private readonly callbacks: ConfigPanelCallbacks) {
        this.overlay = document.createElement("div");
        this.overlay.className = "nsm-config-overlay";
        this.overlay.addEventListener("click", (e) => {
            if (e.target === this.overlay) {
                this.applyAndClose();
            }
        });

        this.panel = document.createElement("div");
        this.panel.className = "nsm-config-panel";

        // Header.
        const header = document.createElement("div");
        header.className = "nsm-config-header";
        const title = document.createElement("div");
        title.className = "nsm-config-title";
        title.textContent = "Setup";
        const closeBtn = document.createElement("button");
        closeBtn.className = "nsm-config-close";
        closeBtn.setAttribute("aria-label", "Close");
        closeBtn.textContent = "✕";
        closeBtn.addEventListener("click", () => this.applyAndClose());
        header.appendChild(title);
        header.appendChild(closeBtn);

        // Tabs.
        this.tabBar = document.createElement("div");
        this.tabBar.className = "nsm-config-tabs";

        // List host.
        this.listHost = document.createElement("div");
        this.listHost.className = "nsm-config-list";

        // Footer.
        const footer = document.createElement("div");
        footer.className = "nsm-config-footer";
        const resetBtn = document.createElement("button");
        resetBtn.className = "nsm-config-btn nsm-config-reset";
        resetBtn.textContent = "Reset";
        resetBtn.addEventListener("click", () => {
            this.working = this.callbacks.onReset();
            this.renderList();
        });
        const doneBtn = document.createElement("button");
        doneBtn.className = "nsm-config-btn nsm-config-done";
        doneBtn.textContent = "Done";
        doneBtn.addEventListener("click", () => this.applyAndClose());
        footer.appendChild(resetBtn);
        footer.appendChild(doneBtn);

        this.panel.appendChild(header);
        this.panel.appendChild(this.tabBar);
        this.panel.appendChild(this.listHost);
        this.panel.appendChild(footer);
        this.overlay.appendChild(this.panel);
        container.appendChild(this.overlay);
    }

    public isOpen(): boolean {
        return this.open_;
    }

    public open(model: ConfigModel): void {
        this.working = cloneModel(model);
        // Default to the first tab that has entries.
        if (this.working.rows.length > 0) {
            this.activeTab = "rows";
        } else if (this.working.values.length > 0) {
            this.activeTab = "values";
        } else {
            this.activeTab = "cols";
        }
        this.renderTabs();
        this.renderList();
        this.open_ = true;
        this.overlay.classList.add("open");
    }

    public close(): void {
        this.open_ = false;
        this.overlay.classList.remove("open");
    }

    public destroy(): void {
        if (this.overlay.parentNode) {
            this.overlay.parentNode.removeChild(this.overlay);
        }
    }

    private applyAndClose(): void {
        this.callbacks.onApply(this.working);
        this.close();
        this.callbacks.onClose();
    }

    private currentList(): SlotEntry[] {
        if (this.activeTab === "rows") {
            return this.working.rows;
        }
        if (this.activeTab === "values") {
            return this.working.values;
        }
        return this.working.cols;
    }

    private renderTabs(): void {
        this.tabBar.textContent = "";
        const defs: { key: TabKey; label: string; count: number }[] = [
            { key: "rows", label: "Rows", count: this.working.rows.length },
            { key: "values", label: "Values", count: this.working.values.length },
            { key: "cols", label: "Column Fields", count: this.working.cols.length }
        ];
        defs.forEach((d) => {
            const tab = document.createElement("button");
            tab.className = "nsm-config-tab";
            tab.textContent = `${d.label} (${d.count})`;
            const disabled = d.count === 0;
            if (disabled) {
                tab.classList.add("disabled");
                tab.disabled = true;
            }
            if (d.key === this.activeTab) {
                tab.classList.add("active");
            }
            tab.addEventListener("click", () => {
                if (disabled) {
                    return;
                }
                this.activeTab = d.key;
                this.renderTabs();
                this.renderList();
            });
            this.tabBar.appendChild(tab);
        });
    }

    private renderList(): void {
        this.listHost.textContent = "";
        const list = this.currentList();

        if (list.length === 0) {
            const empty = document.createElement("div");
            empty.className = "nsm-config-empty";
            empty.textContent = "No fields in this section.";
            this.listHost.appendChild(empty);
            return;
        }

        list.forEach((entry, index) => {
            this.listHost.appendChild(this.renderEntry(entry, index, list));
        });
    }

    private renderEntry(entry: SlotEntry, index: number, list: SlotEntry[]): HTMLElement {
        const row = document.createElement("div");
        row.className = "nsm-config-item";
        row.setAttribute("draggable", "true");
        row.dataset.index = String(index);
        if (!entry.visible) {
            row.classList.add("hidden-field");
        }

        // --- Native HTML5 drag-and-drop ---
        row.addEventListener("dragstart", (e: DragEvent) => {
            this.dragIndex = index;
            row.classList.add("dragging");
            if (e.dataTransfer) {
                e.dataTransfer.effectAllowed = "move";
                e.dataTransfer.setData("text/plain", String(index));
            }
        });
        row.addEventListener("dragend", () => {
            this.dragIndex = -1;
            row.classList.remove("dragging");
            this.listHost
                .querySelectorAll(".nsm-config-item.drop-target")
                .forEach((el) => el.classList.remove("drop-target"));
        });
        row.addEventListener("dragover", (e: DragEvent) => {
            e.preventDefault();
            if (e.dataTransfer) {
                e.dataTransfer.dropEffect = "move";
            }
            row.classList.add("drop-target");
        });
        row.addEventListener("dragleave", () => {
            row.classList.remove("drop-target");
        });
        row.addEventListener("drop", (e: DragEvent) => {
            e.preventDefault();
            row.classList.remove("drop-target");
            const from = this.dragIndex;
            const to = index;
            if (from >= 0 && from !== to) {
                this.reorder(list, from, to);
                this.renderList();
            }
        });

        // Drag handle.
        const handle = document.createElement("span");
        handle.className = "nsm-config-handle";
        handle.textContent = "⋮⋮";
        handle.title = "Drag to reorder";

        // Visibility toggle.
        const eye = document.createElement("button");
        eye.className = "nsm-config-eye";
        eye.title = entry.visible ? "Visible — click to hide" : "Hidden — click to show";
        eye.textContent = entry.visible ? "👁" : "🚫";
        eye.addEventListener("click", (ev) => {
            ev.stopPropagation();
            entry.visible = !entry.visible;
            this.renderList();
        });

        // Name label.
        const name = document.createElement("span");
        name.className = "nsm-config-name";
        name.textContent = entry.originalName || `(slot ${entry.slotIndex + 1})`;

        // Rename input.
        const input = document.createElement("input");
        input.className = "nsm-config-rename";
        input.type = "text";
        input.placeholder = entry.originalName;
        input.value = entry.rename != null ? entry.rename : "";
        input.addEventListener("click", (ev) => ev.stopPropagation());
        input.addEventListener("input", () => {
            const v = input.value.trim();
            entry.rename = v.length > 0 ? v : null;
        });

        row.appendChild(handle);
        row.appendChild(eye);
        row.appendChild(name);
        row.appendChild(input);
        return row;
    }

    private reorder(list: SlotEntry[], from: number, to: number): void {
        const [moved] = list.splice(from, 1);
        list.splice(to, 0, moved);
    }
}
