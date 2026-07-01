/*
 * statusBar.ts
 * ------------
 * The 24px Windows-style status bar pinned to the bottom edge of the visual.
 * Dark background (#1e1e1e — the single permitted hardcoded color), light text.
 *
 * Left to right: ⚙ Setup button, Expand/Collapse All toggle, sort-stack summary
 * (truncated with ellipsis). Right side: row count. Toggleable via format pane;
 * when hidden, height() returns 0 so the renderer reclaims the space.
 */

export interface StatusBarCallbacks {
    onSetup: () => void;
    onToggleExpandAll: (expand: boolean) => void;
}

export interface StatusBarState {
    visible: boolean;
    sortText: string;
    rowCount: number;
    allExpanded: boolean;
}

const BAR_HEIGHT = 24;

export class StatusBar {
    private readonly root: HTMLElement;
    private readonly setupBtn: HTMLElement;
    private readonly expandBtn: HTMLElement;
    private readonly sortLabel: HTMLElement;
    private readonly rowCountLabel: HTMLElement;
    private allExpanded = false;
    private visible = true;

    constructor(container: HTMLElement, private readonly callbacks: StatusBarCallbacks) {
        this.root = document.createElement("div");
        this.root.className = "nsm-statusbar";

        const left = document.createElement("div");
        left.className = "nsm-statusbar-left";

        this.setupBtn = document.createElement("button");
        this.setupBtn.className = "nsm-statusbar-btn";
        this.setupBtn.title = "Open setup / configuration panel";
        this.setupBtn.setAttribute("aria-label", "Open setup panel");
        const gear = document.createElement("span");
        gear.className = "nsm-gear";
        gear.textContent = "⚙";
        const setupLabel = document.createElement("span");
        setupLabel.textContent = "Setup";
        this.setupBtn.appendChild(gear);
        this.setupBtn.appendChild(setupLabel);
        this.setupBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            this.callbacks.onSetup();
        });

        this.expandBtn = document.createElement("button");
        this.expandBtn.className = "nsm-statusbar-btn";
        this.expandBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            this.allExpanded = !this.allExpanded;
            this.updateExpandLabel();
            this.callbacks.onToggleExpandAll(this.allExpanded);
        });

        this.sortLabel = document.createElement("span");
        this.sortLabel.className = "nsm-statusbar-sort";

        left.appendChild(this.setupBtn);
        left.appendChild(this.expandBtn);
        left.appendChild(this.sortLabel);

        const right = document.createElement("div");
        right.className = "nsm-statusbar-right";
        this.rowCountLabel = document.createElement("span");
        this.rowCountLabel.className = "nsm-statusbar-rowcount";
        right.appendChild(this.rowCountLabel);

        this.root.appendChild(left);
        this.root.appendChild(right);
        container.appendChild(this.root);

        this.updateExpandLabel();
    }

    public height(): number {
        return this.visible ? BAR_HEIGHT : 0;
    }

    public render(state: StatusBarState): void {
        this.visible = state.visible;
        this.root.style.display = state.visible ? "flex" : "none";
        if (!state.visible) {
            return;
        }
        this.allExpanded = state.allExpanded;
        this.updateExpandLabel();

        if (state.sortText) {
            this.sortLabel.textContent = "Sorted: " + state.sortText;
            this.sortLabel.title = "Sorted: " + state.sortText;
        } else {
            this.sortLabel.textContent = "";
            this.sortLabel.title = "";
        }

        this.rowCountLabel.textContent = formatCount(state.rowCount) + " rows";
    }

    public destroy(): void {
        if (this.root.parentNode) {
            this.root.parentNode.removeChild(this.root);
        }
    }

    private updateExpandLabel(): void {
        this.expandBtn.textContent = "";
        const chev = document.createElement("span");
        chev.className = "nsm-chev";
        const label = document.createElement("span");
        if (this.allExpanded) {
            chev.textContent = "−";
            label.textContent = "Collapse All";
            this.expandBtn.title = "Collapse all groups";
        } else {
            chev.textContent = "+";
            label.textContent = "Expand All";
            this.expandBtn.title = "Expand all groups";
        }
        this.expandBtn.appendChild(chev);
        this.expandBtn.appendChild(label);
    }
}

function formatCount(n: number): string {
    // Locale-independent thousands grouping to avoid host-locale surprises.
    const sign = n < 0 ? "-" : "";
    let s = Math.abs(Math.floor(n)).toString();
    let out = "";
    while (s.length > 3) {
        out = "," + s.slice(-3) + out;
        s = s.slice(0, -3);
    }
    return sign + s + out;
}
