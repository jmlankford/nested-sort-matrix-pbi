/*
 * cfPanel.ts
 * ----------
 * In-visual conditional-formatting panel. Opens on right-click of a value column
 * header, owns all of its own UI + interaction, and calls back into visual.ts to
 * persist changes. Replaces the cfSettings format-pane section entirely.
 *
 * Framework-free DOM (no innerHTML, to satisfy the Power BI lint rules).
 */

export type CfType = "none" | "colorScale" | "rules" | "fieldValue";
export type CfApplyTo = "background" | "font" | "both";
export type RuleOperator =
    | ">="
    | ">"
    | "<="
    | "<"
    | "="
    | "!="
    | "between"
    | "isBlank"
    | "isNotBlank";

export interface CfRule {
    operator: RuleOperator;
    value1: number | string;
    value2?: number | string; // "between" second bound only
    hasAnd?: boolean; // whether a second AND condition exists
    operator2?: RuleOperator; // second AND condition operator
    andValue1?: number | string; // second AND condition value
    andValue2?: number | string; // second AND condition "between" bound
    color: string; // hex string e.g. "#FF6B6B"
}

export interface CfColorScale {
    lowColor: string;
    useMid: boolean;
    midColor: string;
    highColor: string;
    basis: "value" | "percent";
}

export interface CfFieldValue {
    measureSlotIndex: number; // slotIndex of the measure whose value is a hex string
    applyAs: CfApplyTo;
}

export interface CfPanelState {
    cfType: CfType;
    applyTo: CfApplyTo;
    applyToTotals: boolean;
    colorScale: CfColorScale;
    rules: CfRule[];
    defaultColor: string; // fallback color when no rule matches
    fieldValue: CfFieldValue;
}

export type CfPanelSaveCallback = (slotIndex: number, state: CfPanelState) => void;

const HEX_RE = /^#[0-9A-Fa-f]{6}$/;

const OPERATORS: { value: RuleOperator; label: string }[] = [
    { value: ">=", label: ">= (greater or equal)" },
    { value: ">", label: "> (greater than)" },
    { value: "<=", label: "<= (less or equal)" },
    { value: "<", label: "< (less than)" },
    { value: "=", label: "= (equals)" },
    { value: "!=", label: "!= (not equal)" },
    { value: "between", label: "between" },
    { value: "isBlank", label: "is blank" },
    { value: "isNotBlank", label: "is not blank" }
];

const APPLY_TO_OPTIONS: { value: CfApplyTo; label: string }[] = [
    { value: "background", label: "Background" },
    { value: "font", label: "Font" },
    { value: "both", label: "Both" }
];

const TYPE_OPTIONS: { value: CfType; label: string }[] = [
    { value: "none", label: "None" },
    { value: "colorScale", label: "Color Scale" },
    { value: "rules", label: "Rules" },
    { value: "fieldValue", label: "Field value" }
];

const BASIS_OPTIONS: { value: "value" | "percent"; label: string }[] = [
    { value: "value", label: "Value" },
    { value: "percent", label: "Percent of range" }
];

function el<K extends keyof HTMLElementTagNameMap>(
    tag: K,
    className?: string,
    text?: string
): HTMLElementTagNameMap[K] {
    const node = document.createElement(tag);
    if (className) {
        node.className = className;
    }
    if (text !== undefined) {
        node.textContent = text;
    }
    return node;
}

function clone(state: CfPanelState): CfPanelState {
    return {
        cfType: state.cfType,
        applyTo: state.applyTo,
        applyToTotals: state.applyToTotals,
        colorScale: { ...state.colorScale },
        rules: state.rules.map((r) => ({ ...r })),
        defaultColor: state.defaultColor,
        fieldValue: { ...state.fieldValue }
    };
}

export class CfPanel {
    private container: HTMLElement;
    private overlay: HTMLElement;
    private currentSlot: number = -1;
    private measureName: string = "";
    private state: CfPanelState = CfPanel.emptyState();
    private onSave: CfPanelSaveCallback;
    private availableMeasures: { slotIndex: number; displayName: string }[] = [];

    constructor(hostElement: HTMLElement, onSave: CfPanelSaveCallback) {
        this.container = hostElement;
        this.onSave = onSave;
        this.overlay = el("div", "nsm-cf-panel");
        // Catch-all: clicks inside the panel must not reach the host's
        // click-outside-closes listener.
        this.overlay.addEventListener("click", (e: MouseEvent) => e.stopPropagation());
        this.container.appendChild(this.overlay);
    }

    private static emptyState(): CfPanelState {
        return {
            cfType: "none",
            applyTo: "background",
            applyToTotals: false,
            colorScale: {
                lowColor: "#FFFFFF",
                useMid: false,
                midColor: "#FFEB84",
                highColor: "#63BE7B",
                basis: "value"
            },
            rules: [],
            defaultColor: "#CCCCCC",
            fieldValue: { measureSlotIndex: -1, applyAs: "background" }
        };
    }

    public open(slotIndex: number, displayName: string, state: CfPanelState): void {
        this.currentSlot = slotIndex;
        this.measureName = displayName;
        this.state = clone(state);
        this.render();
        this.overlay.classList.add("nsm-cf-panel-open");
        this.applyOverlayStyles();
    }

    public close(): void {
        this.overlay.classList.remove("nsm-cf-panel-open");
        // Clear the fail-safe inline display so the class-based display:none applies.
        this.overlay.style.display = "";
    }

    /** The panel's root element (used by the host's click-outside check). */
    public getOverlayElement(): HTMLElement {
        return this.overlay;
    }

    /**
     * Fail-safe inline styles so a re-render can never drop the visible state,
     * even if the stylesheet class is lost or not yet applied.
     */
    private applyOverlayStyles(): void {
        const s = this.overlay.style;
        s.position = "absolute";
        s.right = "8px";
        s.bottom = "32px";
        s.zIndex = "1000";
        if (this.isOpen()) {
            s.display = "block";
        }
    }

    public setAvailableMeasures(measures: { slotIndex: number; displayName: string }[]): void {
        this.availableMeasures = measures.slice();
    }

    public isOpen(): boolean {
        return this.overlay.classList.contains("nsm-cf-panel-open");
    }

    /** Whether a DOM node lives inside the panel (used for click-outside close). */
    public contains(node: Node | null): boolean {
        return !!node && this.overlay.contains(node);
    }

    // -----------------------------------------------------------------------
    // Rendering.
    // -----------------------------------------------------------------------

    private render(): void {
        this.overlay.textContent = "";

        // Header.
        const header = el("div", "nsm-cf-panel-header");
        header.appendChild(el("span", undefined, "Conditional Formatting — " + this.measureName));
        const closeBtn = el("button", "nsm-cf-close", "✕");
        closeBtn.setAttribute("type", "button");
        closeBtn.setAttribute("aria-label", "Close");
        closeBtn.onclick = (e: MouseEvent) => {
            e.stopPropagation();
            this.cancel();
        };
        header.appendChild(closeBtn);
        this.overlay.appendChild(header);

        // Body.
        const body = el("div", "nsm-cf-panel-body");

        // Apply to.
        body.appendChild(
            this.row(
                "Apply to",
                this.toggleGroup(APPLY_TO_OPTIONS, this.state.applyTo, (v) => {
                    this.state.applyTo = v;
                })
            )
        );

        // Type.
        const typeSelect = this.select(
            TYPE_OPTIONS,
            this.state.cfType,
            (v) => {
                this.state.cfType = v as CfType;
                this.render();
            }
        );
        body.appendChild(this.row("Type", typeSelect));

        // Apply to totals.
        const totalsCheck = el("input") as HTMLInputElement;
        totalsCheck.type = "checkbox";
        totalsCheck.checked = this.state.applyToTotals;
        totalsCheck.onchange = () => {
            this.state.applyToTotals = totalsCheck.checked;
        };
        body.appendChild(this.row("Apply to totals", totalsCheck));

        // Dynamic content.
        if (this.state.cfType === "colorScale") {
            body.appendChild(this.renderColorScale());
        } else if (this.state.cfType === "rules") {
            body.appendChild(this.renderRules());
        } else if (this.state.cfType === "fieldValue") {
            body.appendChild(this.renderFieldValue());
        }

        this.overlay.appendChild(body);

        // Footer.
        const footer = el("div", "nsm-cf-panel-footer");
        const cancelBtn = el("button", "nsm-cf-btn-secondary", "Cancel");
        cancelBtn.setAttribute("type", "button");
        cancelBtn.onclick = (e: MouseEvent) => {
            e.stopPropagation();
            this.cancel();
        };
        const applyBtn = el("button", "nsm-cf-btn-primary", "Apply");
        applyBtn.setAttribute("type", "button");
        applyBtn.onclick = (e: MouseEvent) => {
            e.stopPropagation();
            this.save();
        };
        footer.appendChild(cancelBtn);
        footer.appendChild(applyBtn);
        this.overlay.appendChild(footer);

        this.applyOverlayStyles();
    }

    private renderColorScale(): HTMLElement {
        const wrap = el("div", "nsm-cf-dynamic");
        const cs = this.state.colorScale;

        wrap.appendChild(
            this.row(
                "Low color",
                this.buildColorInput(cs.lowColor, (hex) => {
                    cs.lowColor = hex;
                })
            )
        );

        const midCheck = el("input") as HTMLInputElement;
        midCheck.type = "checkbox";
        midCheck.checked = cs.useMid;
        midCheck.onchange = () => {
            cs.useMid = midCheck.checked;
            this.render();
        };
        wrap.appendChild(this.row("Use midpoint", midCheck));

        const midInput = this.buildColorInput(cs.midColor, (hex) => {
            cs.midColor = hex;
        });
        if (!cs.useMid) {
            midInput.style.opacity = "0.5";
            midInput.style.pointerEvents = "none";
        }
        wrap.appendChild(this.row("Mid color", midInput));

        wrap.appendChild(
            this.row(
                "High color",
                this.buildColorInput(cs.highColor, (hex) => {
                    cs.highColor = hex;
                })
            )
        );

        wrap.appendChild(
            this.row(
                "Basis",
                this.select(BASIS_OPTIONS, cs.basis, (v) => {
                    cs.basis = v as "value" | "percent";
                })
            )
        );

        return wrap;
    }

    private renderRules(): HTMLElement {
        const wrap = el("div", "nsm-cf-dynamic");
        const list = el("div", "nsm-cf-rule-list");

        this.state.rules.forEach((rule, index) => {
            list.appendChild(this.renderRuleRow(rule, index));
        });
        wrap.appendChild(list);

        const addBtn = el("button", "nsm-cf-add-rule", "+ Add rule");
        addBtn.setAttribute("type", "button");
        addBtn.onclick = (e: MouseEvent) => {
            e.stopPropagation();
            this.addRule();
        };
        wrap.appendChild(addBtn);

        wrap.appendChild(
            this.row(
                "Default color",
                this.buildColorInput(this.state.defaultColor, (hex) => {
                    this.state.defaultColor = hex;
                })
            )
        );

        return wrap;
    }

    private renderRuleRow(rule: CfRule, index: number): HTMLElement {
        const rowEl = el("div", "nsm-cf-rule-row");

        // Operator.
        rowEl.appendChild(
            this.select(OPERATORS, rule.operator, (v) => {
                rule.operator = v as RuleOperator;
                this.render();
            })
        );

        const noValue = rule.operator === "isBlank" || rule.operator === "isNotBlank";
        const between = rule.operator === "between";

        if (!noValue) {
            const v1 = el("input") as HTMLInputElement;
            v1.type = "text";
            v1.value = String(rule.value1 ?? "");
            v1.oninput = () => {
                rule.value1 = v1.value;
            };
            rowEl.appendChild(v1);

            if (between) {
                rowEl.appendChild(el("span", undefined, "and"));
                const v2 = el("input") as HTMLInputElement;
                v2.type = "text";
                v2.value = rule.value2 !== undefined ? String(rule.value2) : "";
                v2.oninput = () => {
                    rule.value2 = v2.value;
                };
                rowEl.appendChild(v2);
            }
        }

        // AND toggle button
        const andBtn = el("button", "nsm-cf-and-btn", rule.hasAnd ? "− AND" : "+ AND");
        andBtn.setAttribute("type", "button");
        andBtn.onclick = (e: MouseEvent) => {
            e.stopPropagation();
            rule.hasAnd = !rule.hasAnd;
            if (!rule.hasAnd) {
                delete rule.operator2;
                delete rule.andValue1;
                delete rule.andValue2;
            } else {
                rule.operator2 = "<=";
                rule.andValue1 = 0;
            }
            this.render();
        };
        rowEl.appendChild(andBtn);

        // Second condition row — only when hasAnd is true
        if (rule.hasAnd) {
            const andRow = el("div", "nsm-cf-and-row");
            andRow.appendChild(el("span", "nsm-cf-and-label", "AND"));
            andRow.appendChild(
                this.select(OPERATORS, rule.operator2 ?? "<=", (v) => {
                    rule.operator2 = v as RuleOperator;
                    this.render();
                })
            );
            const noValue2 = rule.operator2 === "isBlank" || rule.operator2 === "isNotBlank";
            const between2 = rule.operator2 === "between";
            if (!noValue2) {
                const av1 = el("input") as HTMLInputElement;
                av1.type = "text";
                av1.value = String(rule.andValue1 ?? "");
                av1.oninput = () => {
                    rule.andValue1 = av1.value;
                };
                andRow.appendChild(av1);
                if (between2) {
                    andRow.appendChild(el("span", undefined, "and"));
                    const av2 = el("input") as HTMLInputElement;
                    av2.type = "text";
                    av2.value = String(rule.andValue2 ?? "");
                    av2.oninput = () => {
                        rule.andValue2 = av2.value;
                    };
                    andRow.appendChild(av2);
                }
            }
            rowEl.appendChild(andRow);
        }

        rowEl.appendChild(
            this.buildColorInput(rule.color, (hex) => {
                rule.color = hex;
            })
        );

        const removeBtn = el("button", "nsm-cf-rule-remove", "✕");
        removeBtn.setAttribute("type", "button");
        removeBtn.setAttribute("aria-label", "Remove rule");
        removeBtn.onclick = (e: MouseEvent) => {
            e.stopPropagation();
            this.removeRule(index);
        };
        rowEl.appendChild(removeBtn);

        return rowEl;
    }

    private renderFieldValue(): HTMLElement {
        const wrap = el("div", "nsm-cf-dynamic");
        const fv = this.state.fieldValue;

        const measureSelect = el("select") as HTMLSelectElement;
        const noneOpt = el("option", undefined, "(none)") as HTMLOptionElement;
        noneOpt.value = "-1";
        measureSelect.appendChild(noneOpt);
        this.availableMeasures.forEach((m) => {
            const opt = el("option", undefined, m.displayName) as HTMLOptionElement;
            opt.value = String(m.slotIndex);
            measureSelect.appendChild(opt);
        });
        measureSelect.value = String(fv.measureSlotIndex);
        measureSelect.onchange = () => {
            fv.measureSlotIndex = parseInt(measureSelect.value, 10);
        };
        wrap.appendChild(this.row("Measure", measureSelect));

        wrap.appendChild(
            this.row(
                "Apply as",
                this.toggleGroup(APPLY_TO_OPTIONS, fv.applyAs, (v) => {
                    fv.applyAs = v;
                })
            )
        );

        return wrap;
    }

    // -----------------------------------------------------------------------
    // Rule list mutation.
    // -----------------------------------------------------------------------

    private addRule(): void {
        this.state.rules.push({ operator: ">=", value1: 0, color: "#FF6B6B" });
        this.render();
    }

    private removeRule(index: number): void {
        if (index >= 0 && index < this.state.rules.length) {
            this.state.rules.splice(index, 1);
            this.render();
        }
    }

    // -----------------------------------------------------------------------
    // Small control builders.
    // -----------------------------------------------------------------------

    private row(labelText: string, control: HTMLElement): HTMLElement {
        const rowEl = el("div", "nsm-cf-row");
        rowEl.appendChild(el("label", undefined, labelText));
        rowEl.appendChild(control);
        return rowEl;
    }

    private select<T extends string>(
        options: { value: T; label: string }[],
        current: T,
        onChange: (value: T) => void
    ): HTMLElement {
        const sel = el("select") as HTMLSelectElement;
        options.forEach((o) => {
            const opt = el("option", undefined, o.label) as HTMLOptionElement;
            opt.value = o.value;
            sel.appendChild(opt);
        });
        sel.value = current;
        sel.onchange = () => onChange(sel.value as T);
        return sel;
    }

    private toggleGroup<T extends string>(
        options: { value: T; label: string }[],
        current: T,
        onChange: (value: T) => void
    ): HTMLElement {
        const group = el("div", "nsm-cf-toggle-group");
        const buttons: HTMLButtonElement[] = [];
        options.forEach((o) => {
            const btn = el("button", "nsm-cf-toggle", o.label);
            btn.setAttribute("type", "button");
            if (o.value === current) {
                btn.classList.add("nsm-cf-toggle-active");
            }
            btn.onclick = (e: MouseEvent) => {
                e.stopPropagation();
                buttons.forEach((b) => b.classList.remove("nsm-cf-toggle-active"));
                btn.classList.add("nsm-cf-toggle-active");
                onChange(o.value);
            };
            buttons.push(btn);
            group.appendChild(btn);
        });
        return group;
    }

    private buildColorInput(value: string, onChange: (hex: string) => void): HTMLElement {
        const wrap = el("div", "nsm-cf-color-input");

        const swatch = el("div", "nsm-cf-color-swatch");
        swatch.style.background = HEX_RE.test(value) ? value : "#ffffff";

        const nativeColor = el("input") as HTMLInputElement;
        nativeColor.type = "color";
        nativeColor.value = HEX_RE.test(value) ? value : "#ffffff";
        nativeColor.style.position = "absolute";
        nativeColor.style.width = "0";
        nativeColor.style.height = "0";
        nativeColor.style.opacity = "0";
        nativeColor.style.pointerEvents = "none";

        const textInput = el("input") as HTMLInputElement;
        textInput.type = "text";
        textInput.value = value;

        const setColor = (hex: string): void => {
            swatch.style.background = hex;
            nativeColor.value = hex;
            onChange(hex);
        };

        swatch.onclick = () => nativeColor.click();
        nativeColor.oninput = () => {
            const hex = nativeColor.value.toUpperCase();
            textInput.value = hex;
            textInput.classList.remove("nsm-cf-hex-invalid");
            setColor(hex);
        };
        textInput.onblur = () => {
            const v = textInput.value.trim();
            if (HEX_RE.test(v)) {
                textInput.classList.remove("nsm-cf-hex-invalid");
                setColor(v);
            } else {
                textInput.classList.add("nsm-cf-hex-invalid");
            }
        };
        textInput.oninput = () => {
            const v = textInput.value.trim();
            if (HEX_RE.test(v)) {
                textInput.classList.remove("nsm-cf-hex-invalid");
                swatch.style.background = v;
            }
        };

        wrap.appendChild(swatch);
        wrap.appendChild(nativeColor);
        wrap.appendChild(textInput);
        return wrap;
    }

    // -----------------------------------------------------------------------
    // Commit.
    // -----------------------------------------------------------------------

    private save(): void {
        this.onSave(this.currentSlot, clone(this.state));
        this.close();
    }

    private cancel(): void {
        this.close();
    }
}
