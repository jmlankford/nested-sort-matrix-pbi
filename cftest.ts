/*
 * cftest.ts
 * ---------
 * Standalone, minimal IVisual ("CF Test") whose ONLY purpose is to probe the
 * conditional-formatting ("fx") dialog. It implements getFormattingModel()
 * (NOT enumerateObjectInstances) and exposes exactly ONE format card with ONE
 * color property whose descriptor sets instanceKind = ConstantOrRule, which is
 * what surfaces the fx button on the color slice.
 *
 * Bind a measure to the "Measure" field well so the fx dialog has data to offer
 * "Format by" options against, then click the fx button on the "Fill color"
 * slice and observe the dialog.
 *
 * This file is built as a SEPARATE temporary .pbiviz — it is not part of the
 * main visual and is never referenced by src/visual.ts.
 */

import powerbi from "powerbi-visuals-api";

import IVisual = powerbi.extensibility.visual.IVisual;
import VisualConstructorOptions = powerbi.extensibility.visual.VisualConstructorOptions;
import VisualUpdateOptions = powerbi.extensibility.visual.VisualUpdateOptions;
import Fill = powerbi.Fill;
import VisualEnumerationInstanceKinds = powerbi.VisualEnumerationInstanceKinds;
import FormattingModel = powerbi.visuals.FormattingModel;
import FormattingCard = powerbi.visuals.FormattingCard;
import FormattingGroup = powerbi.visuals.FormattingGroup;
import SimpleVisualFormattingSlice = powerbi.visuals.SimpleVisualFormattingSlice;
import FormattingDescriptor = powerbi.visuals.FormattingDescriptor;
import FormattingComponent = powerbi.visuals.FormattingComponent;

// Object + property names — must match capabilities.json (the test capabilities).
const CARD_OBJECT = "card1";
const COLOR_PROPERTY = "fillColor";
const DEFAULT_COLOR = "#01B8AA";

export class Visual implements IVisual {
    private readonly root: HTMLElement;
    private readonly swatch: HTMLElement;
    private readonly label: HTMLElement;
    private color: string = DEFAULT_COLOR;

    constructor(options?: VisualConstructorOptions) {
        // The generated plugin passes options as optional; guard for strict null checks.
        if (!options) {
            throw new Error("VisualConstructorOptions are required.");
        }
        this.root = options.element;
        this.root.style.fontFamily = "'Segoe UI', sans-serif";
        this.root.style.fontSize = "12px";
        this.root.style.padding = "10px";
        this.root.style.display = "flex";
        this.root.style.alignItems = "center";
        this.root.style.gap = "10px";

        this.swatch = document.createElement("div");
        this.swatch.style.width = "48px";
        this.swatch.style.height = "48px";
        this.swatch.style.borderRadius = "4px";
        this.swatch.style.border = "1px solid #c8c6c4";

        this.label = document.createElement("div");

        this.root.appendChild(this.swatch);
        this.root.appendChild(this.label);
    }

    public update(options: VisualUpdateOptions): void {
        const dataView =
            options.dataViews && options.dataViews.length ? options.dataViews[0] : undefined;
        const objects = dataView && dataView.metadata ? dataView.metadata.objects : undefined;
        const card = objects ? objects[CARD_OBJECT] : undefined;
        const raw = card ? card[COLOR_PROPERTY] : undefined;

        this.color = DEFAULT_COLOR;
        if (raw && typeof raw === "object") {
            const fill = raw as Fill;
            if (fill.solid && fill.solid.color) {
                this.color = String(fill.solid.color);
            }
        }

        this.swatch.style.background = this.color;
        this.label.textContent =
            "Fill color = " + this.color + ". Open Format pane → 'CF Test Card' → click the fx on 'Fill color'.";
    }

    public getFormattingModel(): FormattingModel {
        // The descriptor ties the slice to capabilities object card1.fillColor.
        // instanceKind: ConstantOrRule is what makes the fx button appear.
        const descriptor: FormattingDescriptor = {
            objectName: CARD_OBJECT,
            propertyName: COLOR_PROPERTY,
            instanceKind: VisualEnumerationInstanceKinds.ConstantOrRule
        };

        const colorSlice: SimpleVisualFormattingSlice = {
            uid: CARD_OBJECT + "-" + COLOR_PROPERTY,
            displayName: "Fill color",
            control: {
                type: FormattingComponent.ColorPicker,
                properties: {
                    descriptor,
                    value: { value: this.color }
                }
            }
        };

        const group: FormattingGroup = {
            uid: CARD_OBJECT + "-group",
            displayName: "",
            slices: [colorSlice]
        };

        const card: FormattingCard = {
            uid: CARD_OBJECT + "-card",
            displayName: "CF Test Card",
            groups: [group],
            revertToDefaultDescriptors: [descriptor]
        };

        return { cards: [card] };
    }
}
