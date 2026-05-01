import { MODULE } from '../module.js';
import { jqueryFromHtmlLike } from '../foundryCompat.js';

export class EditDamageType extends FormApplication {
    constructor(item, index, onSubmit) {
        super();
        this.item = item;
        this.index = index;
        this.onSubmitCallback = onSubmit;

        this.initialImg = item.value.imgOriginal || item.value.img;
        this.initialIcon = item.value.icon;
    }

    static get defaultOptions() {
        return foundry.utils.mergeObject(super.defaultOptions, {
            id: "edit-damage-type",
            title: game.i18n.localize("NAS.forms.drTypeEditor.title"),
            template: "modules/nevelas-automation-suite/src/templates/damage-type-editor.html",
            width: 400,
            height: "auto",
            closeOnSubmit: true
        });
    }

    getData() {
        let categoryDisplay = this.item.value.category;
    
        if (!["physical", "energy", "misc"].includes(this.item.value.category)) {
            categoryDisplay = this.item.value.category.capitalize();
        }
    
        return {
            item: {
                ...this.item.value,
                imgInput: this.item.value.imgOriginal || this.item.value.img || "",
                imgPreview: this.item.value.img || this.item.value.imgOriginal || "",
                imgColor: this.item.value.imgColor || this.item.value.color || "#000000",
                category: categoryDisplay  
            },
            moduleId: MODULE.ID
        };
    }

    activateListeners(html) {
        super.activateListeners(html);
        html = jqueryFromHtmlLike(html) ?? html;
        this.selectedRadio = html.find('input[name="flag-type"]:checked');
    
        html.find(`input[name="flag-type"][value="${this.item.value.flags[MODULE.ID]?.type}"]`).prop('checked', true);
    
        html.find(`select[name="flag-ability"]`).val(this.item.value.flags[MODULE.ID]?.abilities || '');
    
        html.find('button.file-picker').click(this._onFilePicker.bind(this));
        html.find('input[name="custom-category"]').on('focus', this._onCustomCategoryFocus.bind(this));
        html.find('input[name="flag-type"]').on('click', this._onRadioClick.bind(this));
        html.find('input[name="img-color"]').on('input', () => this._syncImagePreview());
    
        html.find('button[name="save"]').on('click', (event) => {
            this._onSubmit(event);
        });

        this._syncImagePreview();
    }

    async _onFilePicker(event) {
        event.preventDefault();
        const options = {};
        const filePicker = new FilePicker({
            type: event.currentTarget.dataset.type,
            current: this.form.img.value,
            callback: (path) => {
                this.form.img.value = path;
                this._syncImagePreview();
            },
            options: options
        });

        if (typeof filePicker.browse === "function") await filePicker.browse();
        else if (typeof filePicker.render === "function") filePicker.render(true);
    }

    _onCustomCategoryFocus(event) {
        $('input[name="category"]').prop('checked', false);
    }

    _onRadioClick(event) {
        const clickedRadio = $(event.currentTarget);
        if (this.selectedRadio && this.selectedRadio[0] === clickedRadio[0]) {
            clickedRadio.prop('checked', false);
            this.selectedRadio = null;
        } else {
            this.selectedRadio = clickedRadio;
        }
    }

    async _onSubmit(event, options = {}) {
        event.preventDefault();
        const formData = new FormData(this.element.find('form')[0]);

        const name = formData.get("name").trim().capitalize();
        const img = formData.get("img").trim();
        let category = formData.get("category");
        const customCategory = formData.get("custom-category").trim();
        if (!["physical", "energy", "misc"].includes(category)) {
            category = customCategory.toLowerCase().trim()
        }

        const abbr = formData.get("abbr").trim();
        const icon = formData.get("icon").trim();
        const color = formData.get("color").trim();
        const imgColor = (formData.get("img-color") || "#000000").trim();
        const isModifier = formData.get("isModifier") === "on";

        const flagType = formData.get("flag-type");
        const flagAbility = formData.get("flag-ability");

        let vsAbility = false;
        let abilities = "";

        if (!name) {
            ui.notifications.error(game.i18n.localize("NAS.common.errors.nameRequired"));
            return;
        }
        if (!img && !icon) {
            ui.notifications.error(game.i18n.localize("NAS.common.errors.imgOrIconRequired"));
            return;
        }
        if (!category) {
            ui.notifications.error(game.i18n.localize("NAS.common.errors.categoryRequired"));
            return;
        }
        if (flagType && !flagAbility) {
            ui.notifications.error(game.i18n.localize("NAS.forms.damageTypeForm.errors.selectAbility"));
            return;
        }

        if (!flagType && flagAbility) {
            ui.notifications.error(game.i18n.localize("NAS.forms.damageTypeForm.errors.flagTypeRequired"));
            return;
        }

        if (flagType && flagAbility) {
            vsAbility = true;
            abilities = flagAbility;
        }

        const updatedFlags = {
            ...this.item.value.flags,
            [MODULE.ID]: {
                vsAbility: vsAbility,
                abilities: abilities,
                type: flagType || ""
            }
        };

        const imageData = await this._prepareImageValue(img, imgColor);
        const resolvedImg = (this.initialImg && img === this.initialImg && icon && !this.initialIcon) ? "" : imageData.img;

        const updatedItem = {
            name,
            img: resolvedImg,
            imgOriginal: resolvedImg ? imageData.imgOriginal : "",
            imgColor,
            category,
            abbr,
            icon: (this.initialIcon && icon === this.initialIcon && img && !this.initialImg) ? "" : icon, 
            color,
            isModifier,
            flags: updatedFlags
        };

        this.item.value = updatedItem;

        this.onSubmitCallback(updatedItem);
        await this.onSubmitCallback(updatedItem);

        this.close();
    }

    async _updateObject(event, formData) {
    }

    _isSvgPath(path) {
        if (!path || typeof path !== "string") return false;
        if (/^data:image\/svg\+xml/i.test(path)) return true;
        return /\.svg(?:$|[?#])/i.test(path);
    }

    _injectSvgTint(svgText, color) {
        if (!svgText || typeof svgText !== "string") return null;
        const match = svgText.match(/<svg\b[^>]*>/i);
        if (!match) return null;

        const style = `<style>*:not([fill="none"]):not([stroke="none"]){fill:${color} !important;stroke:${color} !important;}</style>`;
        return svgText.replace(match[0], `${match[0]}${style}`);
    }

    async _createTintedSvgDataUrl(path, color) {
        try {
            if (/^data:image\/svg\+xml/i.test(path)) return path;
            const response = await fetch(path);
            if (!response.ok) return null;
            const svgText = await response.text();
            const tintedSvg = this._injectSvgTint(svgText, color);
            if (!tintedSvg) return null;
            return `data:image/svg+xml;charset=utf-8,${encodeURIComponent(tintedSvg)}`;
        } catch (_err) {
            return null;
        }
    }

    async _prepareImageValue(path, color) {
        if (!path) return { img: "", imgOriginal: "" };
        if (!this._isSvgPath(path) || /^data:image\/svg\+xml/i.test(path)) {
            return { img: path, imgOriginal: "" };
        }

        const tinted = await this._createTintedSvgDataUrl(path, color);
        if (!tinted) return { img: path, imgOriginal: "" };
        return { img: tinted, imgOriginal: path };
    }

    async _syncImagePreview() {
        const preview = this.element?.find?.("[data-image-preview]")?.[0];
        const colorPicker = this.element?.find?.('input[name="img-color"]')?.[0];
        if (!preview) return;

        const path = (this.form?.img?.value ?? "").trim();
        const color = this.form?.["img-color"]?.value ?? "#000000";
        const isSvg = this._isSvgPath(path);

        if (colorPicker) colorPicker.style.display = isSvg ? "inline-block" : "none";

        if (!path) {
            preview.removeAttribute("src");
            preview.style.display = "none";
            return;
        }

        let previewSrc = path;
        if (isSvg) {
            const tinted = await this._createTintedSvgDataUrl(path, color);
            if (tinted) previewSrc = tinted;
        }

        preview.src = previewSrc;
        preview.style.display = "block";
    }
}



