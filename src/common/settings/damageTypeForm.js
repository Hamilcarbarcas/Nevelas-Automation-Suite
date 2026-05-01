import { MODULE } from '../module.js';
import { jqueryFromHtmlLike } from '../foundryCompat.js';
import { EditDamageType } from './editDamageType.js';
import { updateDamageTypeReferences } from '../migrations.js';

export class DamageTypeFormApplication extends FormApplication {
    constructor(...args) {
        super(...args);
        this.customData = {
            name: "", img: "", category: "", abbr: "", icon: "", isModifier: false, color: "#000000", flag: ""
        };
        this.savedDamageTypes = game.settings.get(MODULE.ID, "customDamageTypes");
    }
  
    static get defaultOptions() {
        return foundry.utils.mergeObject(super.defaultOptions, {
            id: "damage-type-form",
            title: game.i18n.localize("NAS.damageTypes.title"),
            template: `modules/${MODULE.ID}/src/templates/custom-damage-type-form.html`,
            width: 600,
            height: "auto",
            closeOnSubmit: true
        });
    }
  
    getData() {
        return {
            customData: this.customData,
            savedDamageTypes: this.savedDamageTypes,
            hassavedDamageTypes: this.savedDamageTypes.length > 0 
        };
    }
  
    activateListeners(html) {
        super.activateListeners(html);
        html = jqueryFromHtmlLike(html) ?? html;
        html.find('button[name="save"]').click(this._onSave.bind(this));
        html.find('button[name="clear"]').click(this._onClear.bind(this));
        html.find('button.file-picker').click(this._onFilePicker.bind(this));
        html.find('button.edit-btn').click(this._onEdit.bind(this));
        html.find('button.delete-btn').click(this._onDelete.bind(this));
        
        html.find('input[name="custom-category"]').on('focus', this._onCustomCategoryFocus.bind(this));

        html.find('input[name="flag-type"]').on('click', this._onRadioClick.bind(this));
    
        html.find('input[name="custom-category"]').on('focus', this._onCustomCategoryFocus.bind(this));

        html.find('input[name="img-color"]').on('input', () => this._syncImagePreview());
        this._syncImagePreview();
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
  
    async _onSave(event) {
        event.preventDefault(); 
        const form = $(event.currentTarget).parents("form")[0]; 
        const name = form.name.value.capitalize().trim();
        const img = form.img.value.trim();
        let category = form.category.value.trim();
        const customCategory = form["custom-category"].value.trim();
        if (customCategory) {
            category = customCategory.toLowerCase().trim()
        }
        const abbr = form.abbr.value.trim();
        const icon = form.icon.value.trim();
        const color = form.color.value; 
        const imgColor = form["img-color"]?.value ?? "#000000";
        const isModifier = form.isModifier.checked;
  
        const flagType = form["flag-type"].value;
        const flagAbility = form["flag-ability"].value;
        let vsAbility = false;
        let abilities = [];

        if (flagType && !flagAbility) {
            return ui.notifications.error(game.i18n.localize("NAS.forms.damageTypeForm.errors.selectAbility"));
        }

        if (!flagType && flagAbility) {
            ui.notifications.error(game.i18n.localize("NAS.forms.damageTypeForm.errors.flagTypeRequired"));
            return;
        }
  
        if (flagType && flagAbility) {
            vsAbility = true;
            abilities.push(flagAbility); 
        }
    
        const flags = {
            [MODULE.ID]: {
                vsAbility: vsAbility,
                abilities: abilities.join(','), 
                type: flagType || "" 
            }
        };
  
        if (!name) return ui.notifications.error(game.i18n.localize("NAS.common.errors.nameRequired"));
        if (!img && !icon) return ui.notifications.error(game.i18n.localize("NAS.common.errors.imgOrIconRequired"));
        if (!category) return ui.notifications.error(game.i18n.localize("NAS.common.errors.categoryRequired"));
  
        const key = name.toLowerCase(); 
        const imageData = await this._prepareImageValue(img, imgColor);
        const newDamageType = {
            key,
            value: {
                name,
                img: imageData.img,
                imgOriginal: imageData.imgOriginal,
                imgColor,
                category,
                flags,
                namespace: MODULE.ID,
                _id: key,
                abbr,
                icon,
                isModifier,
                color
            }
        };
  
        this.savedDamageTypes.push(newDamageType);
        await game.settings.set(MODULE.ID, "customDamageTypes", this.savedDamageTypes);
        this.render(); 
              
        this._promptReload();
    }
  
    async _onClear(event) {
        event.preventDefault(); 
        const dialog = new Dialog({
            title: game.i18n.localize("NAS.forms.damageTypeForm.clearTitle"),
            content: `<p>${game.i18n.localize("NAS.common.confirmations.clearAll")}</p>`,
            buttons: {
                yes: {
                    icon: '<i class="fas fa-check"></i>',
                    label: game.i18n.localize("NAS.common.buttons.yes"),
                    callback: async () => {
                        await game.settings.set(MODULE.ID, "customDamageTypes", []);
                        this.savedDamageTypes = game.settings.get(MODULE.ID, "customDamageTypes");
                        this.render(); 
                    }
                },
                no: {
                    icon: '<i class="fas fa-times"></i>',
                    label: game.i18n.localize("NAS.common.buttons.no")
                }
            },
            default: "no"
        });
        dialog.render(true);
    }
  
    async _onFilePicker(event) {
        event.preventDefault(); 
        const options = {}; 
        const filePicker = new FilePicker({
            type: event.currentTarget.dataset.type,
            current: this.form.img.value,
            callback: path => {
                this.form.img.value = path; 
                this._syncImagePreview();
            },
            options: options
        });

        if (typeof filePicker.browse === "function") await filePicker.browse();
        else if (typeof filePicker.render === "function") filePicker.render(true);
    }
  
    async _onEdit(event) {
        event.preventDefault(); 
        const index = event.currentTarget.dataset.index; 
        const item = this.savedDamageTypes[index]; 
    
        new EditDamageType(item, index, async (newValues) => {
            const oldKey = String(item.key ?? item.value?._id ?? item.value?.name ?? "").toLowerCase().trim();
            const key = newValues.name.toLowerCase().trim();
    
            this.savedDamageTypes[index] = {
                key,
                value: {
                    ...newValues,
                    flags: {
                        ...item.value.flags, 
                        [MODULE.ID]: {
                            ...item.value.flags[MODULE.ID], 
                            ...newValues.flags[MODULE.ID], 
                        }
                    },
                    namespace: MODULE.ID,
                    _id: key,
                    color: newValues.color 
                }
            };
    
            await game.settings.set(MODULE.ID, "customDamageTypes", this.savedDamageTypes);
            if (oldKey && key && oldKey !== key) {
                await updateDamageTypeReferences(new Map([[oldKey, key]]));
            }
            this.render(); 
    
            this._promptReload();
        }).render(true);
    }
  
    _promptReload() {
        new Dialog({
            title: game.i18n.localize("NAS.common.confirmations.reloadRequired"),
            content: `<p>${game.i18n.localize("NAS.common.confirmations.reload")}</p>`,
            buttons: {
                yes: {
                    icon: '<i class="fas fa-check"></i>',
                    label: game.i18n.localize("NAS.common.buttons.yes"),
                    callback: () => window.location.reload()
                },
                no: {
                    icon: '<i class="fas fa-times"></i>',
                    label: game.i18n.localize("NAS.common.buttons.no")
                }
            },
            default: "yes"
        }).render(true);
    }
  
    async _onDelete(event) {
        event.preventDefault(); 
        const index = event.currentTarget.dataset.index; 
        const item = this.savedDamageTypes[index]; 
        const dialog = new Dialog({
            title: game.i18n.localize("NAS.forms.damageTypeForm.deleteTitle").replace("{value}", item.value.name),
            content: `<p>${game.i18n.localize("NAS.common.confirmations.deleteValue").replace("{value}", item.value.name)}</p>`,
            buttons: {
                yes: {
                    icon: '<i class="fas fa-check"></i>',
                    label: game.i18n.localize("NAS.common.buttons.yes"),
                    callback: async () => {
                        this.savedDamageTypes.splice(index, 1);
                        await game.settings.set(MODULE.ID, "customDamageTypes", this.savedDamageTypes);
                        this.render(); 
                    }
                },
                no: {
                    icon: '<i class="fas fa-times"></i>',
                    label: game.i18n.localize("NAS.common.buttons.no")
                }
            },
            default: "no"
        });
        dialog.render(true);
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
  
    async _updateObject(event, formData) {
    }
}



