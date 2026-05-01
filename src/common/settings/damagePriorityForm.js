import { MODULE } from '../module.js';
import { jqueryFromHtmlLike } from '../foundryCompat.js';
import { populateDefaultTypes } from '../config.js';
import { DRTypeEditor } from './drTypeEditor.js';

export class DamagePriorityForm extends FormApplication {
    static get defaultOptions() {
        return foundry.utils.mergeObject(super.defaultOptions, {
            id: "damage-priority-form",
            title: game.i18n.localize("NAS.forms.damagePriorityForm.title"),
            template: "modules/nevelas-automation-suite/src/templates/damage-priority-form.html",
            width: 500,
            height: "auto", 
            closeOnSubmit: true
        });
    }

    constructor(...args) {
        super(...args);
        this.originalPriorityLevels = JSON.parse(JSON.stringify(game.settings.get(MODULE.ID, "damageTypePriority"))); 
    }

    getData() {
        const data = super.getData();
        this.priorityLevels = JSON.parse(game.settings.get(MODULE.ID, "damageTypePriority"));
        data.priorityLevels = this.priorityLevels;
        return data;
    }

    activateListeners(html) {
        super.activateListeners(html);
        html = jqueryFromHtmlLike(html) ?? html;
        html.find('.add-row').click(this._onAddRow.bind(this));
        html.find('.delete-row').click(this._onDeleteRow.bind(this));
        html.find('.reset-defaults').click(this._onResetDefaults.bind(this));
        html.find('.edit-row').click(this._onEditRow.bind(this));

        html.find('.edit-row, .delete-row').on('click', function(event) {
            setTimeout(() => {
                event.currentTarget.blur();
            }, 100);
        });
    }

    async _onAddRow(event) {
        event.preventDefault();
        if (!this.priorityLevels) {
            console.error("priorityLevels is undefined in _onAddRow");
            return;
        }
        this.priorityLevels.push([]);
        await game.settings.set(MODULE.ID, "damageTypePriority", JSON.stringify(this.priorityLevels));
        this.render(false);
    }

    async _onDeleteRow(event) {
        event.preventDefault();
        const row = event.currentTarget.closest('tr');
        const index = row.rowIndex - 1; 
        if (!this.priorityLevels) {
            console.error("priorityLevels is undefined in _onDeleteRow");
            return;
        }

        this._showDeleteConfirmationDialog(index);
    }

    async _onResetDefaults(event) {
        event.preventDefault();
        this._showResetConfirmationDialog();
    }

    async _onEditRow(event) {
        event.preventDefault();
        const index = $(event.currentTarget).data('index');
        const drTypes = this.priorityLevels[index];

        new DRTypeEditor(drTypes, async (updatedTypes) => {
            this.priorityLevels[index] = updatedTypes;
            await game.settings.set(MODULE.ID, "damageTypePriority", JSON.stringify(this.priorityLevels));
            this.render(false);
        }).render(true);
    }

    async _updateObject(event, formData) {
        event.preventDefault(); 

        const form = event.currentTarget;
        const disabledFields = form.querySelectorAll('input[disabled]');
        disabledFields.forEach(field => field.disabled = false);

        const priorityLevels = [];
        const formDataUpdated = new FormData(form);

        formDataUpdated.forEach((value, key) => {
            if (key.startsWith('priority')) {
                const index = parseInt(key.split('.')[1]);
                const types = value.split(',')
                    .map(type => type.trim())
                    .filter(type => type !== ''); 
                priorityLevels[index] = types; 
            }
        });

        await game.settings.set(MODULE.ID, "damageTypePriority", JSON.stringify(priorityLevels));

        this.priorityLevels = priorityLevels;

        this.render(false);

        this._promptReload();
    }

    _showDeleteConfirmationDialog(index) {
        new Dialog({
            title: game.i18n.localize("NAS.common.confirmations.deleteTitle"),
            content: `<p>${game.i18n.localize("NAS.common.confirmations.deleteRow")}</p>`,
            buttons: {
                yes: {
                    icon: '<i class="fas fa-check"></i>',
                    label: game.i18n.localize("NAS.common.buttons.yes"),
                    callback: async () => {
                        this.priorityLevels.splice(index, 1);
                        await game.settings.set(MODULE.ID, "damageTypePriority", JSON.stringify(this.priorityLevels));
                        this.render(false);
                    }
                },
                no: {
                    icon: '<i class="fas fa-times"></i>',
                    label: game.i18n.localize("NAS.common.buttons.no")
                }
            },
            default: "no"
        }).render(true);
    }

    _showResetConfirmationDialog() {
        new Dialog({
            title: game.i18n.localize("NAS.common.confirmations.resetTitle"),
            content: `<p>${game.i18n.localize("NAS.common.confirmations.resetDefaults")}</p>`,
            buttons: {
                yes: {
                    icon: '<i class="fas fa-check"></i>',
                    label: game.i18n.localize("NAS.common.buttons.yes"),
                    callback: async () => {
                        await populateDefaultTypes(); 
                        this.priorityLevels = JSON.parse(await game.settings.get(MODULE.ID, "damageTypePriority")); 
                        this.render(true); 
                    }
                },
                no: {
                    icon: '<i class="fas fa-times"></i>',
                    label: game.i18n.localize("NAS.common.buttons.no")
                }
            },
            default: "yes"
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

    close(options = {}) {
        if (!options.force && !options.submit) {
            game.settings.set(MODULE.ID, "damageTypePriority", this.originalPriorityLevels); 
        }
        super.close(options);
    }
}



