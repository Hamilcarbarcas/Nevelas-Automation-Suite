import { jqueryFromHtmlLike } from '../foundryCompat.js';

export class DRTypeEditor extends FormApplication {
    constructor(drTypes, onSubmit) {
        super();
        this.drTypes = drTypes;
        this.onSubmit = onSubmit;
        this.materialTypes = pf1.registry.materialTypes;
        this.damageResistances = pf1.config.damageResistances;
        this.availableTypes = ["Custom", ...this.materialTypes.map(m => m.name), ...Object.values(this.damageResistances).sort((a, b) => a.localeCompare(b))];
        this.originalDrTypes = [...drTypes]; 
    }

    static get defaultOptions() {
        return foundry.utils.mergeObject(super.defaultOptions, {
            id: "dr-type-editor",
            title: game.i18n.localize("NAS.forms.drTypeEditor.title"),
            template: "modules/nevelas-automation-suite/src/templates/dr-type-editor.html",
            width: 300,
            height: "auto",
            closeOnSubmit: true
        });
    }

    getData() {
        return {
            drTypes: this.drTypes,
            availableTypes: this.availableTypes
        };
    }

    activateListeners(html) {
        super.activateListeners(html);
        html = jqueryFromHtmlLike(html) ?? html;
        html.find('.add-entry').click(this._onAddType.bind(this));
        html.find('.delete-entry').click(this._onRemoveType.bind(this));
        html.find('#new-type-select').change(this._onNewTypeSelect.bind(this));
        html.find('form').submit(this._onSubmit.bind(this));
    }

    _onAddType(event) {
        event.preventDefault();
        const newTypeSelect = document.getElementById('new-type-select');
        const newType = newTypeSelect.value;
        if (newType && newType !== "Custom") {
            this.drTypes.push(newType);
        } else if (newType === "Custom") {
            const customTypeInput = document.getElementById('custom-type-input');
            const customType = customTypeInput.value.trim();
            if (customType) {
                this.drTypes.push(customType);
                customTypeInput.value = ''; 
            }
        }
        this.render(false);
    }

    _onRemoveType(event) {
        event.preventDefault();
        const index = $(event.currentTarget).data('index');
        this.drTypes.splice(index, 1);
        this.render(false);
    }

    _onNewTypeSelect(event) {
        const selectedType = event.target.value;
        const customTypeInput = document.getElementById('custom-type-input');
        if (selectedType === "Custom") {
            customTypeInput.style.display = 'block';
        } else {
            customTypeInput.style.display = 'none';
        }
    }

    _onSubmit(event) {
        event.preventDefault();

        const inputs = event.currentTarget.querySelectorAll('input[disabled]');
        inputs.forEach(input => input.disabled = false);

        const formData = {};
        const formElements = event.currentTarget.elements;
        for (let element of formElements) {
            if (element.name) {
                formData[element.name] = element.value;
            }
        }

        this._updateObject(event, formData);

        inputs.forEach(input => input.disabled = true);
    }

    async _updateObject(event, formData) {
        const updatedTypes = [];

        for (const key in formData) {
            if (formData.hasOwnProperty(key) && key.startsWith('type')) {
                let value = formData[key].trim();
                if (value) {
                    if (this.availableTypes.includes(value)) {
                        const material = this.materialTypes.find(m => m.name === value);
                        const resistanceKey = Object.keys(this.damageResistances).find(key => this.damageResistances[key] === value);

                        if (material) {
                            updatedTypes.push(material.name); 
                        } else if (resistanceKey) {
                            updatedTypes.push(this.damageResistances[resistanceKey]); 
                        } else {
                            updatedTypes.push(value.capitalize()); 
                        }
                    } else {
                        updatedTypes.push(value.capitalize()); 
                    }
                }
            }
        }

        this.onSubmit(updatedTypes);
        this.close();
    }

    close(options = {}) {
        if (!options.force && !options.submit) {
            this.drTypes.splice(0, this.drTypes.length, ...this.originalDrTypes);
        }
        super.close(options);
    }
}

