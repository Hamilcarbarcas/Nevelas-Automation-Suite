import { MODULE } from '../module.js';

export class TranslationForm extends FormApplication {
    static get defaultOptions() {
        return foundry.utils.mergeObject(super.defaultOptions, {
            id: "translation-form",
            title: game.i18n.localize("NAS.forms.translationForm.title"),
            template: "modules/nevelas-automation-suite/src/templates/translation-form.html",
            width: 400,
            height: "auto",
            closeOnSubmit: true
        });
    }

    getData() {
        const translations = game.settings.get(MODULE.ID, "translations") || {};
        translations.hardness ||= "Hardness";
        translations.construct ||= "Construct Traits";
        translations.undead ||= "Undead Traits";

        return translations;
    }

    async _updateObject(event, formData) {
        const translations = {
            hardness: formData.hardness.trim() || "Hardness",
            construct: formData.construct.trim() || "Construct Traits",
            undead: formData.undead.trim() || "Undead Traits"
        };

        await game.settings.set(MODULE.ID, "translations", translations);
    }
}



