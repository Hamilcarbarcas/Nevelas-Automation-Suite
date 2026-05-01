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
        translations.construct ||= "Construct Traits";
        translations.undead ||= "Undead Traits";
        translations.constructClassNames ||= "";
        translations.undeadClassNames ||= "";
        translations.constructRaceNames ||= "";
        translations.undeadRaceNames ||= "";

        return translations;
    }

    async _updateObject(event, formData) {
        const translations = {
            construct: formData.construct.trim() || "Construct Traits",
            undead: formData.undead.trim() || "Undead Traits",
            constructClassNames: String(formData.constructClassNames ?? "").trim(),
            undeadClassNames: String(formData.undeadClassNames ?? "").trim(),
            constructRaceNames: String(formData.constructRaceNames ?? "").trim(),
            undeadRaceNames: String(formData.undeadRaceNames ?? "").trim()
        };

        await game.settings.set(MODULE.ID, "translations", translations);
    }
}



