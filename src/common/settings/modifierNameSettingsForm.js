import { MODULE } from '../module.js';

export class ModifierNameSettingsForm extends FormApplication {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      id: 'modifier-name-settings-form',
      title: game.i18n.localize("NAS.forms.modifierNames.title"),
      template: `modules/${MODULE.ID}/src/templates/modifier-name-settings-form.html`,
      width: 400,
      height: 'auto',
      closeOnSubmit: true
    });
  }

  getData() {
    const modifierNames = game.settings.get(MODULE.ID, 'modifierNames') || {};
    modifierNames.lesser ||= game.i18n.localize("NAS.forms.modifierNames.labels.lesser");
    modifierNames.minor ||= game.i18n.localize("NAS.forms.modifierNames.labels.minor");
    modifierNames.improved ||= game.i18n.localize("NAS.forms.modifierNames.labels.improved");
    modifierNames.greater ||= game.i18n.localize("NAS.forms.modifierNames.labels.greater");
    modifierNames.major ||= game.i18n.localize("NAS.forms.modifierNames.labels.major");
    modifierNames.supreme ||= game.i18n.localize("NAS.forms.modifierNames.labels.supreme");
    modifierNames.mass ||= game.i18n.localize("NAS.forms.modifierNames.labels.mass");
    modifierNames.communal ||= game.i18n.localize("NAS.forms.modifierNames.labels.communal");
    return modifierNames;
  }

  async _updateObject(event, formData) {
    const modifierNames = {
      lesser: formData.lesser?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.lesser"),
      minor: formData.minor?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.minor"),
      improved: formData.improved?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.improved"),
      greater: formData.greater?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.greater"),
      major: formData.major?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.major"),
      supreme: formData.supreme?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.supreme"),
      mass: formData.mass?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.mass"),
      communal: formData.communal?.trim() || game.i18n.localize("NAS.forms.modifierNames.labels.communal")
    };
    await game.settings.set(MODULE.ID, 'modifierNames', modifierNames);
    ui.notifications.info(game.i18n.localize("NAS.forms.modifierNames.savedInfo"));
  }
}
