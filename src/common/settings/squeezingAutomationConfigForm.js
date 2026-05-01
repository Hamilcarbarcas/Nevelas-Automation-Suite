import { MODULE } from '../module.js';

export class SqueezingAutomationConfigForm extends FormApplication {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      id: 'nas-squeezing-automation-config-form',
      title: game.i18n.localize("NAS.forms.squeezingAutomationConfig.title"),
      template: `modules/${MODULE.ID}/src/templates/squeezing-automation-config-form.html`,
      width: 460,
      height: 'auto',
      closeOnSubmit: true
    });
  }

  getData() {
    const squeezedHandling = game.settings.get(MODULE.ID, 'squeezedHandling');
    const squeezedExitHandling = game.settings.get(MODULE.ID, 'squeezedExitHandling');
    const squeezedEscapeFailureHandling = game.settings.get(MODULE.ID, 'squeezedEscapeFailureHandling');
    const squeezingMediumBodyWidth = game.settings.get(MODULE.ID, 'squeezingMediumBodyWidth');
    const squeezingMediumHeadWidth = game.settings.get(MODULE.ID, 'squeezingMediumHeadWidth');
    const squeezingEscapeArtistDC = game.settings.get(MODULE.ID, 'squeezingEscapeArtistDC');

    return {
      squeezedHandling,
      squeezedExitHandling,
      squeezedEscapeFailureHandling,
      squeezingMediumBodyWidth,
      squeezingMediumHeadWidth,
      squeezingEscapeArtistDC,
      squeezedHandlingChoices: {
        disabled: game.i18n.localize("NAS.common.choices.handling.disabledNoRestrictions"),
        strict: game.i18n.localize("NAS.settings.squeezedHandling.choices.strict"),
        lenient: game.i18n.localize("NAS.common.choices.handling.lenientWarning")
      },
      squeezedExitHandlingChoices: {
        count: game.i18n.localize("NAS.settings.squeezedExitHandling.choices.count"),
        ignore: game.i18n.localize("NAS.settings.squeezedExitHandling.choices.ignore")
      },
      squeezedEscapeFailureHandlingChoices: {
        stopBeforeNarrow: game.i18n.localize("NAS.settings.squeezedEscapeFailureHandling.choices.stopBeforeNarrow"),
        enterFirstNarrowSquare: game.i18n.localize("NAS.settings.squeezedEscapeFailureHandling.choices.enterFirstNarrowSquare")
      }
    };
  }

  async _updateObject(_event, formData) {
    await game.settings.set(MODULE.ID, 'squeezedHandling', formData.squeezedHandling);
    await game.settings.set(MODULE.ID, 'squeezedExitHandling', formData.squeezedExitHandling);
    await game.settings.set(MODULE.ID, 'squeezedEscapeFailureHandling', formData.squeezedEscapeFailureHandling);
    await game.settings.set(MODULE.ID, 'squeezingMediumBodyWidth', Number(formData.squeezingMediumBodyWidth));
    await game.settings.set(MODULE.ID, 'squeezingMediumHeadWidth', Number(formData.squeezingMediumHeadWidth));
    await game.settings.set(MODULE.ID, 'squeezingEscapeArtistDC', Number(formData.squeezingEscapeArtistDC));
  }
}
