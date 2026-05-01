import { MODULE } from '../module.js';

export class BuffCompendiaSelector extends FormApplication {
  static get defaultOptions() {
    return foundry.utils.mergeObject(super.defaultOptions, {
      id: "buff-compendia-selector",
      title: game.i18n.localize("NAS.forms.buffCompendiaSelector.title"),
      template: `modules/${MODULE.ID}/src/templates/buff-compendia-selector.html`,
      classes: ["sheet"],
      width: 500,
      height: "auto",
      closeOnSubmit: true
    });
  }
  
  async getData() {
    const selectedCompendia = game.settings.get(MODULE.ID, 'customBuffCompendia');

    const includeWorldBuffs = selectedCompendia.includes("__world__");

    const systemBuffsPack = game.packs.get("pf1.buffs");
    const pfContentBuffsPack = game.packs.get("pf-content.pf-buffs");
    const specialCompendia = [];
    if (systemBuffsPack && systemBuffsPack.ownership !== ("LIMITED" || CONST.DOCUMENT_OWNERSHIP_LEVELS.LIMITED)) {
      let displayName = systemBuffsPack.title;
      if (displayName && displayName.includes('.')) displayName = game.i18n.localize(displayName);
      specialCompendia.push({
        id: systemBuffsPack.collection,
        name: displayName,
        isSelected: selectedCompendia.includes(systemBuffsPack.collection)
      });
    }
    if (pfContentBuffsPack && pfContentBuffsPack.ownership !== ("LIMITED" || CONST.DOCUMENT_OWNERSHIP_LEVELS.LIMITED)) {
      let displayName = pfContentBuffsPack.title;
      if (displayName && displayName.includes('.')) displayName = game.i18n.localize(displayName);
      specialCompendia.push({
        id: pfContentBuffsPack.collection,
        name: displayName,
        isSelected: selectedCompendia.includes(pfContentBuffsPack.collection)
      });
    }

    const itemCompendia = game.packs.filter(pack =>
      (pack.metadata.type === "Item" || pack.documentName === "Item") &&
      pack.collection !== "pf1.buffs" &&
      pack.collection !== "pf-content.pf-buffs" &&
      pack.ownership?.PLAYER !== ("LIMITED" || CONST.DOCUMENT_OWNERSHIP_LEVELS.LIMITED)
    );

    const compendiaWithBuffs = [];
    for (const pack of itemCompendia) {
      try {
        const index = await pack.getIndex();
        if (index.some(entry => entry.type === "buff")) {
          let displayName = pack.title;
          if (displayName && displayName.includes('.')) {
            displayName = game.i18n.localize(displayName);
          }
          compendiaWithBuffs.push({
            id: pack.collection,
            name: displayName,
            isSelected: selectedCompendia.includes(pack.collection)
          });
        }
      } catch (e) {
        console.warn(`${MODULE.ID} | Could not index compendium ${pack.collection}:`, e);
      }
    }

    compendiaWithBuffs.sort((a, b) => a.name.localeCompare(b.name));

    const allCompendia = [...specialCompendia, ...compendiaWithBuffs];

    return {
      compendia: allCompendia,
      includeWorldBuffs,
      worldIdLabel: game.i18n.localize("NAS.forms.buffCompendiaSelector.worldId")
    };
  }
  
  async _updateObject(event, formData) {
    const selectedCompendia = [];
    for (const [key, value] of Object.entries(formData)) {
      if (key.startsWith('compendium-') && value) {
        const compendiumId = key.substring(11); 
        selectedCompendia.push(compendiumId);
      }
    }
    if (formData.includeWorldBuffs) {
      selectedCompendia.push("__world__");
    }
    await game.settings.set(MODULE.ID, 'customBuffCompendia', selectedCompendia);
    ui.notifications.info(game.i18n.format("NAS.forms.buffCompendiaSelector.savedInfo", {
      count: selectedCompendia.length
    }));
  }
}
