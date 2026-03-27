import { MODULE } from '../../../common/module.js';
import { getDamageTypes } from '../../../common/settings/damageSettingsForms.js';

let _damageFootnoteHooksRegistered = false;

export function registerDamageFootnoteHooks() {
  if (_damageFootnoteHooksRegistered) return;
  _damageFootnoteHooksRegistered = true;
  Hooks.on('pf1PreActionUse', handlePreActionUse);
}

function handlePreActionUse(action) {
  const { shared, item, actor, formData, token } = action;
  if (!shared) return;

  const moduleId = MODULE.ID;
  const actionId = action.action?.id;
  const flags = item?.flags?.[moduleId] || {};

  if (!shared.templateData) shared.templateData = {};
  if (!Array.isArray(shared.templateData.footnotes)) shared.templateData.footnotes = [];
  if (Array.isArray(shared.chatAttacks)) {
    for (const chatAttack of shared.chatAttacks) {
      if (!Array.isArray(chatAttack.effectNotes)) chatAttack.effectNotes = [];
    }
  }

  const global = flags.globalItemSettings || {};
  const itemActionSettings = flags.itemActionSettings || {};
  const actions = Array.isArray(itemActionSettings.actions) ? itemActionSettings.actions : [];
  const actionSettings = actions.find(a => a.id === actionId);

  function addBypassFootnotes(source, category, label, pluralLabel, addFn) {
    const cat = source[category];
    if (cat?.bypass?.enabled && Array.isArray(cat.bypass.types) && cat.bypass.types.length > 0) {
      const typeMap = {};
      for (const dt of getDamageTypes(category)) typeMap[dt.id] = dt.label;
      if (cat.bypass.types.includes("all")) {
        addFn(`Bypass All ${pluralLabel}`);
      } else {
        for (const type of cat.bypass.types) {
          const typeLabel = typeMap[type] || type;
          if (category === "damageReduction" && type === "dr-none") {
            addFn("Bypass DR/-");
          } else {
            addFn(`Bypass ${typeLabel} ${label}`);
          }
        }
      }
    }
  }

  if (actionSettings && actionSettings.hardness && actionSettings.hardness.bypass?.inherit === false) {
    if (actionSettings.hardness.bypass.enabled) {
      addGlobalFootnote(shared, "Bypass Hardness");
    }
  } else if (global.hardness?.bypass) {
    addGlobalFootnote(shared, "Bypass Hardness");
  }

  if (actionSettings && actionSettings.immunity && actionSettings.immunity.inherit === false) {
    addBypassFootnotes(actionSettings, "immunity", "Immunity", "Immunities", t => addGlobalFootnote(shared, t));
  } else {
    addBypassFootnotes(global, "immunity", "Immunity", "Immunities", t => addGlobalFootnote(shared, t));
  }

  if (actionSettings && actionSettings.resistance && actionSettings.resistance.inherit === false) {
    addBypassFootnotes(actionSettings, "resistance", "Resistance", "Resistances", t => addGlobalFootnote(shared, t));
  } else {
    addBypassFootnotes(global, "resistance", "Resistance", "Resistances", t => addGlobalFootnote(shared, t));
  }

  if (actionSettings && actionSettings.damageReduction && actionSettings.damageReduction.inherit === false) {
    addBypassFootnotes(actionSettings, "damageReduction", "DR", "DRs", t => addGlobalFootnote(shared, t));
  } else {
    addBypassFootnotes(global, "damageReduction", "DR", "DRs", t => addGlobalFootnote(shared, t));
  }

  if (Array.isArray(shared.chatAttacks) && actionSettings && Array.isArray(actionSettings.attacks)) {
    const hasteLabel = game.i18n.localize("PF1.Haste");
    const rapidShotLabel = game.i18n.localize("PF1.RapidShot");
    for (let i = 0; i < shared.chatAttacks.length; i++) {
      const chatAttack = shared.chatAttacks[i];
      if (!Array.isArray(chatAttack.effectNotes)) chatAttack.effectNotes = [];
      const chatLabel = chatAttack.label;
      let attackSetting = null;
      if (chatLabel === hasteLabel || chatLabel === "Haste") {
        attackSetting = actionSettings.attacks.find(a => a.name === "haste");
      } else if (chatLabel === rapidShotLabel || chatLabel === "Rapid Shot") {
        attackSetting = actionSettings.attacks.find(a => a.name === "rapid_shot");
      } else {
        attackSetting = actionSettings.attacks[i];
        if (chatLabel && actionSettings.attacks.some(a => a.name === chatLabel)) {
          attackSetting = actionSettings.attacks.find(a => a.name === chatLabel);
        }
      }
      if (!attackSetting) continue;
      function addAttackBypassFootnotes(source, category, label, pluralLabel) {
        addBypassFootnotes(source, category, label, pluralLabel, t => addAttackFootnote(shared, i, t));
      }
      if (attackSetting.hardness && attackSetting.hardness.bypass?.inherit === false) {
        if (attackSetting.hardness.bypass.enabled) {
          addAttackFootnote(shared, i, "Bypass Hardness");
        }
      }
      if (attackSetting.immunity && attackSetting.immunity.inherit === false) {
        addAttackBypassFootnotes(attackSetting, "immunity", "Immunity", "Immunities");
      }
      if (attackSetting.resistance && attackSetting.resistance.inherit === false) {
        addAttackBypassFootnotes(attackSetting, "resistance", "Resistance", "Resistances");
      }
      if (attackSetting.damageReduction && attackSetting.damageReduction.inherit === false) {
        addAttackBypassFootnotes(attackSetting, "damageReduction", "DR", "DRs");
      }
    }
  }
}

export function addGlobalFootnote(shared, text) {
  if (!shared.templateData) shared.templateData = {};
  if (!Array.isArray(shared.templateData.footnotes)) shared.templateData.footnotes = [];
  shared.templateData.footnotes.push({ text });
}

export function addAttackFootnote(shared, attackIndex, text) {
  if (!Array.isArray(shared.chatAttacks)) return;
  const chatAttack = shared.chatAttacks[attackIndex];
  if (!chatAttack) return;
  if (!Array.isArray(chatAttack.effectNotes)) chatAttack.effectNotes = [];
  chatAttack.effectNotes.push({ text });
} 



