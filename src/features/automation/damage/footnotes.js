import { MODULE } from '../../../common/module.js';
import { getDamageTypes } from '../../../common/settings/damageSettingsForms.js';
import {
  captureMirrorImageEffectNoteRoll,
  prepareMirrorImageEffectNote
} from '../buffs/mirrorImage.js';

let _damageFootnoteHooksRegistered = false;
let _attackFootnoteWrapperRegistered = false;
let _attackFootnoteQueueWrapperRegistered = false;

const ATTACK_FOOTNOTE_QUEUE_KEY = "__nasPendingAttackFootnotes";

export function registerDamageFootnoteHooks() {
  if (_damageFootnoteHooksRegistered) return;
  _damageFootnoteHooksRegistered = true;
  Hooks.on('pf1PreActionUse', handlePreActionUse);
  registerAttackFootnoteQueueWrapper();
  registerAttackFootnoteRenderWrapper();
}

function registerAttackFootnoteQueueWrapper() {
  if (_attackFootnoteQueueWrapperRegistered) return;
  _attackFootnoteQueueWrapperRegistered = true;

  if (!game.modules.get("lib-wrapper")?.active) {
    console.warn(`${MODULE.ID} | libWrapper missing; attack-level footnote queue fallback disabled.`);
    return;
  }

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ActionUse.prototype.addEffectNotes",
    async function (wrapped, ...args) {
      try {
        this.shared[ATTACK_FOOTNOTE_QUEUE_KEY] = Object.create(null);
        queueAttackFootnotesFromSettings(this);
      } catch (err) {
        console.error(`${MODULE.ID} | Failed preparing attack footnote queue`, err);
      }
      return wrapped(...args);
    },
    "WRAPPER"
  );
}

function registerAttackFootnoteRenderWrapper() {
  if (_attackFootnoteWrapperRegistered) return;
  _attackFootnoteWrapperRegistered = true;

  if (!game.modules.get("lib-wrapper")?.active) {
    console.warn(`${MODULE.ID} | libWrapper missing; attack-level footnote rendering fallback disabled.`);
    return;
  }

  libWrapper.register(
    MODULE.ID,
    "pf1.actionUse.ChatAttack.prototype.setEffectNotesHTML",
    async function (wrapped, ...args) {
      try {
        prepareMirrorImageEffectNote(this);
        const actionUse = this.actionUse;
        const shared = actionUse?.shared;
        const pendingByIndex = shared?.[ATTACK_FOOTNOTE_QUEUE_KEY];
        if (pendingByIndex && Array.isArray(shared?.chatAttacks)) {
          const attackIndex = shared.chatAttacks.indexOf(this);
          if (attackIndex >= 0) {
            const pendingNotes = pendingByIndex[attackIndex] ?? [];
            if (pendingNotes.length > 0) {
              if (!Array.isArray(this.effectNotes)) this.effectNotes = [];
              const existingText = new Set(this.effectNotes.map((entry) => entry?.text).filter(Boolean));
              for (const text of pendingNotes) {
                if (typeof text !== "string") continue;
                const trimmed = text.trim();
                if (!trimmed || existingText.has(trimmed)) continue;
                this.effectNotes.push({ text: trimmed });
                existingText.add(trimmed);
              }
            }
          }
        }
      } catch (err) {
        console.error(`${MODULE.ID} | Failed injecting queued attack footnotes`, err);
      }

      const result = await wrapped(...args);

      try {
        captureMirrorImageEffectNoteRoll(this);
      } catch (err) {
        console.error(`${MODULE.ID} | Failed capturing Mirror Image effect note roll`, err);
      }

      return result;
    },
    "WRAPPER"
  );
}

function getPendingAttackFootnoteQueue(shared) {
  if (!shared) return null;
  shared[ATTACK_FOOTNOTE_QUEUE_KEY] ??= Object.create(null);
  return shared[ATTACK_FOOTNOTE_QUEUE_KEY];
}

function queueAttackFootnote(shared, attackIndex, text) {
  if (!shared || !Number.isInteger(attackIndex)) return;
  if (typeof text !== "string" || !text.trim()) return;

  const queue = getPendingAttackFootnoteQueue(shared);
  if (!queue) return;
  queue[attackIndex] ??= [];
  if (!queue[attackIndex].includes(text)) {
    queue[attackIndex].push(text);
  }
}

function queueAttackFootnotesFromSettings(actionUse) {
  const shared = actionUse?.shared;
  const item = actionUse?.item;
  if (!shared || !item || !Array.isArray(shared.chatAttacks)) return;

  const actionId = actionUse?.action?.id;
  const flags = item?.flags?.[MODULE.ID] || {};
  const itemActionSettings = flags.itemActionSettings || {};
  const actions = Array.isArray(itemActionSettings.actions) ? itemActionSettings.actions : [];
  const actionSettings = actions.find((a) => a.id === actionId);
  if (!actionSettings || !Array.isArray(actionSettings.attacks)) return;

  const hasteLabel = game.i18n.localize("PF1.Haste");
  const rapidShotLabel = game.i18n.localize("PF1.RapidShot");

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

  for (let i = 0; i < shared.chatAttacks.length; i++) {
    const chatAttack = shared.chatAttacks[i];
    if (!Array.isArray(chatAttack.effectNotes)) chatAttack.effectNotes = [];
    const chatLabel = chatAttack?.label;
    let attackSetting = null;
    if (chatLabel === hasteLabel || chatLabel === "Haste") {
      attackSetting = actionSettings.attacks.find((a) => a.name === "haste");
    } else if (chatLabel === rapidShotLabel || chatLabel === "Rapid Shot") {
      attackSetting = actionSettings.attacks.find((a) => a.name === "rapid_shot");
    } else {
      attackSetting = actionSettings.attacks[i];
      if (chatLabel && actionSettings.attacks.some((a) => a.name === chatLabel)) {
        attackSetting = actionSettings.attacks.find((a) => a.name === chatLabel);
      }
    }
    if (!attackSetting) continue;

    function addAttackBypassFootnotes(source, category, label, pluralLabel) {
      addBypassFootnotes(source, category, label, pluralLabel, (t) => addAttackFootnote(shared, i, t));
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

function handlePreActionUse(action) {
  const { shared, item } = action;
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

}

export function addGlobalFootnote(shared, text) {
  if (!shared.templateData) shared.templateData = {};
  if (!Array.isArray(shared.templateData.footnotes)) shared.templateData.footnotes = [];
  shared.templateData.footnotes.push({ text });
}

export function addAttackFootnote(shared, attackIndex, text) {
  queueAttackFootnote(shared, attackIndex, text);

  if (!Array.isArray(shared.chatAttacks)) return;
  const chatAttack = shared.chatAttacks[attackIndex];
  if (!chatAttack) return;
  if (!Array.isArray(chatAttack.effectNotes)) chatAttack.effectNotes = [];
  if (!chatAttack.effectNotes.some((entry) => entry?.text === text)) {
    chatAttack.effectNotes.push({ text });
  }
} 



