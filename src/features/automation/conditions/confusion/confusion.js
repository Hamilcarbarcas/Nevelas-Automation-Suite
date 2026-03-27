
import { MODULE } from '../../../../common/module.js';

function normalizeDamageTypeId(typeRef) {
  const needle = String(typeRef).trim().toLowerCase();
  if (!needle) return null;

  const reg = pf1?.registry?.damageTypes;
  if (reg?.get) {
    const direct = reg.get(typeRef);
    if (direct) return typeRef;
    for (const [key, value] of reg.entries()) {
      const name = value?.name?.toLowerCase();
      const shortName = value?.shortName?.toLowerCase();
      if (needle === String(key).toLowerCase() || needle === name || needle === shortName) {
        return key;
      }
    }
  }

  const alignments = pf1?.config?.damageResistances || {};
  for (const [key, label] of Object.entries(alignments)) {
    if (needle === String(key).toLowerCase()) return key;
    if (needle === String(label).toLowerCase()) return key;
  }

  return null;
}

function resolveDamageTypeIds(damageTypes) {
  const raw = Array.isArray(damageTypes) ? damageTypes : [damageTypes];
  const resolved = [];
  for (const typeRef of raw) {
    if (!typeRef || typeof typeRef !== "string") continue;
    const id = normalizeDamageTypeId(typeRef);
    if (id) resolved.push(id);
  }
  return resolved.length ? Array.from(new Set(resolved)) : ["untyped"];
}

function getDamageTypeLabels(typeIds) {
  const labels = [];
  for (const typeId of typeIds) {
    const entry = pf1?.registry?.damageTypes?.get?.(typeId);
    if (entry?.name) labels.push(entry.name);
    else labels.push(typeId);
  }
  return labels;
}

function getStrengthLabel() {
  return pf1?.config?.abilities?.str || "Strength";
}

function applyDamageTypeFlavors(terms, damageTypeFlavor, strengthLabel) {
  for (const term of terms) {
    const termClass = term?.constructor?.name;
    if (termClass === "OperatorTerm") continue;
    term.options ||= {};
    if (termClass === "NumericTerm") {
      term.options.flavor = strengthLabel;
    } else {
      term.options.flavor = damageTypeFlavor;
    }
  }
}

export function getBehaviorData(actor, rollResult) {
  if (rollResult <= 25) {
    return { id: 1, description: game.i18n.localize("NAS.conditions.confused.effects.1") };
  } else if (rollResult <= 50) {
    return { id: 2, description: game.i18n.localize("NAS.conditions.confused.effects.2") };
  } else if (rollResult <= 75) {
    const meleeItems = actor.items.filter(i => 
      i.type === "weapon" &&
      i.system.actions.some(a => a.actionType === "mwak") &&
      !i.system.broken
    );
    
    const strMod = actor.system.abilities.str.mod;
    if (meleeItems.length > 0) {
      const selectedItem = meleeItems[Math.floor(Math.random() * meleeItems.length)];
      
      const baseDamage = selectedItem.system.actions[0].damage.parts.length ? selectedItem.system.actions[0].damage.parts[0].formula : "1d8";
      const damageTypes = selectedItem.system.actions[0].damage.parts.length && selectedItem.system.actions[0].damage.parts[0].types.length ? 
                          selectedItem.system.actions[0].damage.parts[0].types : ["bludgeoning"];
      
      return { 
        id: 3, 
        description: game.i18n.localize("NAS.conditions.confused.effects.3a"),
        itemName: selectedItem.name,
        damageFormula: `1d8 + ${strMod}`,
        damageTypes: damageTypes,
        itemUsed: selectedItem
      };
    } else {
      return { 
        id: 3, 
        description: game.i18n.localize("NAS.conditions.confused.effects.3b"),
        damageFormula: `1d8 + ${strMod}`,
        damageTypes: damageTypes
      };
    }
  } else {
    return { id: 4, description: game.i18n.localize("NAS.conditions.confused.effects.4") };
  }
}

export async function handleConfusionCondition(combat, combatData) {
  if (!game.settings.get(MODULE.ID, 'handleConfused')) return;

  const token = canvas.tokens.get(combatData.tokenId);
  if (!token) return;

  const actor = token.actor;
  if (!actor) return;

  const isConfused = actor.statuses.has("confused");
  if (!isConfused) return;

  const roll = new Roll("1d100");
  await roll.evaluate();
  const result = roll.total;

  const behavior = getBehaviorData(actor, result);

  const isHiddenOrInvisible = token.document.hidden || actor.statuses.has("invisible");
  const whisperTargets = getWhisperTargets(actor, isHiddenOrInvisible);

  if (behavior.id === 3) {
    const resolvedTypes = resolveDamageTypeIds(behavior.damageTypes);
    const damageRoll = new pf1.dice.DamageRoll(behavior.damageFormula, {}, {
      damageType: resolvedTypes,
      type: "normal"
    });
    
    await damageRoll.evaluate();
    const damageTypeLabels = getDamageTypeLabels(resolvedTypes);
    const damageTypeFlavor = damageTypeLabels.join(", ");
    const strengthLabel = getStrengthLabel();
    applyDamageTypeFlavors(damageRoll.terms, damageTypeFlavor, strengthLabel);
    
    createConfusionEffectMessage({
      token,
      behavior,
      damageRoll,
      damageTypes: resolvedTypes,
      itemUsed: behavior.itemUsed,
      whisper: whisperTargets,
      isPrivate: isHiddenOrInvisible
    });
  } else {
    createConfusionEffectMessage({
      token,
      behavior,
      whisper: whisperTargets,
      isPrivate: isHiddenOrInvisible
    });
  }
}

function getWhisperTargets(actor, isHiddenOrInvisible) {
  if (!isHiddenOrInvisible) return [];
  
  const whisperTargets = [];
  
  const gmUsers = game.users.filter(u => u.isGM);
  whisperTargets.push(...gmUsers.map(u => u.id));
  
  if (actor.hasPlayerOwner) {
    const ownerUsers = game.users.filter(u => actor.testUserPermission(u, "OWNER") && !u.isGM);
    whisperTargets.push(...ownerUsers.map(u => u.id));
  }
  
  return [...new Set(whisperTargets)]; 
}

export function createConfusionEffectMessage({
  token,
  behavior,
  damageRoll = null,
  damageTypes = null,
  itemUsed = null,
  speakerAlias = "Confusion Effect",
  whisper = [],
  isPrivate = false,
}) {
  if (!token) return;

  let resolvedTypes = null;
  let damageTypeLabels = null;
  let damageTypeFlavor = null;
  let damageInstances = null;

  let tokenContent = `
    <div class="NAS-token" data-uuid="${token.document.uuid}" style="margin-bottom: 8px;">
      <div style="display: flex; justify-content: center;">
        <img src="${token.document.texture.src}" title="${token.name}" width="72" height="72" style="margin-bottom: 8px; cursor: pointer;"/>
      </div>
      <span style="text-align: center; display: block;">${token.name} ${behavior.description}</span>
    </div>
  `;

  if (damageRoll) {
    resolvedTypes = resolveDamageTypeIds(damageTypes);
    damageTypeLabels = getDamageTypeLabels(resolvedTypes);
    damageTypeFlavor = damageTypeLabels.join(", ");
    const strengthLabel = getStrengthLabel();
    applyDamageTypeFlavors(damageRoll.terms, damageTypeFlavor, strengthLabel);

    damageInstances = [{
      types: resolvedTypes,
      value: damageRoll.total,
      formula: String(damageRoll.total)
    }];

    const description = behavior.description
      .replace("{itemName}", behavior.itemName || "fists")
      .replace("{damage}", `<span class="confusion-damage">${damageRoll.total}</span>`);

    tokenContent = `
      <div class="NAS-token" data-uuid="${token.document.uuid}" style="margin-bottom: 8px;">
        <div style="display: flex; justify-content: center;">
          <img src="${token.document.texture.src}" title="${token.name}" width="72" height="72" style="margin-bottom: 8px; cursor: pointer;"/>
        </div>
        <span style="text-align: center; display: block;">${token.name} ${description}</span>
      </div>
    `;
  }

  const messageData = {
    user: game.user.id,
    speaker: {
      scene: canvas.scene.id,
      token: token.id,
      alias: speakerAlias
    },
    type: CONST.CHAT_MESSAGE_STYLES.OTHER,
    content: tokenContent,
    whisper: whisper.length > 0 ? whisper : null
  };

  if (damageRoll) {
    const fullDamage = damageRoll.total;
    const halfDamage = Math.floor(damageRoll.total / 2);
    
    messageData.rolls = [damageRoll];
    
    messageData.flavor = `Damage: ${damageTypeLabels.join(", ")}`;
    messageData.system = {
      subject: { health: "damage" }
    };
    
    messageData.flags = {
      "better-damage": {
        damageRoll: true,
        damageTotal: damageRoll.total,
        damageTypes: resolvedTypes
      },
      [MODULE.ID]: {
        damageRoll: true,
        damageTotal: damageRoll.total,
        damageTypes: resolvedTypes,
        damageInstances,
        source: "command",
        sourceDetail: "confusion",
        subject: { health: "damage" }
      }
    };
    
    damageRoll.render().then(rollContent => {
      const cleanedRollContent = rollContent.replace(
        /<span class="part-prefix">\+<\/span>/g,
        ""
      );
      const buttonsHtml = `
        <div class="card-buttons flexrow">
          <button type="button" data-action="applyDamage" data-ratio="1" data-value="${fullDamage}" data-damage-types="${resolvedTypes.join(',')}">Apply</button>
          <button type="button" data-action="applyDamage" data-ratio="0.5" data-value="${halfDamage}" data-damage-types="${resolvedTypes.join(',')}">Apply Half</button>
        </div>
      `;
      
      messageData.content += cleanedRollContent + buttonsHtml;
      
      ChatMessage.create(messageData);
    });
  } else {
  ChatMessage.create(messageData);
  }
}

export async function handleConfusionOnCombatStart(combatant, token, turnOrder) {
  if (!game.settings.get(MODULE.ID, 'handleConfused')) return;
  
  const firstTurn = turnOrder[0];
  if (!firstTurn || combatant.id !== firstTurn.id) return;
  
  const actor = token.actor;
  if (!actor) return;
  
  const isConfused = actor.statuses.has("confused");
  if (!isConfused) return;
  
  const roll = new Roll("1d100");
  await roll.evaluate();
  const result = roll.total;
  
  const behavior = getBehaviorData(actor, result);
  
  const isHiddenOrInvisible = token.document.hidden || actor.statuses.has("invisible");
  const whisperTargets = getWhisperTargets(actor, isHiddenOrInvisible);
  
  if (behavior.id === 3) {
    const resolvedTypes = resolveDamageTypeIds(behavior.damageTypes);
    const damageRoll = new pf1.dice.DamageRoll(behavior.damageFormula, {}, {
      damageType: resolvedTypes,
      type: "normal"
    });
    
    await damageRoll.evaluate();
    const damageTypeLabels = getDamageTypeLabels(resolvedTypes);
    const damageTypeFlavor = damageTypeLabels.join(", ");
    const strengthLabel = getStrengthLabel();
    applyDamageTypeFlavors(damageRoll.terms, damageTypeFlavor, strengthLabel);
    
    createConfusionEffectMessage({
      token,
      behavior,
      damageRoll,
      damageTypes: resolvedTypes,
      itemUsed: behavior.itemUsed,
      whisper: whisperTargets,
      isPrivate: isHiddenOrInvisible
    });
  } else {
    createConfusionEffectMessage({
      token,
      behavior,
      whisper: whisperTargets,
      isPrivate: isHiddenOrInvisible
    });
  }
}

export async function handleConfusionForFirstToken(token) {
  if (!game.settings.get(MODULE.ID, 'handleConfused')) return;
  
  const actor = token.actor;
  if (!actor) return;
  
  const isConfused = actor.statuses.has("confused");
  if (!isConfused) return;
  
  const roll = new Roll("1d100");
  await roll.evaluate();
  const result = roll.total;
  
  const behavior = getBehaviorData(actor, result);
  
  const isHiddenOrInvisible = token.document.hidden || actor.statuses.has("invisible");
  const whisperTargets = getWhisperTargets(actor, isHiddenOrInvisible);
  
  if (behavior.id === 3) {
    const resolvedTypes = resolveDamageTypeIds(behavior.damageTypes);
    const damageRoll = new pf1.dice.DamageRoll(behavior.damageFormula, {}, {
      damageType: resolvedTypes,
      type: "normal"
    });
    
    await damageRoll.evaluate();
    const damageTypeLabels = getDamageTypeLabels(resolvedTypes);
    const damageTypeFlavor = damageTypeLabels.join(", ");
    const strengthLabel = getStrengthLabel();
    applyDamageTypeFlavors(damageRoll.terms, damageTypeFlavor, strengthLabel);
    
    createConfusionEffectMessage({
      token,
      behavior,
      damageRoll,
      damageTypes: resolvedTypes,
      itemUsed: behavior.itemUsed,
      whisper: whisperTargets,
      isPrivate: isHiddenOrInvisible
    });
  } else {
    createConfusionEffectMessage({
      token,
      behavior,
      whisper: whisperTargets,
      isPrivate: isHiddenOrInvisible
    });
  }
}

export function handlePrivateMessage(actor, token, tokenContent, privateMessages) {
  const activeOwner = actor.activeOwner?.id;
  const gmId = game.users.find(user => user.isGM).id;
  const whisperIds = new Set([activeOwner, gmId]);
  const whisperKey = Array.from(whisperIds).sort().join(',');

  if (!privateMessages[whisperKey]) {
    privateMessages[whisperKey] = {
      content: `<div class="confusion-message-content">`,
      whisper: Array.from(whisperIds),
      tokenCount: 0
    };
  }
  privateMessages[whisperKey].content += tokenContent;
  privateMessages[whisperKey].tokenCount += 1;
}

export function sendPrivateMessages(privateMessages) {
  Object.values(privateMessages).forEach(message => {
    if (message.tokenCount > 1) {
      message.content = message.content.replace(/<\/div><div class="NAS-token"/g, '</div><div style="border-top: 2px solid black; margin: 8px 0;"></div><div class="NAS-token"');
    }
    message.content += `</div>`;
    createConfusionEffectMessage({
      content: message.content,
      token: null,
      damageRoll: null,
      behavior: null,
      damageTypes: null,
      itemUsed: null,
      speakerAlias: "Confusion Effect",
      whisper: message.whisper,
      isPrivate: true
    });
  });
} 



