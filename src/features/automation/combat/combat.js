
import { MODULE } from '../../../common/module.js';
import { handleConfusionCondition, handleConfusionOnCombatStart } from '../conditions/confusion/confusion.js';
import { resolveBuffReference, activateVariantForTarget } from '../buffs/buffs.js';
import { handleFlatFootedCondition, resetCombatFlags, resetExemptFlags, updateFlatFootedTracker } from '../conditions/flatfooted/flatfooted.js';



export async function handleFlatFootedOnCombatStart(combat, combatant, token, turnIndex, highestInitiative, isSurprise) {
  if (!game.settings.get(MODULE.ID, 'autoApplyFF')) return;
  
  const actor = token.actor;
  if (!actor) return;
  
  if (actor.statuses.has("flatFooted")) return;
  
  const hasUncanny = actor.items.some(i => 
    i.type === "feat" && 
    i.name.toLowerCase().includes("uncanny dodge")
  );
  
  if (hasUncanny) return;
  
  const processedActors = combat.getFlag(MODULE.ID, "flatFootedProcessed") || [];
  if (processedActors.includes(token.id)) return;
  
  const exemptFromSurprise = actor.getFlag(MODULE.ID, 'exemptFromSurprise') || false;
  const isFlatFootedUntilTurn = (isSurprise && !exemptFromSurprise) || combatant.initiative < highestInitiative;
  
  if (isFlatFootedUntilTurn) {
    await actor.setCondition("flatFooted", true);
    
    const flatFootedInfo = {
      tokenId: token.id,
      wasFlatFooted: true,
      appliedOnRound: 1, 
      appliedOnTurn: turnIndex,
      targetRemovalRound: isSurprise && !exemptFromSurprise ? 2 : 1,
      removalInfo: null
    };
    
    const ffTracker = combat.getFlag(MODULE.ID, "flatFootedTracker") || {};
    ffTracker[token.id] = flatFootedInfo;
    await combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
    
    processedActors.push(token.id);
    await combat.setFlag(MODULE.ID, "flatFootedProcessed", processedActors);
  }
}

export function restoreFlatFootedTracker(combat) {
  const ffTracker = {};
  for (const combatant of combat.combatants) {
    ffTracker[combatant.id] = false;
  }
  combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
}

export function handleCombatTurn(combat, combatData) {
  const previousTurnData = combat.getFlag(MODULE.ID, "previousTurnData") || {
    round: combat.round,
    turn: combat.turn
  };
  
  const enhancedCombatData = {
    ...combatData,
    previousRound: previousTurnData.round,
    previousTurn: previousTurnData.turn
  };
  
  combat.setFlag(MODULE.ID, "previousTurnData", {
    round: combat.round,
    turn: combat.turn
  });
  
  handleFlatFootedCondition(combat, enhancedCombatData);
  handleVariantBuffsOnTurn(combat, enhancedCombatData);
  
  if (!(combat.round === 1 && combat.turn === 0)) {
  handleConfusionCondition(combat, combatData);
  }
}

export async function handleCombatRound(combat, round) {
  if (round === 0) {
    await resetCombatFlags(combat);
    return;
  }
  
  await combat.unsetFlag(MODULE.ID, "flatFootedProcessed");
  
  const previousRoundData = combat.getFlag(MODULE.ID, "previousRoundData") || { round: 0 };
  const currentRound = round;
  const turnOrder = combat.turns;
  const currentTurn = combat.turn;
  const isSurprise = combat.getFlag(MODULE.ID, 'isSurprise') || false;
  
  const isGoingBackwards = previousRoundData.round > currentRound;
  
  await combat.setFlag(MODULE.ID, "previousRoundData", { round: currentRound });
  
  const ffTracker = combat.getFlag(MODULE.ID, "flatFootedTracker") || {};
  
  if (isGoingBackwards) {
    for (let i = 0; i < turnOrder.length; i++) {
      const combatant = turnOrder[i];
      const tokenId = combatant.token.id;
      const token = canvas.tokens.get(tokenId);
      
      if (!token || !token.actor) continue;
      
      const trackerData = ffTracker[tokenId];
      if (!trackerData) continue;
    
      
      const shouldHaveFlatFooted = 
        trackerData.wasFlatFooted &&
        (i >= currentTurn || currentRound < trackerData.targetRemovalRound);
      
      if (shouldHaveFlatFooted && !token.actor.statuses.has("flatFooted")) {
        await token.actor.setCondition("flatFooted", true);
        
        if (trackerData.removalInfo) {
          trackerData.removalInfo = null;
          ffTracker[tokenId] = trackerData;
        }
      } else if (!shouldHaveFlatFooted && token.actor.statuses.has("flatFooted")) {
        await token.actor.setCondition("flatFooted", false);
      }
    }
    
    await combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
  } else if (currentRound === 1) {
    const sortedTurnOrder = turnOrder.sort((a, b) => b.initiative - a.initiative);
    
    const highestInitiative = sortedTurnOrder.length > 0 ? sortedTurnOrder[0].initiative : 0;
    
    for (let i = 0; i < sortedTurnOrder.length; i++) {
      const combatant = sortedTurnOrder[i];
      const token = canvas.tokens.get(combatant.token.id);
      
      if (token && token.actor) {
        await handleFlatFootedOnCombatStart(combat, combatant, token, i, highestInitiative, isSurprise);
        
      }
    }
  } else {
    
    
    for (const [tokenId, trackerData] of Object.entries(ffTracker)) {
      if (trackerData && !trackerData.currentRound) {
        trackerData.currentRound = currentRound;
        ffTracker[tokenId] = trackerData;
      }
    }
    
    await combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
    
    await resetExemptFlags(combat);
  }
  
  updateFlatFootedTracker(combat);
}

export async function handleVariantBuffsOnTurn(combat, combatData) {
  const tracker = combat.getFlag(MODULE.ID, "variantBuffTracker") || [];
  if (!tracker.length) return;

  const currentTokenId = combatData.tokenId;
  const currentRound = combat.round;
  const currentTurn = combat.turn;
  let changed = false;
  const updatedEntries = [];

  for (const entry of tracker) {
    const variants = await Promise.all((entry.variants || []).map(resolveBuffReference));
    const availableVariants = variants.filter(Boolean);
    if (!availableVariants.length) continue;

    const remainingTargets = [];
    for (const tgt of entry.targets || []) {
      if (tgt.tokenId !== currentTokenId) {
        remainingTargets.push(tgt);
        continue;
      }

      const reached = (currentRound > tgt.applyRound) || (currentRound === tgt.applyRound && currentTurn >= tgt.applyTurn);
      if (!reached) {
        remainingTargets.push(tgt);
        continue;
      }

      const token = canvas.tokens.get(tgt.tokenId);
      if (!token) {
        remainingTargets.push(tgt);
        continue;
      }

      const defaultVariant = availableVariants[tgt.variantIndex] || availableVariants[0];
      let chosenVariant = defaultVariant;

      if (tgt.switching) {
        chosenVariant = await promptVariantChoiceOnTurn(token, availableVariants, defaultVariant);
      }

      if (chosenVariant) {
        await activateVariantForTarget(token, chosenVariant, availableVariants, tgt.duration || entry.defaultDuration || { units: 'round', value: '1' }, entry.caster?.level, { silent: true });
      }

      if (tgt.switching) {
        const idx = availableVariants.findIndex(v => v && chosenVariant && v.id === chosenVariant.id && (v.pack || null) === (chosenVariant.pack || null));
        tgt.variantIndex = idx >= 0 ? idx : tgt.variantIndex;
        const timingChoice = tgt.applyTiming || (tgt.applyOnTurn ? 'turn' : 'cast');
        tgt.applyRound = currentRound + 1;
        tgt.applyTurn = tgt.turnIndex ?? tgt.applyTurn ?? currentTurn;
        tgt.applyTiming = timingChoice;
        remainingTargets.push(tgt);
      }

      changed = true;
    }

    if (remainingTargets.length > 0) {
      updatedEntries.push({ ...entry, targets: remainingTargets });
    }
  }

  if (changed) {
    await combat.setFlag(MODULE.ID, "variantBuffTracker", updatedEntries);
  }
}

async function promptVariantChoiceOnTurn(token, variants, defaultVariant) {
  if (!variants || variants.length === 0) return defaultVariant;
  const defaultIdx = Math.max(variants.findIndex(v => v && defaultVariant && v.id === defaultVariant.id && (v.pack || null) === (defaultVariant.pack || null)), 0);
  const options = variants.map((v, idx) => `<option value="${idx}" ${idx === defaultIdx ? 'selected' : ''}>${v.name}</option>`).join('');
  return new Promise(resolve => {
    const dlg = new Dialog({
      title: `${MODULE.ID} | ${token.name}`,
      content: `<p>${game.i18n.localize('NAS.buffs.SelectBuffVariant') || 'Select variant'}</p>
        <div class="form-group"><select id="ic-variant-choice" style="width:100%;">${options}</select></div>`,
      buttons: {
        apply: {
          icon: '<i class="fas fa-check"></i>',
          label: game.i18n.localize('NAS.buffs.ApplyBuff'),
          callback: html => {
            let idx = defaultIdx;
            if (typeof html.find === 'function') idx = Number(html.find('#ic-variant-choice').val());
            else {
              const sel = html.querySelector('#ic-variant-choice');
              idx = sel ? Number(sel.value) : defaultIdx;
            }
            resolve(variants[idx]);
          }
        },
        cancel: {
          icon: '<i class="fas fa-times"></i>',
          label: game.i18n.localize('NAS.common.buttons.cancel'),
          callback: () => resolve(defaultVariant)
        }
      },
      default: 'apply',
      close: () => resolve(defaultVariant)
    });
    dlg.render(true);
  });
}




