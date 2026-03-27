import { MODULE } from '../../../../common/module.js';
import { socket } from '../../../../integration/moduleSockets.js';

export function checkNextTokenFlatFooted(combat, combatData) {
  const token = canvas.tokens.get(combatData.tokenId);
  if (!token || !token.actor) return;
  
  if (token.actor.statuses.has("flatFooted")) {
    socket.executeAsGM("handleFlatFootedRemoval", combatData.tokenId, combatData.round, combatData.turn);
  }
}

export function restoreFlatFootedTracker(combat) {
  const trackerData = combat.getFlag(MODULE.ID, 'flatFootedTracker') || {};
  flatFootedTracker.clear(); 
  for (const [tokenId, data] of Object.entries(trackerData)) {
    flatFootedTracker.set(tokenId, data);
  }
}

const flatFootedTracker = new Map();

export function updateFlatFootedTracker(combat) {
  const currentId = combat.combatant?.id;
  
  if (!currentId) return;
  
  const combatant = combat.combatants.get(currentId);
  const tokenId = combatant?.token?.id;
  
  if (!tokenId) return;
  
  const ffTracker = combat.getFlag(MODULE.ID, "flatFootedTracker") || {};
  
  if (ffTracker[tokenId]) {
    ffTracker[tokenId].hasActed = true;
    combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
  }
}

export function handleFlatFootedCondition(combat, combatData) {
  if (!game.settings.get(MODULE.ID, 'autoApplyFF')) return;
  
  const currentTokenId = combatData.tokenId;
  const token = canvas.tokens.get(currentTokenId);
  if (!token) return;
  
  const actor = token.actor;
  if (!actor) return;
  
  const currentRound = combat.round;
  const currentTurn = combat.turn;
  const previousTurn = combatData.previousTurn;
  const previousRound = combatData.previousRound;
  
  const ffTracker = combat.getFlag(MODULE.ID, "flatFootedTracker") || {};
  
  const isMovingBackward = 
    (previousRound > currentRound) || 
    (previousRound === currentRound && previousTurn > currentTurn);
  
  if (isMovingBackward) {
    for (const [tokenId, trackerData] of Object.entries(ffTracker)) {
      if (!trackerData || !trackerData.removalInfo) continue;
      
      const { removedOnRound, removedOnTurn } = trackerData.removalInfo;
      
      if (removedOnRound > currentRound || 
          (removedOnRound === currentRound && removedOnTurn > currentTurn)) {
        const targetToken = canvas.tokens.get(tokenId);
        if (targetToken && targetToken.actor) {
          targetToken.actor.setCondition("flatFooted", true);
          
          trackerData.wasFlatFooted = true;
          trackerData.removalInfo = null;
          
          ffTracker[tokenId] = trackerData;
        }
      }
    }
    
    combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
  } else {
    if (actor.statuses.has("flatFooted")) {
      const trackerData = ffTracker[currentTokenId];
      
      if (trackerData) {
        const targetRemovalRound = trackerData.targetRemovalRound || 1;
        
        if (currentRound >= targetRemovalRound) {
          const isBeingProcessed = combat.getFlag(MODULE.ID, "processingFlatFooted") || {};
          
          if (isBeingProcessed[currentTokenId]) return;
          
          isBeingProcessed[currentTokenId] = true;
          combat.setFlag(MODULE.ID, "processingFlatFooted", isBeingProcessed);
          
    actor.setCondition("flatFooted", false);
          
          trackerData.removalInfo = {
            removedOnRound: currentRound,
            removedOnTurn: currentTurn
          };
          trackerData.wasFlatFooted = false;
          
          ffTracker[currentTokenId] = trackerData;
          combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
          
          isBeingProcessed[currentTokenId] = false;
          combat.setFlag(MODULE.ID, "processingFlatFooted", isBeingProcessed);
        }
      }
    }
  }
}

export async function resetExemptFlags(combat) {
  for (const combatant of combat.combatants) {
    const token = canvas.tokens.get(combatant.token.id);
    if (token) {
      await token.document.unsetFlag(MODULE.ID, 'exempt-from-confusion-roll');
    }
  }
}

export async function resetCombatFlags(combat) {
  await combat.unsetFlag(MODULE.ID, "flatFootedTracker");
  await combat.unsetFlag(MODULE.ID, "flatFootedProcessed");
  await combat.unsetFlag(MODULE.ID, "processingFlatFooted");
  await combat.unsetFlag(MODULE.ID, "previousTurnData");
  await combat.unsetFlag(MODULE.ID, "previousRoundData");
  await combat.unsetFlag(MODULE.ID, "isSurprise");
  await combat.unsetFlag(MODULE.ID, "variantBuffTracker");
  
  for (const combatant of combat.combatants) {
    const token = canvas.tokens.get(combatant.token.id);
    if (token && token.actor && token.actor.statuses.has("flatFooted")) {
      await token.actor.setCondition("flatFooted", false);
    }
  }
  
  await resetExemptFlags(combat);
}



