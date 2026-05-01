import { hasWvHealthUpdate, isWoundsVigorActive, isWvNoWoundsActor } from './woundsVigor.js';

export function hasHpUpdate(updateData) {
  return updateData.system && 
         updateData.system.attributes && 
         updateData.system.attributes.hp && 
         (updateData.system.attributes.hp.value !== undefined || updateData.system.attributes.hp.offset !== undefined);
}

export function hasHpNonlethalUpdate(updateData) {
  return updateData?.system?.attributes?.hp?.nonlethal !== undefined;
}

export function hasWoundsUpdate(updateData) {
  return updateData?.system?.attributes?.wounds &&
    (updateData.system.attributes.wounds.value !== undefined ||
      updateData.system.attributes.wounds.offset !== undefined ||
      updateData.system.attributes.wounds.threshold !== undefined);
}

export function getNewHp(actor, updateData) {
  if (updateData.system.attributes.hp.value !== undefined) {
    return updateData.system.attributes.hp.value;
  }
  else if (updateData.system.attributes.hp.offset !== undefined) {
    return actor.system.attributes.hp.value;
  }
  return actor.system.attributes.hp.value;
}

export function getNewWounds(actor, updateData) {
  const woundsUpdate = updateData?.system?.attributes?.wounds;
  if (woundsUpdate?.value !== undefined) {
    return woundsUpdate.value;
  }
  return actor?.system?.attributes?.wounds?.value ?? 0;
}

export function getNewVigor(actor, updateData) {
  const vigorUpdate = updateData?.system?.attributes?.vigor;
  if (vigorUpdate?.value !== undefined) {
    return vigorUpdate.value;
  }
  return actor?.system?.attributes?.vigor?.value ?? 0;
}

export function getWoundsThreshold(actor, updateData) {
  const woundsUpdate = updateData?.system?.attributes?.wounds;
  if (woundsUpdate?.threshold !== undefined) {
    return woundsUpdate.threshold;
  }
  return actor?.system?.attributes?.wounds?.threshold ?? 0;
}

export function hasPrimaryHealthUpdate(actor, updateData) {
  if (isWoundsVigorActive(actor)) return hasWvHealthUpdate(updateData);
  return hasHpUpdate(updateData);
}

export function getPrimaryHealthValue(actor, updateData) {
  if (isWoundsVigorActive(actor)) {
    if (isWvNoWoundsActor(actor)) return getNewVigor(actor, updateData);
    return getNewWounds(actor, updateData);
  }
  return getNewHp(actor, updateData);
}

export function hasNegativeLevelUpdate(updateData) {
  return updateData.system && 
         updateData.system.attributes && 
         updateData.system.attributes.energyDrain !== undefined;
}

export function getNewNegativeLevels(actor, updateData) {
  return updateData.system.attributes.energyDrain;
}
