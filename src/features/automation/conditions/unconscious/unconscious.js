import { MODULE } from '../../../../common/module.js';
import { getPrimaryHealthValue, hasHpNonlethalUpdate, hasPrimaryHealthUpdate } from '../../utils/healthUpdates.js';
import { detectHardToKillItems, getHtkFlag } from '../../utils/hardToKill.js';
import { socket, getPreferredOwnerUserId } from '../../../../integration/moduleSockets.js';
import { getWvState, isWoundsVigorActive } from '../../utils/woundsVigor.js';

const NONLETHAL_UNCONSCIOUS_FLAG = "nonlethalThresholdUnconscious";
const NONLETHAL_STAGGERED_FLAG = "nonlethalThresholdStaggered";

function getPrevHpFromOptions(actorDocument, options) {
  const prev = options?._nasPrevHp;
  if (Number.isFinite(prev)) return prev;
  return actorDocument?.system?.attributes?.hp?.value ?? null;
}

function getCurrentHpState(actorDocument) {
  const hp = Number(actorDocument?.system?.attributes?.hp?.value ?? 0) || 0;
  const nonlethal = Number(actorDocument?.system?.attributes?.hp?.nonlethal ?? 0) || 0;
  return { hp, nonlethal };
}

function getNonlethalThresholdState(actorDocument) {
  const { hp, nonlethal } = getCurrentHpState(actorDocument);
  return {
    hp,
    nonlethal,
    shouldUnconscious: hp >= 0 && nonlethal > hp,
    shouldStaggered: hp > 0 && nonlethal === hp
  };
}

async function syncNonlethalThresholdStates(actorDocument, change) {
  const hasRelevantUpdate =
    hasPrimaryHealthUpdate(actorDocument, change) ||
    hasHpNonlethalUpdate(change);
  if (!hasRelevantUpdate) return;

  const isDead = actorDocument.statuses?.has?.("dead");
  const { hp, shouldUnconscious, shouldStaggered } = getNonlethalThresholdState(actorDocument);
  const hadNlUnconscious = Boolean(actorDocument.getFlag(MODULE.ID, NONLETHAL_UNCONSCIOUS_FLAG));
  const hadNlStaggered = Boolean(actorDocument.getFlag(MODULE.ID, NONLETHAL_STAGGERED_FLAG));

  if (isDead || hp < 0) {
    if (hadNlUnconscious) await actorDocument.unsetFlag(MODULE.ID, NONLETHAL_UNCONSCIOUS_FLAG);
    if (hadNlStaggered) await actorDocument.unsetFlag(MODULE.ID, NONLETHAL_STAGGERED_FLAG);
    return;
  }

  if (shouldUnconscious) {
    if (!actorDocument.statuses?.has?.("unconscious")) {
      await actorDocument.setCondition("unconscious", true);
    }
    if (!actorDocument.statuses?.has?.("prone")) {
      await actorDocument.setCondition("prone", true);
    }
    if (actorDocument.statuses?.has?.("dying")) {
      await actorDocument.setCondition("dying", false);
    }
    if (actorDocument.statuses?.has?.("staggered") && hadNlStaggered) {
      await actorDocument.setCondition("staggered", false);
    }
    if (!hadNlUnconscious) await actorDocument.setFlag(MODULE.ID, NONLETHAL_UNCONSCIOUS_FLAG, true);
    if (hadNlStaggered) await actorDocument.unsetFlag(MODULE.ID, NONLETHAL_STAGGERED_FLAG);
    return;
  }

  if (hadNlUnconscious) {
    if (actorDocument.statuses?.has?.("unconscious")) {
      await actorDocument.setCondition("unconscious", false);
    }
    await actorDocument.unsetFlag(MODULE.ID, NONLETHAL_UNCONSCIOUS_FLAG);
  }

  if (shouldStaggered) {
    if (!actorDocument.statuses?.has?.("staggered")) {
      await actorDocument.setCondition("staggered", true);
    }
    if (!hadNlStaggered) await actorDocument.setFlag(MODULE.ID, NONLETHAL_STAGGERED_FLAG, true);
    return;
  }

  if (hadNlStaggered) {
    if (actorDocument.statuses?.has?.("staggered")) {
      await actorDocument.setCondition("staggered", false);
    }
    await actorDocument.unsetFlag(MODULE.ID, NONLETHAL_STAGGERED_FLAG);
  }
}

async function promptUseFeatureOnOwner(actorDocument, titleKey, promptKey) {
  if (!socket) return false;
  const ownerId = getPreferredOwnerUserId(actorDocument);
  if (!ownerId) return false;
  const out = await socket.executeAsUser("promptHTKUseFeature", ownerId, actorDocument.id, titleKey, promptKey, {});
  return out;
}

async function promptDiehardChoiceOnOwner(actorDocument) {
  if (!socket) return null;
  const ownerId = getPreferredOwnerUserId(actorDocument);
  if (!ownerId) return null;
  const out = await socket.executeAsUser("promptHTKChoice", ownerId, actorDocument.id);
  return out;
}

function getTokenUuidFromContext(actorDocument, options) {
  const tokenId = options?.tokenId ?? null;
  const sceneId = options?.sceneId ?? canvas?.scene?.id ?? null;
  if (tokenId && sceneId) return `Scene.${sceneId}.Token.${tokenId}`;

  const direct = actorDocument?.token?.uuid ?? actorDocument?.token?.document?.uuid ?? null;
  if (direct) return direct;

  const activeToken = actorDocument?.getActiveTokens?.(true, true)?.[0] ?? null;
  return activeToken?.document?.uuid ?? activeToken?.token?.document?.uuid ?? null;
}

async function useItemOnOwner(actorDocument, item, options) {
  if (!socket) return false;
  if (!actorDocument?.id || !item) return false;
  const ownerId = getPreferredOwnerUserId(actorDocument);
  if (!ownerId) return false;

  const itemId = item.id ?? item._id ?? null;
  const tokenUuid = getTokenUuidFromContext(actorDocument, options);
  const itemUuid =
    item.uuid ??
    (tokenUuid && itemId ? `${tokenUuid}.Actor.Item.${itemId}` : null) ??
    (itemId ? `Actor.${actorDocument.id}.Item.${itemId}` : null);
  if (!itemUuid || typeof itemUuid !== "string") return false;

  return socket.executeAsUser("useHtkItem", ownerId, actorDocument.id, itemUuid, tokenUuid);
}

async function showItemCardOnOwner(actorDocument, item, options) {
  if (!socket) return false;
  if (!actorDocument?.id || !item) return false;
  const ownerId = getPreferredOwnerUserId(actorDocument);
  if (!ownerId) return false;

  const itemId = item.id ?? item._id ?? null;
  const tokenUuid = getTokenUuidFromContext(actorDocument, options);
  const itemUuid =
    item.uuid ??
    (tokenUuid && itemId ? `${tokenUuid}.Actor.Item.${itemId}` : null) ??
    (itemId ? `Actor.${actorDocument.id}.Item.${itemId}` : null);
  if (!itemUuid || typeof itemUuid !== "string") return false;

  return socket.executeAsUser("postHtkItemCard", ownerId, actorDocument.id, itemUuid, tokenUuid);
}

async function deductOneDailyUse(item) {
  if (!item) return false;
  const uses = item.system?.uses;
  if (!uses?.per) return false;
  const value = Number(uses.value ?? 0);
  if (!Number.isFinite(value) || value <= 0) return false;
  await item.update({ "system.uses.value": value - 1 });
  return true;
}

async function handleWoundsVigorUnconscious(actorDocument, change) {
  if (!hasPrimaryHealthUpdate(actorDocument, change)) return;

  const woundsValue = getPrimaryHealthValue(actorDocument, change);
  const state = getWvState(actorDocument);
  const isDead = actorDocument.statuses?.has?.('dead');

  if (woundsValue <= 0 || isDead) {
    if (!actorDocument.statuses?.has?.('unconscious')) {
      await actorDocument.setCondition('unconscious', true);
    }
    if (!actorDocument.statuses?.has?.('prone')) {
      await actorDocument.setCondition('prone', true);
    }
    if (actorDocument.statuses?.has?.('dying')) {
      await actorDocument.setCondition('dying', false);
    }
    if (actorDocument.statuses?.has?.('staggered')) {
      await actorDocument.setCondition('staggered', false);
    }
    return;
  }

  // Wounded is represented by staggered in W&V; it is not unconscious by itself.
  if (state.isWounded && actorDocument.statuses?.has?.('dying')) {
    await actorDocument.setCondition('dying', false);
  }

  if (actorDocument.statuses?.has?.('unconscious')) {
    await actorDocument.setCondition('unconscious', false);
  }
}

export async function handleUnconsciousOnUpdate(actorDocument, change, options = {}) {
  if (isWoundsVigorActive(actorDocument)) {
    await handleWoundsVigorUnconscious(actorDocument, change);
    return;
  }

  const dyingAutomationEnabled = game.settings.get(MODULE.ID, "enableDyingAutomation");

  if (hasPrimaryHealthUpdate(actorDocument, change)) {
    const newHp = getPrimaryHealthValue(actorDocument, change);
    if (!dyingAutomationEnabled && newHp >= 0 && actorDocument.statuses?.has?.("dying")) {
      await actorDocument.setCondition("dying", false);
    }

    // Clear staggered when no longer negative HP (or if dead).
    if (newHp >= 0 || actorDocument.statuses?.has?.("dead")) {
      const nlState = getNonlethalThresholdState(actorDocument);
      if (actorDocument.statuses?.has?.("staggered") && !nlState.shouldStaggered && !nlState.shouldUnconscious) {
        await actorDocument.setCondition("staggered", false);
      }
    }

    // Diehard choice should not persist once the actor is no longer at negative HP.
    if (newHp >= 0 || actorDocument.statuses?.has?.("dead")) {
      const flag = getHtkFlag(actorDocument);
      if (flag?.diehardMode) {
        await actorDocument.setFlag(MODULE.ID, "htk", { ...flag, diehardMode: null });
      }
    }

    if (newHp < 0) {
      const flag = getHtkFlag(actorDocument);
      if (flag?.diehardMode && flag.diehardMode !== "fight") {
        await actorDocument.setFlag(MODULE.ID, "htk", { ...flag, diehardMode: null });
      }
    }
  }

  if (hasPrimaryHealthUpdate(actorDocument, change)) {
    const newHp = getPrimaryHealthValue(actorDocument, change);
    const nlState = getNonlethalThresholdState(actorDocument);
    if (newHp >= 0 && actorDocument.statuses.has("unconscious") && !nlState.shouldUnconscious) {
      await actorDocument.setCondition("unconscious", false);
    }
  }

  await syncNonlethalThresholdStates(actorDocument, change);

  const unconsciousSetting = game.settings.get(MODULE.ID, 'unconsciousAtNegativeHP');
  if (unconsciousSetting !== 'none' && hasPrimaryHealthUpdate(actorDocument, change)) {
    const newHp = getPrimaryHealthValue(actorDocument, change);
    if (newHp < 0) {
      const isNPC = actorDocument.type === 'npc';
      const shouldApply = 
        unconsciousSetting === 'everyone' || 
        (unconsciousSetting === 'npc' && isNPC) ||
        (unconsciousSetting === 'player' && !isNPC);
      
      if (shouldApply) {
        const conScore = Number(actorDocument.system?.abilities?.con?.total ?? actorDocument.system?.abilities?.con?.value ?? 0) || 0;
        if (conScore > 0 && newHp <= -conScore) {
          if (actorDocument.statuses?.has?.("staggered")) await actorDocument.setCondition("staggered", false);
          if (!dyingAutomationEnabled && actorDocument.statuses?.has?.("dying")) {
            await actorDocument.setCondition("dying", false);
          }
          await actorDocument.setCondition("unconscious", true);
          await actorDocument.setCondition("prone", true);
          return;
        }

        const prevHp = getPrevHpFromOptions(actorDocument, options);
        const droppedBelowZero = Number.isFinite(prevHp) ? (prevHp >= 0 && newHp < 0) : false;

        const htkItems = await detectHardToKillItems(actorDocument);
        const htkFlag = getHtkFlag(actorDocument);

        const hasHonorableStandActive = Boolean(htkItems.honorableStand) && actorDocument.effects?.some?.((e) => (e?.name ?? e?.label) === "Honorable Stand");
        const hasLastStandActive = Boolean(htkItems.lastStand) && actorDocument.effects?.some?.((e) => (e?.name ?? e?.label) === "Last Stand");
        const hasDeathsHostPossessed = Boolean(htkItems.deathsHost) && (actorDocument.statuses?.has?.("possessed") ?? false);

        const hasFerocity =
          Boolean(htkItems.ferocityUmr) ||
          Boolean(htkItems.ferocityOrc) ||
          Boolean(htkItems.ferociousResolve);
        const ferocityItem = htkItems.ferociousResolve || htkItems.ferocityOrc || htkItems.ferocityUmr || null;

        const diehardMode = htkFlag?.diehardMode === "fight" ? "fight" : null;
        const diehardFighting = diehardMode === "fight" && Boolean(htkItems.diehard);

        const fightOnRoundActive = Boolean(htkFlag?.fightOnRoundActive || htkFlag?.fightOnDropAfterTurn);

        const fightOnTempActive = Boolean(htkFlag?.fightOnTemp?.expiresAt && game.time?.worldTime < htkFlag.fightOnTemp.expiresAt);

        const stayConscious =
          hasHonorableStandActive ||
          hasLastStandActive ||
          hasDeathsHostPossessed ||
          hasFerocity ||
          diehardFighting ||
          fightOnRoundActive ||
          fightOnTempActive;

        if (droppedBelowZero) {
          if (hasFerocity && ferocityItem) {
            await showItemCardOnOwner(actorDocument, ferocityItem, options);
          }

          if (htkItems.diehard && !diehardMode) {
            const choice = await promptDiehardChoiceOnOwner(actorDocument);
            if (choice === "fight") {
              await actorDocument.setFlag(MODULE.ID, "htk", { ...(htkFlag ?? {}), diehardMode: "fight" });
              await showItemCardOnOwner(actorDocument, htkItems.diehard, options);
            } else if (choice === "unconscious") {
              const next = { ...(htkFlag ?? {}) };
              if (next.diehardMode) next.diehardMode = null;
              await actorDocument.setFlag(MODULE.ID, "htk", next);
            }
          }

          const orcFerocityItem = htkItems.orcFerocityOrc || htkItems.orcFerocityHalfOrc || htkItems.orcFerocityGrachukk;
          if (orcFerocityItem && Number(orcFerocityItem.system?.uses?.value ?? 0) > 0) {
            const useIt = await promptUseFeatureOnOwner(actorDocument, "NAS.conditions.sockets.OrcFerocityTitle", "NAS.conditions.sockets.OrcFerocityPrompt");
            if (useIt) {
              const used = await useItemOnOwner(actorDocument, orcFerocityItem, options);
              if (!used) {
                await deductOneDailyUse(orcFerocityItem);
                await showItemCardOnOwner(actorDocument, orcFerocityItem, options);
              }
              const triggeredDuringOwnTurn = game.combat?.combatant?.actorId === actorDocument.id;
              await actorDocument.setFlag(MODULE.ID, "htk", {
                ...(getHtkFlag(actorDocument) ?? {}),
                fightOnRoundPending: false,
                fightOnRoundActive: true,
                fightOnDropAfterTurn: true,
                fightOnSource: "orcFerocity",
                fightOnSkipTurnEnds: triggeredDuringOwnTurn ? 1 : 0,
              });
            }
          }
          if (htkItems.akitonianFerocity && Number(htkItems.akitonianFerocity.system?.uses?.value ?? 0) > 0) {
            const useIt = await promptUseFeatureOnOwner(actorDocument, "NAS.conditions.sockets.AkitonianFerocityTitle", "NAS.conditions.sockets.AkitonianFerocityPrompt");
            if (useIt) {
              const used = await useItemOnOwner(actorDocument, htkItems.akitonianFerocity, options);
              if (!used) {
                await deductOneDailyUse(htkItems.akitonianFerocity);
                await showItemCardOnOwner(actorDocument, htkItems.akitonianFerocity, options);
              }
              const triggeredDuringOwnTurn = game.combat?.combatant?.actorId === actorDocument.id;
              await actorDocument.setFlag(MODULE.ID, "htk", {
                ...(getHtkFlag(actorDocument) ?? {}),
                fightOnRoundPending: false,
                fightOnRoundActive: true,
                fightOnDropAfterTurn: true,
                fightOnSource: "akitonianFerocity",
                fightOnSkipTurnEnds: triggeredDuringOwnTurn ? 1 : 0,
              });
            }
          }

          if (htkItems.fightOn && Number(htkItems.fightOn.system?.uses?.value ?? 0) > 0) {
            const useIt = await promptUseFeatureOnOwner(actorDocument, "NAS.conditions.sockets.FightOnTitle", "NAS.conditions.sockets.FightOnPrompt");
            if (useIt) {
              const used = await useItemOnOwner(actorDocument, htkItems.fightOn, options);
              if (!used) {
                await deductOneDailyUse(htkItems.fightOn);
                await showItemCardOnOwner(actorDocument, htkItems.fightOn, options);
              }
              const conMod = Number(actorDocument.system?.abilities?.con?.mod ?? actorDocument.system?.abilities?.con?.total ?? 0) || 0;
              const amount = Math.max(0, conMod);
              const hp = actorDocument.system?.attributes?.hp;
              const currentTemp = Number(hp?.temp ?? 0) || 0;
              await actorDocument.update({ "system.attributes.hp.temp": currentTemp + amount });
              const expiresAt = (game.time?.worldTime ?? 0) + 60;
              await actorDocument.setFlag(MODULE.ID, "htk", { ...(getHtkFlag(actorDocument) ?? {}), fightOnTemp: { amount, expiresAt } });
            }
          }

          if (htkItems.resolve && Number(htkItems.resolve.system?.uses?.value ?? 0) > 0) {
            const useIt = await promptUseFeatureOnOwner(actorDocument, "NAS.conditions.sockets.ResolveUnstoppableTitle", "NAS.conditions.sockets.ResolveUnstoppablePrompt");
            if (useIt) {
              const used = await useItemOnOwner(actorDocument, htkItems.resolve, options);
              if (!used) {
                await deductOneDailyUse(htkItems.resolve);
                await showItemCardOnOwner(actorDocument, htkItems.resolve, options);
              }
              await actorDocument.setFlag(MODULE.ID, "htk", { ...(getHtkFlag(actorDocument) ?? {}), resolveUnstoppable: true });
            }
          }
        }

        const refreshedFlag = getHtkFlag(actorDocument);
        const refreshedDiehard = refreshedFlag?.diehardMode === "fight" && Boolean(htkItems.diehard);
        const refreshedFightOnTempActive = Boolean(refreshedFlag?.fightOnTemp?.expiresAt && game.time?.worldTime < refreshedFlag.fightOnTemp.expiresAt);
        const refreshedFightOnRoundActive = Boolean(refreshedFlag?.fightOnRoundActive || refreshedFlag?.fightOnDropAfterTurn);
        const refreshedStayConscious =
          hasHonorableStandActive ||
          hasLastStandActive ||
          hasDeathsHostPossessed ||
          hasFerocity ||
          refreshedDiehard ||
          refreshedFightOnRoundActive ||
          refreshedFightOnTempActive ||
          Boolean(refreshedFlag?.resolveUnstoppable);

        if (refreshedDiehard && newHp < 0 && !actorDocument.statuses?.has?.("dead")) {
          if (!actorDocument.statuses?.has?.("staggered")) await actorDocument.setCondition("staggered", true);
        }

        if (!dyingAutomationEnabled) {
          if (refreshedStayConscious) {
            if (actorDocument.statuses?.has?.("dying")) await actorDocument.setCondition("dying", false);
          } else if (newHp < 0 && !actorDocument.statuses?.has?.("dead") && !actorDocument.statuses?.has?.("dying")) {
            await actorDocument.setCondition("dying", true);
          }
        }

        if (!refreshedStayConscious) {
          if (actorDocument.statuses?.has?.("staggered")) await actorDocument.setCondition("staggered", false);
          if (refreshedFlag?.diehardMode) {
            await actorDocument.setFlag(MODULE.ID, "htk", { ...refreshedFlag, diehardMode: null });
          }
          await actorDocument.setCondition('unconscious', true);
          await actorDocument.setCondition('prone', true);
        } else {
          if (newHp <= 0 && !actorDocument.statuses?.has?.("dead") && !actorDocument.statuses?.has?.("unconscious")) {
            if (!actorDocument.statuses?.has?.("staggered")) {
              await actorDocument.setCondition("staggered", true);
            }

            const disableSetting = game.settings.get(MODULE.ID, 'disableAtZeroHP');
            if (disableSetting && !actorDocument.statuses?.has?.("disabled")) {
              await actorDocument.setCondition("disabled", true);
            }
          }
        }
      }
    }
  }
}



