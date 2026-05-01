import { MODULE } from '../../../../common/module.js';
import { socket, getPreferredOwnerUserId } from '../../../../integration/moduleSockets.js';
import { getPrimaryHealthValue, hasPrimaryHealthUpdate } from '../../utils/healthUpdates.js';
import { detectHardToKillItems, getHtkFlag } from '../../utils/hardToKill.js';
import { isWoundsVigorActive, shouldBeStaggeredFromWv } from '../../utils/woundsVigor.js';

export async function handleDisabledOnUpdate(actorDocument, change) {
  const disableSetting = game.settings.get(MODULE.ID, 'disableAtZeroHP');
  const actorType = actorDocument.type;

  if ((disableSetting === 'everyone' || 
     (disableSetting === 'player' && actorType === 'character') || 
     (disableSetting === 'npc' && actorType === 'npc')) && 
     hasPrimaryHealthUpdate(actorDocument, change)) {
      const isWv = isWoundsVigorActive(actorDocument);
      const newHealth = getPrimaryHealthValue(actorDocument, change);
      const shouldDisable = isWv
        ? shouldBeStaggeredFromWv(actorDocument)
        : newHealth === 0;
      if (shouldDisable) {
        await actorDocument.setCondition("disabled", true);
    }
  }
}

export async function handleDisabledStrenuousAction(action) {
  const itemSource = action.item;
  const token = action.token;
  const actor = token?.actor;

  if (game.settings.get(MODULE.ID, 'disableAtZeroHP')) {
    if (isWoundsVigorActive(actor)) return;
    const activationTypes = ["nonaction", "passive", "free", "swift", "immediate", "move", "standard", "full", "attack", "aoo", "round", "minute", "hour", "special"]
    const strenuousTypes = ["standard", "full", "attack", "aoo", "round", "minute", "hour"]
    if (strenuousTypes.includes(action.action?.activation?.type)) {
      if (actor && actor?.statuses?.has("disabled")) {
        let hp = actor.system?.attributes?.hp;
        const conScore = actor.system?.abilities?.con?.total;
        let newHp = hp.value - 1;
        
        if(hp.max > 0 && hp.value == 0 && !actor.statuses.has("unconscious")) {
          await handleHTK(actor, newHp, conScore);
        } else if(hp.max > 0 && hp.value < 0 && !actor.statuses.has("unconscious") && hp.value >= (conScore * -1)) {
          await handleHTK(actor, newHp, conScore);
        }
      }
    }
  }
}

async function handleHTK(actor, newHp, conScore) {
  const htkItems = await detectHardToKillItems(actor);
  const htkFlag = getHtkFlag(actor);

  const diehardActive = htkFlag?.diehardMode === "fight" && Boolean(htkItems.diehard);
  const hasFerocity =
    Boolean(htkItems.ferocityUmr) ||
    Boolean(htkItems.ferocityOrc) ||
    Boolean(htkItems.ferociousResolve);

  const hasDeathlessMaster = Boolean(htkItems.deathlessMaster);
  if (hasDeathlessMaster && (diehardActive || actor.system?.attributes?.hp?.value <= 0)) {
    return;
  }

  if (newHp < 0 && newHp > -conScore) {
    if (!diehardActive) {
      await actor.update({"system.attributes.hp.value": newHp});
      return;
    }

    let choice = null;
    if (socket) {
      const ownerId = getPreferredOwnerUserId(actor);
      if (ownerId) {
        choice = await socket.executeAsUser("promptHTKChoice", ownerId, actor.id);
      }
    }

    if (choice === "fight") {
      await actor.setFlag(MODULE.ID, "htk", { ...(htkFlag ?? {}), diehardMode: "fight" });
      await actor.update({"system.attributes.hp.value": newHp});
      return;
    }

    if (htkFlag?.diehardMode) {
      await actor.setFlag(MODULE.ID, "htk", { ...(htkFlag ?? {}), diehardMode: null });
    }
    await actor.update({"system.attributes.hp.value": newHp});
  } else if (newHp <= -conScore) {
    await actor.update({"system.attributes.hp.value": newHp});
  }
}



