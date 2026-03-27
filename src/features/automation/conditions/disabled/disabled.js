import { MODULE } from '../../../../common/module.js';
import { socket } from '../../../../integration/moduleSockets.js';
import { getNewHp, hasHpUpdate } from '../../utils/healthUpdates.js';

export async function handleDisabledOnUpdate(actorDocument, change) {
  const disableSetting = game.settings.get(MODULE.ID, 'disableAtZeroHP');
  const actorType = actorDocument.type;  

  if ((disableSetting === 'everyone' || 
     (disableSetting === 'player' && actorType === 'character') || 
     (disableSetting === 'npc' && actorType === 'npc')) && 
    hasHpUpdate(change)) {
    const newHp = getNewHp(actorDocument, change);
    if (newHp === 0) {
        await actorDocument.setCondition("disabled", true);
    }
  }
}

export async function handleDisabledStrenuousAction(action) {
  const itemSource = action.item;
  const token = action.token;
  const actor = token?.actor;

  if (game.settings.get(MODULE.ID, 'disableAtZeroHP')) {
    const activationTypes = ["nonaction", "passive", "free", "swift", "immediate", "move", "standard", "full", "attack", "aoo", "round", "minute", "hour", "special"]
    const strenuousTypes = ["standard", "full", "attack", "aoo", "round", "minute", "hour"]
    if (strenuousTypes.includes(action.action?.activation?.type)) {
      if (actor && actor?.statuses?.has("disabled")) {
        let hp = actor.system?.attributes?.hp;
        const conScore = actor.system?.abilities?.con?.total;
        const hardToKill = ["diehard", "ferocity (orc)", "orc ferocity", "ferocity"]
        const ability = actor.items.find(item => hardToKill.some(htk => htk === item.name.toLowerCase()));
        let newHp = hp.value - 1;
        
        if(hp.max > 0 && hp.value == 0 && !actor.statuses.has("unconscious")) {
          handleHTK(actor, ability, newHp, conScore);
        } else if(hp.max > 0 && hp.value < 0 && !actor.statuses.has("unconscious") && hp.value >= (conScore * -1)) {
          handleHTK(actor, ability, newHp, conScore);
        }
      }
    }
  }
}

async function handleHTK(actor, ability, newHp, conScore) {
  if (newHp < 0 && newHp > -conScore) {
    const choice = await socket.executeAsGM("promptHTKChoice", actor.id);
    if (choice === "fight") {
      await actor.update({"system.attributes.hp.value": newHp});
      return;
    } else {
      await actor.update({"system.attributes.hp.value": newHp});
    }
  } else if (newHp <= -conScore) {
    await actor.update({"system.attributes.hp.value": newHp});
  }
}



