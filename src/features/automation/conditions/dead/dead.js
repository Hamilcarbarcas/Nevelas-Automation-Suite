import { MODULE } from '../../../../common/module.js';
import { getPrimaryHealthValue, hasPrimaryHealthUpdate } from '../../utils/healthUpdates.js';
import { isWoundsVigorActive } from '../../utils/woundsVigor.js';

export async function handleDeadOnUpdate(actorDocument, change) {
  const deadConditionSetting = game.settings.get(MODULE.ID, 'applyDeadCondition');
  const isWv = isWoundsVigorActive(actorDocument);

  if (deadConditionSetting !== 'none' && hasPrimaryHealthUpdate(actorDocument, change)) {
    const newHealth = getPrimaryHealthValue(actorDocument, change);
    const conScore = actorDocument.system.abilities.con.total;
    const isNPC = actorDocument.type === 'npc';
    
    let shouldApply = false;
    
    if (deadConditionSetting === 'everyone') {
      shouldApply = isWv ? newHealth <= 0 : newHealth <= -conScore;
    } else if (deadConditionSetting === 'npc' && isNPC) {
      shouldApply = isWv ? newHealth <= 0 : newHealth <= -conScore;
    } else if (deadConditionSetting === 'player' && !isNPC) {
      shouldApply = isWv ? newHealth <= 0 : newHealth <= -conScore;
    } else if (deadConditionSetting === 'player-negative-con-npc-negative-hp') {
      if (isNPC) {
        shouldApply = isWv ? newHealth <= 0 : newHealth < 0;
      } else {
        shouldApply = isWv ? newHealth <= 0 : newHealth <= -conScore;
      }
    }
    
    if (shouldApply) {
      await actorDocument.setCondition('dead', {overlay: true});
      await actorDocument.setCondition('prone', true); 
      await actorDocument.setCondition('dying', false);
      await actorDocument.setCondition('staggered', false);
      await actorDocument.setCondition('unconscious', true);
    }
  }

  const removeDeadSetting = game.settings.get(MODULE.ID, 'removeDeadCondition');
  if (removeDeadSetting !== 'disabled' && hasPrimaryHealthUpdate(actorDocument, change) && actorDocument.statuses?.has?.('dead')) {
    const newHealth = getPrimaryHealthValue(actorDocument, change);
    const conScore = actorDocument.system.abilities.con.total;

    let shouldRemove = false;
    if (removeDeadSetting === 'aboveNegativeCon') {
      shouldRemove = isWv ? newHealth > 0 : newHealth > -conScore;
    } else if (removeDeadSetting === 'nonNegative') {
      shouldRemove = isWv ? newHealth > 0 : newHealth >= 0;
    }

    if (shouldRemove) {
      await actorDocument.setCondition('dead', false);
    }
  }
}



