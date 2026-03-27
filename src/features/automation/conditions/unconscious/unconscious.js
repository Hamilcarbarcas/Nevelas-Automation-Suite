import { MODULE } from '../../../../common/module.js';
import { getNewHp, hasHpUpdate } from '../../utils/healthUpdates.js';

export async function handleUnconsciousOnUpdate(actorDocument, change) {
  if (hasHpUpdate(change)) {
    const newHp = getNewHp(actorDocument, change);
    if (newHp >= 0 && actorDocument.statuses.has("unconscious")) {
      await actorDocument.setCondition("unconscious", false);
    }
  }

  const unconsciousSetting = game.settings.get(MODULE.ID, 'unconsciousAtNegativeHP');
  if (unconsciousSetting !== 'none' && hasHpUpdate(change)) {
    const newHp = getNewHp(actorDocument, change);
    if (newHp < 0) {
      const isNPC = actorDocument.type === 'npc';
      const shouldApply = 
        unconsciousSetting === 'everyone' || 
        (unconsciousSetting === 'npc' && isNPC) ||
        (unconsciousSetting === 'player' && !isNPC);
      
      if (shouldApply) {
        const hardToKill = ["diehard", "ferocity (orc)", "orc ferocity", "ferocity"];
        const hasHTK = actorDocument.items.some(i => 
          i.type === 'feat' && 
          hardToKill.some(htk => htk === i.name.toLowerCase())
        );
        
        if (!hasHTK) {
          await actorDocument.setCondition('unconscious', true);
          await actorDocument.setCondition('prone', true); 
        }
      }
    }
  }
}



