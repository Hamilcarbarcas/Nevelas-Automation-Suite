import { MODULE } from '../../../common/module.js';
import { elementFromHtmlLike } from '../../../common/foundryCompat.js';

export function reorderTokenHUDConditions(html, data) {
  const root = elementFromHtmlLike(html);
  const conditions = root?.querySelectorAll?.('.status-effects') ?? [];
  const reorderAllConditions = game.settings.get(MODULE.ID, 'reorderAllConditions');
  const allConditions = pf1.registry.conditions.map(condition => condition._id);

  const conditionEffects = Object.values(data.statusEffects).filter(effect => allConditions.includes(effect.id));
  const buffEffects = Object.values(data.statusEffects).filter(effect => !allConditions.includes(effect.id) && effect.id !== "dead");
  const deadCondition = Object.values(data.statusEffects).filter(effect => effect.id === "dead");

  let sortedEffects;
  if (reorderAllConditions) {
    sortedEffects = Object.values(data.statusEffects).sort((a, b) => a.title.localeCompare(b.title));
  } else {
    const otherConditions = conditionEffects.filter(effect => effect.id !== "dead");
    sortedEffects = otherConditions.sort((a, b) => a.title.localeCompare(b.title));
  }

  conditions.forEach(el => {
      const icons = Array.from(el.querySelectorAll('img'));
      const iconMap = {};
      icons.forEach(img => {
        iconMap[img.dataset.statusId] = img;
      });
      let newOrder = [];
      if (deadCondition && !reorderAllConditions && deadCondition[0]) {
        if (iconMap[deadCondition[0].id]) newOrder.push(iconMap[deadCondition[0].id]);
      }
      sortedEffects.forEach(effect => {
        if (iconMap[effect.id]) newOrder.push(iconMap[effect.id]);
      });
      if (!reorderAllConditions) {
        buffEffects.forEach(effect => {
          if (iconMap[effect.id]) newOrder.push(iconMap[effect.id]);
        });
      }
      newOrder.forEach(img => el.appendChild(img));
  });
}
