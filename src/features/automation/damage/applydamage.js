import { MODULE } from '../../../common/module.js';
import { checkMassiveDamage } from '../../../integration/moduleSockets.js';
import { abilityDamageCalculation } from './ability.js';
import { elementalResistancesCalculation } from './elemental.js';
import { hardnessCalculation } from './hardness.js';
import { damageImmunityCalculation } from './immunity.js';
import { damageReductionCalculation } from './reduction.js';
import { sortDamage } from './sortdamage.js';
import { damageVulnerabilityCalculation } from './vulnerability.js';

export function customApplyDamage(originalApplyDamage, value, config, targetInfo) {
    canvas.tokens.controlled.forEach(token => {
        let totalDamage = 0;
        const maxHP = token.actor.system.attributes.hp.max; 
        const traits = token.actor.system.traits;
        const abilities = token.actor.system.abilities 
        const isClusteredShots = config.flags?.[MODULE.ID]?.clusteredShots != null;
        const eRes = traits.eres; 
        const conditionImmunities = traits.ci 
        const damageImmunities = traits.di; 
        const damageReductions = traits.dr; 
        const damageVulnerabilities = traits.dv; 
        const hardness = traits.hardness
        const messageId = targetInfo.id;
        const message = game.messages.get(messageId);
        let systemRolls = message.systemRolls;
        if(Object.keys(systemRolls).length == 0 && systemRolls.constructor == Object && message.rolls) {
            systemRolls = message.rolls;
        };
        const itemSource = message.itemSource;
        const itemType = itemSource?.type;
        const itemOptions = itemSource?.flags?.[MODULE.ID] || {};
        let damageMult = targetInfo.buttonType == "PF1.ApplyHalf" || targetInfo.buttonType == "Apply Half" ? 0.5 : 1;

        if (targetInfo.buttonType.includes("percentage")) {
            const percentageMatch = targetInfo.buttonType.match(/(\d+)%/);
            if (percentageMatch) {
                const percentageValue = parseInt(percentageMatch[1], 10);
                if (percentageValue === 50) {
                    damageMult = 0.5;
                } else if (percentageValue === 100) {
                    damageMult = 1;
                };
            };
        };

        let actionId = message.system?.action?.id;
        let itemOptionsForSort = {
            token,
            itemSource,
            itemType,
            damageMult,
            actionId,
            ammoItem: null
        };

        if (systemRolls?.attacks?.length > 0) { 
            let attackDamage;
            let attackName;
            if (isClusteredShots) {
                const clusteredShotsData = config.flags[MODULE.ID].clusteredShots;
                attackDamage = [
                    {
                        options: {
                            damageType: clusteredShotsData.damageTypes || []
                        },
                        total: clusteredShotsData.totalDamage
                    }
                ];
                let uniqueAmmoItems = [];
                if (Array.isArray(clusteredShotsData.criticalHits)) {
                    const seenIds = new Set();
                    for (const hit of clusteredShotsData.criticalHits) {
                        const ammo = hit.ammoItem;
                        if (ammo && ammo._id && !seenIds.has(ammo._id)) {
                            uniqueAmmoItems.push(ammo);
                            seenIds.add(ammo._id);
                        }
                    }
                }
                itemOptionsForSort = {
                    ...itemOptionsForSort,
                    globalItemSettings: itemOptions.globalItemSettings,
                    itemActionSettings: itemOptions.itemActionSettings,
                    attackSettings: null, 
                    ammoItem: uniqueAmmoItems
                };
            } else {
                const attack = systemRolls.attacks[targetInfo.attackIndex];
                attackName = attack?.attack?.options?.flavor || "";
                if (attack.damage?.length > 0) {
                    attackDamage = targetInfo.isCritical ? JSON.parse(JSON.stringify([...attack.damage, ...attack.critDamage])) : JSON.parse(JSON.stringify(attack.damage));
                };

                let attackSettings = null;
                let actionSettings = null;
                let foundAction = null;
                let ammoItem = null;
                if (attack?.ammo?.id) {
                    ammoItem = itemSource.parent.items.get(attack.ammo.id);
                }
                if (itemOptions?.itemActionSettings?.actions?.length > 0 && actionId) {
                    foundAction = itemOptions.itemActionSettings.actions.find(a => a.id === actionId);
                    if (foundAction && foundAction.attacks && Array.isArray(foundAction.attacks)) {
                        attackSettings = foundAction.attacks.find(a => a.name === attackName);
                    }
                    actionSettings = JSON.parse(JSON.stringify(itemOptions.itemActionSettings));
                    for (const action of actionSettings.actions) {
                        if ('attacks' in action) {
                            delete action.attacks;
                        }
                    }
                }
                itemOptionsForSort = {
                    ...itemOptionsForSort,
                    globalItemSettings: itemOptions.globalItemSettings,
                    itemActionSettings: actionSettings || itemOptions.itemActionSettings,
                    attackSettings: attackSettings,
                    ammoItem: ammoItem
                };
            };
            const {damageSortObjects, damageTypes, itemAction, abilityDmg} = sortDamage(attackDamage, itemOptionsForSort);
            damageImmunityCalculation(damageImmunities, attackDamage, damageSortObjects, itemOptionsForSort);
            damageVulnerabilityCalculation(damageVulnerabilities, attackDamage, damageSortObjects);
            hardnessCalculation(attackDamage, eRes, damageReductions, damageTypes, damageSortObjects, itemSource, hardness, itemOptionsForSort);
            elementalResistancesCalculation(eRes, attackDamage, damageTypes, damageSortObjects, itemOptionsForSort);
            damageReductionCalculation(attackDamage, damageReductions, damageTypes, damageSortObjects, itemSource, itemAction, message, hardness, itemOptionsForSort);
            abilityDamageCalculation(damageImmunities, conditionImmunities, abilities, abilityDmg);
            attackDamage.forEach(damage => {
                const damageTypes = damage.options?.damageType || [];
                if(damageTypes.length < 1) {
                    damageTypes[0] = "untyped";
                };
                const type = damageTypes[0];
                if (!abilityDmg.some(dmgType => dmgType.type === type) || abilityDmg.length == 0) { 
                    let damageForType = damage.total || 0; 

                    totalDamage += Math.max(0, damageForType);
                    
                    checkMassiveDamage(damageForType, maxHP, token);
                } else if (abilityDmg && abilityDmg.length > 0) {
                    let updates = {};
                    for (const key in abilities) {
                        updates[`system.abilities.${key}.damage`] = abilities[key].damage || 0;
                        updates[`system.abilities.${key}.drain`] = abilities[key].drain || 0;
                        updates[`system.abilities.${key}.userPenalty`] = abilities[key].userPenalty || 0;
                    };
                    token.actor.update(updates);
                };
            });
        } else {
            systemRolls.forEach(roll => {
                const attackDamage = JSON.parse(JSON.stringify(roll?.terms));

                const hasOptions = roll?.options && Object.keys(roll.options).length > 0;
                
                if (hasOptions) {
                  attackDamage.forEach(term => {
                    if (term.class === "Die" || term.class === "NumericTerm") {
                      term.options = {
                        ...roll.options,
                        ...term.options
                      };
                    };
                  });
                };
                const {damageSortObjects, damageTypes} = sortDamage(attackDamage, itemOptionsForSort);
                damageImmunityCalculation(damageImmunities, attackDamage, damageSortObjects);
                damageVulnerabilityCalculation(damageVulnerabilities, attackDamage, damageSortObjects);
                hardnessCalculation(attackDamage, eRes, damageReductions, damageTypes, damageSortObjects, itemSource, hardness);
                elementalResistancesCalculation(eRes, attackDamage, damageTypes, damageSortObjects);
                damageReductionCalculation(attackDamage, damageReductions, damageTypes, damageSortObjects, itemSource, hardness);
        
                attackDamage.forEach(damage => {
                    const flags = message.system?.subject?.health || message.flags?.[MODULE.ID]?.subject?.health || message.flags?.pf1?.subject?.health;
                    const healthFlag = (flags === undefined || flags === null || Object.keys(message.flags).length === 0 || flags === "damage") ? 1 : -1;
                    let rolledDamage = (damage.total * healthFlag) || 0;

                    totalDamage += rolledDamage;
                    
                    checkMassiveDamage(rolledDamage, maxHP, token);
                });
            });
        };
        
        return originalApplyDamage(totalDamage, config);
    });
    
    Object.keys(targetInfo).forEach(key => delete targetInfo[key]);
};



