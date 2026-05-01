import { MODULE } from '../../common/module.js';
import { elementFromHtmlLike } from '../../common/foundryCompat.js';

export function onRenderChatMessage(html) {
    const root = elementFromHtmlLike(html);
    const messages = root?.querySelectorAll('div.chat-attack');
    if (!messages?.length) return;
    messages.forEach(message => {
        const rows = Array.from(message.querySelectorAll('tr'));
        const normalDamageInfo = [];
        const criticalDamageInfo = [];
        for (let i = 0; i < rows.length; i++) {
            const row = rows[i];
            if (row.classList.contains('damage')) {
                let j = i + 1;
                let foundComponent = false;
                while (j < rows.length && rows[j].querySelector('td.roll.damage.normal')) {
                    foundComponent = true;
                    const rollCell = rows[j].querySelector('td.roll.damage.normal a.inline-roll');
                    const typeCell = rows[j].querySelector('td.damage-types');
                    let value = null, types = [];
                    if (rollCell) value = parseInt(rollCell.textContent.trim(), 10);
                    if (typeCell) {
                        types = Array.from(typeCell.querySelectorAll('.damage-type, .custom')).map(dt =>
                            dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim()
                        );
                    }
                    if (value !== null && types.length > 0) {
                        normalDamageInfo.push({ damageType: types, totalDamage: value });
                    }
                    j++;
                }
                if (!foundComponent) {
                    const normalRollElement = row.querySelector('td.roll.damage.normal a[data-tooltip]');
                    let normalDamageTypes = [];
                    const normalTDs = row.querySelectorAll('td.damage-types');
                    if (normalTDs.length > 0) {
                        normalDamageTypes = Array.from(normalTDs).flatMap(td =>
                            Array.from(td.querySelectorAll('.damage-type, .custom')).map(dt =>
                                dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim()
                            )
                        );
                    } else {
                        normalDamageTypes = Array.from(row.querySelectorAll('td.damage-types .damage-type, td.damage-types .custom'))
                            .map(dt => dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim());
                    }
                    if (normalRollElement && normalDamageTypes.length > 0) {
                        const totalDamage = parseInt(normalRollElement.textContent.trim(), 10);
                        normalDamageInfo.push({ damageType: normalDamageTypes, totalDamage });
                    }
                    const criticalRollElement = row.querySelector('td.roll.damage.critical a[data-tooltip]');
                    let criticalDamageTypes = [];
                    const criticalTDs = row.querySelectorAll('td.damage-type');
                    if (criticalTDs.length > 0) {
                        criticalDamageTypes = Array.from(criticalTDs).flatMap(td =>
                            Array.from(td.querySelectorAll('.damage-type, .custom')).map(dt =>
                                dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim()
                            )
                        );
                    } else {
                        criticalDamageTypes = Array.from(row.querySelectorAll('td.damage-type .damage-type, td.damage-type .custom'))
                            .map(dt => dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim());
                    }
                    if (criticalRollElement && criticalDamageTypes.length > 0) {
                        const totalDamage = parseInt(criticalRollElement.textContent.trim(), 10);
                        criticalDamageInfo.push({ damageType: criticalDamageTypes, totalDamage });
                    }
                }
            }
        }

        const sections = message.querySelectorAll('tr.damage > th, th.attack-damage');

        sections.forEach((section, index) => {
            const applyDamageElements = section.querySelectorAll('a.inline-action[data-action="applyDamage"]');

            applyDamageElements.forEach(element => {
                const tooltip = element.getAttribute('data-tooltip');
                if (tooltip === 'PF1.ApplyHealing') {
                    element.setAttribute('data-tooltip', 'PF1.ApplyDamage');
                }
            });
            
            const heal = document.createElement('div');
            heal.innerHTML = "❤️";
            const healHalf = document.createElement('div');
            healHalf.innerHTML = "🩹";
            modifyElementStyles(heal, true);
            modifyElementStyles(healHalf);
            modifyElementAttributes(heal, "Heal");
            modifyElementAttributes(healHalf, "Heal Half");
            section.appendChild(heal);
            section.appendChild(healHalf);
    
            message.addEventListener('mouseenter', () => {
                heal.style.visibility = "visible";
                healHalf.style.visibility = "visible";
            });
    
            message.addEventListener('mouseleave', () => {
                heal.style.visibility = "hidden";
                healHalf.style.visibility = "hidden";
            });
    
            const isCritical = section.getAttribute('data-damage-type') === 'critical';
    
            heal.addEventListener('click', () => {
                if (isCritical) {
                    applyHealing([...normalDamageInfo, ...criticalDamageInfo], 1);
                } else {
                    applyHealing(normalDamageInfo, 1);
                }
            });
    
            healHalf.addEventListener('click', () => {
                if (isCritical) {
                    applyHealing([...normalDamageInfo, ...criticalDamageInfo], 0.5);
                } else {
                    applyHealing(normalDamageInfo, 0.5);
                }
            });
        });
    });
}


function applyHealing(damageInfo, multiplier) {
    let healDamage = 0;

    damageInfo.forEach(({ damageType, totalDamage }) => {
        let abilityHealingApplied = false;
        damageType.forEach(dt => {
            for (const [key, value] of pf1.registry.damageTypes.entries()) {
                if (dt === value.name) {
                    let healAmount = totalDamage * -1 * multiplier;
                    if (multiplier === 0.5) {
                        healAmount = Math.ceil(healAmount);
                    }

                    if (value.flags?.[MODULE.ID]?.vsAbility) {
                        const ability = value.flags?.[MODULE.ID]?.abilities;
                        const ablDmgType = value.flags?.[MODULE.ID]?.type;

                        canvas.tokens.controlled.forEach(token => {
                            const tokenAbilities = token.actor.system.abilities;
                            const dmg = {
                                vs: ability,
                                amount: healAmount,
                                ablDmgType: ablDmgType
                            };
                            
                            if (tokenAbilities.hasOwnProperty(dmg.vs) && dmg.amount < 0) {
                                switch (dmg.ablDmgType) {
                                    case "damage":
                                        tokenAbilities[dmg.vs].damage = Math.max(tokenAbilities[dmg.vs].damage + dmg.amount, 0);
                                        break;
                                    case "drain":
                                        tokenAbilities[dmg.vs].drain = Math.max(tokenAbilities[dmg.vs].drain + dmg.amount, 0);
                                        break;
                                    case "penalty":
                                        tokenAbilities[dmg.vs].userPenalty = Math.max(tokenAbilities[dmg.vs].userPenalty + dmg.amount, 0);
                                        break;
                                }
                                abilityHealingApplied = true;
                            }

                            let updates = {};
                            for (const key in tokenAbilities) {
                                updates[`system.abilities.${key}.damage`] = tokenAbilities[key].damage;
                                updates[`system.abilities.${key}.drain`] = tokenAbilities[key].drain;
                                updates[`system.abilities.${key}.userPenalty`] = tokenAbilities[key].userPenalty;
                            }
                            token.actor.update(updates);
                        });
                    }
                }
            }
        });
        if (!abilityHealingApplied) {
            healDamage += totalDamage;
        }
    });

    if (healDamage > 0) {
        healDamage = healDamage * -1 * multiplier;
        if (multiplier === 0.5) {
            healDamage = Math.ceil(healDamage);
        }
        const originalApplyDamage = pf1.documents.actor.ActorPF.applyDamage;
        originalApplyDamage(healDamage, { asNonLethal: false, healing: true });
    }
}

function modifyElementStyles (element, pulsating=false) {
    element.style.visibility="hidden";
    element.style.display="inline-block";

    if(pulsating) {
        const keyframes = [
            { transform: 'scale(1)' },
            { transform: 'scale(1.1)' },
            { transform: 'scale(1)' },
          ];
          
          const options = {
            duration: 600, 
            iterations: Infinity,
            easing: 'ease-in-out'
          };

          const animation = element.animate(keyframes, options);
          animation.pause()

          element.addEventListener('mouseenter',()=>{animation.play()});
          element.addEventListener('mouseleave',()=>{animation.pause()});

    } else {
        element.addEventListener('mouseenter',()=>{element.style.transform = 'scale(1.1)'});
        element.style.transition = "transform 150ms";
        element.addEventListener('mouseleave',()=>{element.style.transform = 'scale(1)'});
    };
};

function modifyElementAttributes (element, tooltipText) {
    element.setAttribute("data-tooltip", tooltipText);
};
