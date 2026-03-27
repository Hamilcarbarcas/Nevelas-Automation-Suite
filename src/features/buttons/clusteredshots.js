import { MODULE } from '../../common/module.js';

export async function addClusteredShotsButton(html) {
    const root = (typeof jQuery !== 'undefined' && html instanceof jQuery) ? html[0] : html;
    const cards = root?.querySelectorAll('div.pf1.chat-card.item-card, div.chat-card.item-card.pf1');
    
    if (cards?.length) {
        for (const card of cards) {
            const tokenUuid = card.getAttribute('data-token-uuid');
            const itemId = card.getAttribute('data-item-id');
            const actionId = card.getAttribute('data-action-id');
            let shouldAddButton = false;
            if (tokenUuid && itemId && typeof fromUuid === 'function') {
                try {
                    const token = await fromUuid(tokenUuid);
                    const item = token?.actor?.items?.get(itemId);
                    let action = undefined;
                    if (item && item.actions && typeof item.actions.get === 'function' && actionId) {
                        action = item.actions.get(actionId);
                        if (action && action.isRanged) {
                            shouldAddButton = true;
                        }
                    }
                } catch (e) {
                }
            }
            if (!shouldAddButton) {
                card._skipClusteredShots = true;
            }
        }
    }
    const filteredCards = Array.from(cards).filter(card => {
        if (card._skipClusteredShots) return false;
        const chatAttacks = card.querySelectorAll('.chat-attack');
        return chatAttacks.length >= 2;
    });
    if (!filteredCards.length) {
        return;
    }
    filteredCards.forEach(card => {
        const chatAttacks = card.querySelectorAll('.chat-attack');
        if (!chatAttacks.length) return;
        
        const chatMessage = card.closest('.chat-message');
        if (!chatMessage || !chatMessage.getAttribute('data-message-id')) {
            return;
        }
        
        chatAttacks.forEach((attack, index) => {
            attack.style.position = 'relative';
            
            const hasCriticalDamage = attack.querySelector('tr.damage th[data-damage-type="critical"]') !== null;
            
            const checkboxContainer = document.createElement('div');
            checkboxContainer.className = 'cs-checkbox-container';
            checkboxContainer.style.position = 'absolute';
            checkboxContainer.style.right = '10px';
            checkboxContainer.style.top = '5px';
            checkboxContainer.style.visibility = 'hidden'; 
            checkboxContainer.style.zIndex = '100';
            checkboxContainer.style.backgroundColor = 'rgba(0, 0, 0, 0.6)';
            checkboxContainer.style.padding = '3px 6px';
            checkboxContainer.style.borderRadius = '3px';
            checkboxContainer.style.display = 'flex';
            checkboxContainer.style.flexDirection = 'column';
            checkboxContainer.style.gap = '3px';
            
            const normalCheckbox = document.createElement('input');
            normalCheckbox.type = 'checkbox';
            normalCheckbox.className = 'cs-attack-checkbox cs-normal-checkbox';
            normalCheckbox.dataset.attackIndex = attack.getAttribute('data-index');
            normalCheckbox.dataset.damageType = 'normal';
            normalCheckbox.checked = false; 
            normalCheckbox.style.cursor = 'pointer';
            normalCheckbox.style.verticalAlign = 'middle';
            
            const normalLabel = document.createElement('label');
            normalLabel.htmlFor = `cs-normal-checkbox-${index}`;
            normalLabel.style.marginLeft = '4px';
            normalLabel.style.cursor = 'pointer';
            normalLabel.style.color = 'white';
            normalLabel.style.fontSize = '12px';
            normalLabel.textContent = 'Normal';
            
            const normalContainer = document.createElement('div');
            normalContainer.appendChild(normalCheckbox);
            normalContainer.appendChild(normalLabel);
            
            checkboxContainer.setAttribute('data-tooltip', 'Include in Clustered Shots');
            
            checkboxContainer.appendChild(normalContainer);
            
            if (hasCriticalDamage) {
                const criticalCheckbox = document.createElement('input');
                criticalCheckbox.type = 'checkbox';
                criticalCheckbox.className = 'cs-attack-checkbox cs-critical-checkbox';
                criticalCheckbox.dataset.attackIndex = attack.getAttribute('data-index');
                criticalCheckbox.dataset.damageType = 'critical';
                criticalCheckbox.checked = false; 
                criticalCheckbox.style.cursor = 'pointer';
                criticalCheckbox.style.verticalAlign = 'middle';
                
                const criticalLabel = document.createElement('label');
                criticalLabel.htmlFor = `cs-critical-checkbox-${index}`;
                criticalLabel.style.marginLeft = '4px';
                criticalLabel.style.cursor = 'pointer';
                criticalLabel.style.color = 'white';
                criticalLabel.style.fontSize = '12px';
                criticalLabel.textContent = 'Critical';
                
                const criticalContainer = document.createElement('div');
                criticalContainer.appendChild(criticalCheckbox);
                criticalContainer.appendChild(criticalLabel);
                
                checkboxContainer.appendChild(criticalContainer);
            }
            
            attack.appendChild(checkboxContainer);
            
            attack.addEventListener('mouseenter', () => {
                checkboxContainer.style.visibility = 'visible';
            });
            
            attack.addEventListener('mouseleave', () => {
                const anyChecked = checkboxContainer.querySelectorAll('input:checked').length > 0;
                if (!anyChecked) {
                    checkboxContainer.style.visibility = 'hidden';
                }
            });
            
            checkboxContainer.querySelectorAll('input[type="checkbox"]').forEach(checkbox => {
                checkbox.addEventListener('change', () => {
                    const anyChecked = checkboxContainer.querySelectorAll('input:checked').length > 0;
                    if (anyChecked) {
                        checkboxContainer.style.visibility = 'visible';
                    } else {
                        if (!attack.matches(':hover')) {
                            checkboxContainer.style.visibility = 'hidden';
                        }
                    }
                });
            });
        });
        
        const topButton = document.createElement('div');
        const bottomButton = document.createElement('div');
        
        topButton.innerHTML = "🎯 Clustered Shots 🎯";
        bottomButton.innerHTML = "🎯 Clustered Shots 🎯";
        
        styleClusteredShotButton(topButton);
        styleClusteredShotButton(bottomButton);
        
        [topButton, bottomButton].forEach(button => {
            button.dataset.action = "applyClusteredDamage";
            button.dataset.type = "normal";
            button.dataset.ratio = "1";
            button.dataset.tags = "";
            button.dataset.clusteredShots = "true"; 
        });
        
        topButton.setAttribute("data-tooltip", "Apply damage before DR (Clustered Shots)");
        bottomButton.setAttribute("data-tooltip", "Apply damage before DR (Clustered Shots)");
        
        const firstAttack = chatAttacks[0];
        const lastAttack = chatAttacks[chatAttacks.length - 1];
        
        firstAttack.parentNode.insertBefore(topButton, firstAttack);
        lastAttack.parentNode.insertBefore(bottomButton, lastAttack.nextSibling);
        
        [topButton, bottomButton].forEach(button => {
            button.addEventListener('click', (event) => {
                applyClusteredShots(card, event.currentTarget, event);
            });
        });
    });
}

async function applyClusteredShots(card, button, event) {
    const chatMessage = card.closest('.chat-message');
    const messageId = chatMessage ? chatMessage.getAttribute('data-message-id') : null;
    
    if (!messageId) {
        ui.notifications.error("Could not find message ID.");
        return;
    }
    
    const clickedButton = button;
    
    const chatAttacks = card.querySelectorAll('.chat-attack');
    if (!chatAttacks.length) {
        ui.notifications.warn("No attacks found in this message.");
        return;
    }
    
    const messageObject = game.messages.get(messageId);
    
    if (!messageObject) {
        ui.notifications.error("Could not find message data.");
        return;
    }
    
    const anyAttackSelected = Array.from(chatAttacks).some(attack => {
        const attackIndex = attack.getAttribute('data-index');
        const checkboxes = attack.querySelectorAll(`.cs-attack-checkbox[data-attack-index="${attackIndex}"]`);
        return Array.from(checkboxes).some(checkbox => checkbox.checked);
    });
    
    if (!anyAttackSelected) {
        ui.notifications.warn("No attacks selected for Clustered Shots. Please check at least one attack to include.");
        return;
    }
    
    const criticalHits = [];

    let fallbackTotalDamage = 0;
    const fallbackDamageTypes = new Set();
    const selectedAttacks = [];

    for (const attack of chatAttacks) {
        const attackIndex = attack.getAttribute('data-index');

        const normalCheckbox = attack.querySelector(`.cs-normal-checkbox[data-attack-index="${attackIndex}"]`);
        const includeNormalDamage = normalCheckbox && normalCheckbox.checked;

        const criticalCheckbox = attack.querySelector(`.cs-critical-checkbox[data-attack-index="${attackIndex}"]`);
        const includeCriticalDamage = criticalCheckbox && criticalCheckbox.checked;

        if (!includeNormalDamage && !includeCriticalDamage) {
            continue;
        }

        selectedAttacks.push({
            attackIndex: Number(attackIndex),
            includeNormalDamage,
            includeCriticalDamage
        });

        let normalDamage = 0;
        if (includeNormalDamage) {
            const damageRow = attack.querySelector('tr.damage th:not([data-damage-type="critical"])');
            if (damageRow) {
                const damageValue = damageRow.querySelector('a.fake-inline-roll');
                if (damageValue) {
                    normalDamage = parseInt(damageValue.textContent.trim(), 10);
                    if (!isNaN(normalDamage)) {
                        fallbackTotalDamage += normalDamage;
                    } else {
                        normalDamage = 0;
                    }
                }
            }
        }

        let criticalDamage = 0;
        if (includeCriticalDamage) {
            const criticalRow = attack.querySelector('tr.damage th[data-damage-type="critical"]');
            if (criticalRow) {
                const critDamageValue = criticalRow.querySelector('a.fake-inline-roll');
                if (critDamageValue) {
                    criticalDamage = parseInt(critDamageValue.textContent.trim(), 10);
                    if (!isNaN(criticalDamage)) {
                        fallbackTotalDamage += criticalDamage;
                    }
                }
            }
        }

        if (includeNormalDamage || includeCriticalDamage) {
            const damageTypeElements = attack.querySelectorAll('.damage-type');
            damageTypeElements.forEach(element => {
                const damageType = element.getAttribute('data-tooltip')?.trim() || element.textContent.trim();
                if (damageType) {
                    fallbackDamageTypes.add(damageType.toLowerCase());
                }
            });
        }

        let ammoItem = null;
        const ammoContainer = attack.querySelector('.ammo.group-container');
        if (ammoContainer) {
            const ammoId = ammoContainer.getAttribute('data-ammo-id');
            const cardElem = attack.closest('.pf1.chat-card.item-card');
            if (ammoId && cardElem) {
                const tokenUuid = cardElem.getAttribute('data-token-uuid');
                const actorId = cardElem.getAttribute('data-actor-id');
                if (tokenUuid) {
                    try {
                        const token = await fromUuid(tokenUuid);
                        if (token && token.actor) {
                            ammoItem = token.actor.items.get(ammoId);
                        }
                    } catch (e) {
                        console.warn('Could not fetch token or ammo item from tokenUuid', tokenUuid, e);
                        ammoItem = null;
                    }
                } else if (actorId) {
                    try {
                        const actor = game.actors.get(actorId);
                        if (actor) {
                            ammoItem = actor.items.get(ammoId);
                        }
                    } catch (e) {
                        console.warn('Could not fetch actor or ammo item from actorId', actorId, e);
                        ammoItem = null;
                    }
                }
            }
        }

        criticalHits.push({
            index: attackIndex,
            normalDamage: normalDamage,
            criticalDamage: criticalDamage,
            attackName: `Attack #${parseInt(attackIndex) + 1}`,
            ammoItem: ammoItem
        });
    }

    const instancesByKey = new Map();
    const rolls =
        (messageObject?.systemRolls && Object.keys(messageObject.systemRolls).length)
            ? messageObject.systemRolls
            : messageObject?.rolls;
    const attacks = rolls?.attacks;

    const addInstancePart = (total, types) => {
        if (!Number.isFinite(total) || total === 0) return;
        const resolvedTypes = (Array.isArray(types) && types.length) ? types : ["untyped"];
        const key = [...resolvedTypes].sort().join("|");
        instancesByKey.set(key, (instancesByKey.get(key) || 0) + total);
    };

    if (Array.isArray(attacks)) {
        for (const selection of selectedAttacks) {
            if (!Number.isInteger(selection.attackIndex)) continue;
            const attack = attacks[selection.attackIndex];
            if (!attack) continue;

            if (selection.includeNormalDamage) {
                const baseParts = Array.isArray(attack.damage) ? attack.damage : [];
                for (const p of baseParts) {
                    addInstancePart(p?.total, p?.options?.damageType);
                }
            }

            if (selection.includeCriticalDamage) {
                const critParts = Array.isArray(attack.critDamage) ? attack.critDamage : [];
                for (const p of critParts) {
                    addInstancePart(p?.total, p?.options?.damageType);
                }
            }
        }
    }

    let instances = Array.from(instancesByKey.entries()).map(([key, value]) => ({
        types: key.split("|"),
        value,
        formula: String(value)
    }));

    if (!instances.length) {
        const types = fallbackDamageTypes.size ? Array.from(fallbackDamageTypes) : ["untyped"];
        if (Number.isFinite(fallbackTotalDamage) && fallbackTotalDamage > 0) {
            const formula = String(fallbackTotalDamage);
            instances = [{ types, value: fallbackTotalDamage, formula }];
        }
    }

    const totalDamage = instances.reduce((sum, inst) => sum + (Number(inst?.value) || 0), 0);
    const damageTypes = new Set(instances.flatMap((inst) => inst.types || []));

    if (!Number.isFinite(totalDamage) || totalDamage <= 0) {
        ui.notifications.warn("Clustered Shots found no damage to apply.");
        return;
    }

    if (clickedButton && clickedButton.dataset) {
        clickedButton.dataset.value = totalDamage.toString();
    }

    let action = null;
    const tokenUuid = card.getAttribute('data-token-uuid');
    const itemId = card.getAttribute('data-item-id');
    const actionId = card.getAttribute('data-action-id');
    if (tokenUuid && itemId && actionId && typeof fromUuid === 'function') {
        try {
            const token = await fromUuid(tokenUuid);
            const item = token?.actor?.items?.get(itemId);
            if (item && item.actions && typeof item.actions.get === 'function') {
                action = item.actions.get(actionId) ?? null;
            }
        } catch (e) {
            console.warn('Could not resolve action for clustered shots', tokenUuid, itemId, actionId, e);
        }
    }

    const applyDamage = pf1?.documents?.actor?.ActorPF?.applyDamage;
    if (typeof applyDamage !== 'function') {
        ui.notifications.error("ApplyDamage is not available.");
        return;
    }

    const targets = canvas.tokens.controlled;
    if (!targets?.length && !game.user.character) {
        ui.notifications.warn("Please select at least one token to apply damage to.");
        return;
    }

    await applyDamage(totalDamage, {
        targets,
        instances,
        action,
        message: messageObject,
        element: clickedButton,
        event,
        ratio: 1,
        dialog: false,
        asNonlethal: false,
        flags: {
            [MODULE.ID]: {
                clusteredShots: {
                    totalDamage: totalDamage,
                    damageTypes: Array.from(damageTypes),
                    messageId: messageId,
                    message: messageObject,
                    buttonElement: clickedButton,
                    buttonType: "Clustered Shots",
                    criticalHits: criticalHits
                }
            }
        }
    });

    ui.notifications.info(`Applied ${totalDamage} points of Clustered Shots damage to ${targets?.length || 1} targets.`);
}

function styleClusteredShotButton(button) {
    button.style.display = "flex";
    button.style.alignItems = "center";
    button.style.justifyContent = "center";
    button.style.visibility = "visible";
    button.style.cursor = "pointer";
    button.style.fontSize = "1em";
    button.style.margin = "8px 0";
    button.style.padding = "5px 10px";
    button.style.width = "100%";
    button.style.backgroundColor = "#4b4a44";
    button.style.color = "white";
    button.style.borderRadius = "3px";
    button.style.textAlign = "center";
    button.style.fontWeight = "bold";
    button.style.border = "1px solid #777";
    button.style.boxShadow = "0 1px 3px rgba(0,0,0,0.2)";
    
    button.addEventListener('mouseenter', () => {
        button.style.backgroundColor = "#5e5d57";
        button.style.boxShadow = "0 2px 5px rgba(0,0,0,0.3)";
    });
    
    button.style.transition = "all 150ms";
    
    button.addEventListener('mouseleave', () => {
        button.style.backgroundColor = "#4b4a44";
        button.style.boxShadow = "0 1px 3px rgba(0,0,0,0.2)";
    });
}
