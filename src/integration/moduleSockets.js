import { MODULE } from '../common/module.js';
import { applyBuffToTargets } from '../features/automation/buffs/buffs.js';

export let socket;

export function initializeSockets() {
    if (!game.modules.get("socketlib")?.active) {
        if (game.settings.get(MODULE.ID, "massiveDamage")) {
            ui.notifications.warn("SocketLib is required for the Massive Damage rule to work properly. Please install and activate the socketlib module.");
        }
        return;
    }
    
    socket = socketlib.registerModule(MODULE.ID);
    
    socket.register("rollMassiveDamageSave", rollMassiveDamageSave);
    socket.register("applyImmobilize", applyImmobilize);
    socket.register("sendNotification", sendNotification);
    socket.register("promptHTKChoice", promptHTKChoice);
    socket.register("handleFlatFootedRemoval", handleFlatFootedRemoval);
    socket.register("applyBuffToTargetsSocket", applyBuffToTargetsSocket);
}

async function rollMassiveDamageSave(actorId, damageAmount, threshold) {
    const actor = game.actors.get(actorId);
    if (!actor) return { name: "Unknown", result: "No actor found" };

    const roll = await actor.rollSavingThrow("fort", { dc: 15 });
    
    let total = 0;
    if (roll.rolls && roll.rolls.length > 0) {
        total = roll.rolls[0].total || 0;
    }
    
    const success = total >= 15;
    
    if (!success) {
        await actor.setCondition('dead', {overlay: true});
    }
    
    let content = `<p><strong>${actor.name}</strong> took massive damage (${damageAmount} damage, threshold: ${threshold})!</p>`;
    content += `<p>Fortitude save result: <strong>${total}</strong> - ${success ? "Success! " + actor.name + " survives the massive damage." : "Failure! " + actor.name + " dies from massive damage."}</p>`;
    
    ChatMessage.create({
        user: game.user.id,
        speaker: ChatMessage.getSpeaker({actor}),
        content: content,
        type: CONST.CHAT_MESSAGE_TYPES.OTHER
    });
    
    return { 
        name: actor.name, 
        total: total, 
        success: success,
        damage: damageAmount,
        threshold: threshold
    };
}

export function checkMassiveDamage(damage, maxHP, token) {
    const massiveDamageEnabled = game.settings.get(MODULE.ID, "massiveDamage");
    if (!massiveDamageEnabled) return;
    
    const damageThreshold = Math.max(Math.floor(maxHP / 2), 50);
    
    if (damage >= damageThreshold) {
        
        if (!game.modules.get("socketlib")?.active) {
            ui.notifications.warn(`${token.name} has taken massive damage (${damage} damage)! SocketLib is not available, cannot roll save remotely.`);
            if (game.user.isGM) {
                const actor = token.actor;
                ChatMessage.create({
                    content: `<p><strong>${actor.name}</strong> took massive damage (${damage}). Roll a DC 15 Fortitude save or die.</p>`,
                    speaker: ChatMessage.getSpeaker({token})
                });
            }
            return;
        }
        
        if (typeof socketlib === 'undefined' || !socket) {
            ui.notifications.warn(`${token.name} has taken massive damage (${damage} damage)! Socket not initialized yet.`);
            if (typeof socketlib !== 'undefined' && !socket) {
                initializeSockets();
            }
            return;
        }
        
        const actorId = token.actor.id;
        
        const nonGmOwners = Object.entries(token.actor.ownership)
            .filter(([userId, level]) => level >= CONST.DOCUMENT_OWNERSHIP_LEVELS.OWNER && 
                    userId !== "default" && 
                    game.users.get(userId)?.active && 
                    !game.users.get(userId)?.isGM);
        
        if (nonGmOwners.length > 0) {
            const targetUserId = nonGmOwners[0][0];
            socket.executeAsUser("rollMassiveDamageSave", targetUserId, actorId, damage, damageThreshold)
                .catch(error => {
                    console.error("Error rolling massive damage save:", error);
                    const activeGm = game.users.find(u => u.active && u.isGM);
                    if (activeGm) {
                        socket.executeAsUser("rollMassiveDamageSave", activeGm.id, actorId, damage, damageThreshold);
                    }
                });
        } else {
            const activeGm = game.users.find(u => u.active && u.isGM);
            if (activeGm) {
                socket.executeAsUser("rollMassiveDamageSave", activeGm.id, actorId, damage, damageThreshold);
            } else {
                ui.notifications.warn(`${token.name} has taken massive damage (${damage} damage), but no user was found to roll the save.`);
            }
        }
    }
} 
let immobileConditionIds = new Set();

export function initializeConditionIds() {
    const immobileConditions = ["anchored", "cowering", "dazed", "dying", "grappled", "helpless", "paralyzed", "petrified", "pinned"];

    pf1.registry.conditions.forEach(condition => {
        if (immobileConditions.includes(condition._id)) {
            immobileConditionIds.add(condition._id);
        }
    });
}

function hasImmobileCondition(token) {
    return token.actor.statuses?.some(status => immobileConditionIds.has(status)) ?? false;
}

async function applyImmobilize(tokenId, limit) {
    if (!game.settings.get(MODULE.ID, 'restrictMovement')) {
        return true;
    }
    const token = canvas.tokens.get(tokenId);
    if (!token) return;

    const currentLimit = token.document.getFlag(MODULE.ID, 'immobilized');

    if (currentLimit !== undefined) {
        await token.document.unsetFlag(MODULE.ID, 'immobilized');
        sendNotificationToOwners(token, "info", game.i18n.localize('NAS.conditions.sockets.MovementRestrictionRemoved'));
    } else {
        await token.document.setFlag(MODULE.ID, 'immobilized', limit);  
        sendNotificationToOwners(token, "info", game.i18n.localize('NAS.conditions.sockets.MovementRestrictionApplied'));
    }
}

async function promptHTKChoice(actorId) {
    const actor = game.actors.get(actorId);
    if (!actor) return;

    const content = `<p>${game.i18n.format('NAS.conditions.sockets.HardToKillPrompt', { name: actor.name })}</p>`;
    const options = [game.i18n.localize('NAS.conditions.sockets.ContinueFighting'), game.i18n.localize('NAS.conditions.sockets.FallUnconscious')];
    const choice = await new Promise(resolve => {
        new Dialog({
            title: game.i18n.localize('NAS.conditions.sockets.HardToKillChoice'),
            content,
            buttons: {
                fight: {
                    label: options[0],
                    callback: () => resolve("fight")
                },
                unconscious: {
                    label: options[1],
                    callback: () => resolve("unconscious")
                }
            },
            default: "fight"
        }).render(true);
    });

    return choice;
}

let updatingToken = false;

Hooks.on('preUpdateToken', (tokenDocument, updateData, options, userId) => {
    if (updatingToken) return true; 

    const restrictSetting = game.settings.get(MODULE.ID, 'restrictMovement');
    const limit = tokenDocument.getFlag(MODULE.ID, 'immobilized');
    if (limit !== undefined || hasImmobileCondition(tokenDocument)) {
        const currentX = tokenDocument.x;
        const currentY = tokenDocument.y;
        const newX = updateData.x !== undefined ? updateData.x : currentX;
        const newY = updateData.y !== undefined ? updateData.y : currentY;

        const deltaX = Math.abs(newX - currentX);
        const deltaY = Math.abs(newY - currentY);
        const gridSize = canvas.grid.size;
        const maxMove = gridSize * limit;

        if (deltaX > maxMove || deltaY > maxMove) {
            if (restrictSetting === "all" || (restrictSetting === "players" && !game.user.isGM)) {
                if (game.user.id === userId) {
                    const limitFeet = limit * 5;
                    socket.executeAsUser("sendNotification", userId, "warn", game.i18n.format('NAS.conditions.sockets.MaxMoveWarning', { limitFeet: limitFeet }));
                }
                return false;
            }
        }
    }

    if (game.settings.get(MODULE.ID, 'blindMovementCheck')) {
        const token = canvas.tokens.get(tokenDocument.id);
        if (!token) return true;
    
        const hasBlindCondition = token.actor.statuses?.some(status => status === "blind") ?? false;
    
        if (hasBlindCondition && (updateData.x !== undefined || updateData.y !== undefined)) {
            if (options) options.cancelled = true;
            const currentX = tokenDocument.x;
            const currentY = tokenDocument.y;
            const newX = updateData.x ?? currentX;
            const newY = updateData.y ?? currentY;

            const gridSize = canvas.grid?.size || 1;
            const gridDistance = canvas.dimensions?.distance ?? canvas.scene?.grid?.distance ?? 5;
            const distanceMoved = (Math.hypot(newX - currentX, newY - currentY) / gridSize) * gridDistance;

            if (distanceMoved === 0) return true;

            const speedData = token.actor.system?.attributes?.speed ?? {};
            const availableSpeeds = Object.entries(speedData).filter(([_, data]) => typeof data?.total === "number" && data.total > 0);

            if (!availableSpeeds.length) return true;

            const moveToken = async () => {
                updatingToken = true;
                await token.document.update(updateData);
                updatingToken = false;
            };

            const promptAcrobatics = () => {
                new Dialog({
                    title: game.i18n.localize('NAS.conditions.sockets.BlindMovementCheckTitle'),
                    content: `<p>${game.i18n.format('NAS.conditions.sockets.BlindMovementCheckPrompt', { name: token.name })}</p>`,
                    buttons: {
                        roll: {
                            label: game.i18n.localize('NAS.conditions.sockets.RollAcrobatics'),
                            callback: async () => {
                                const roll = await token.actor.rollSkill("acr");
                                if (roll.rolls[0].total >= 10) {
                                    await moveToken();
                                } else {
                                    token.actor.setCondition("prone", true);
                                    return false;
                                }
                            }
                        },
                        cancel: {
                            label: game.i18n.localize('NAS.common.buttons.cancel'),
                            callback: () => {
                            }
                        }
                    },
                    default: "roll"
                }).render(true);
            };

            const handleSpeedSelection = async (speedTotal) => {
                if (distanceMoved <= speedTotal / 2) {
                    await moveToken();
                    return;
                }
                promptAcrobatics();
            };

            const defaultSpeedTotal = availableSpeeds[0][1].total;
            const speedOptions = availableSpeeds.map(([speedType, data], index) => {
                const label = `${speedType.charAt(0).toUpperCase()}${speedType.slice(1)} (${data.total})`;
                const checked = index === 0 ? "checked" : "";
                return `<label><input type="radio" name="blind-speed-type" value="${speedType}" data-total="${data.total}" ${checked}/> ${label}</label>`;
            }).join("<br/>");

            const dialogContent = `
                <p>${game.i18n.localize('NAS.conditions.sockets.BlindMovementSpeedPrompt')}</p>
                <p class="blind-move-info">${game.i18n.format('NAS.conditions.sockets.BlindMovementSpeedInfo', { halfSpeed: defaultSpeedTotal / 2 })}</p>
                <form class="blind-move-speed-form">
                    ${speedOptions}
                </form>
            `;

            new Dialog({
                title: game.i18n.localize('NAS.conditions.sockets.BlindMovementSpeedTitle'),
                content: dialogContent,
                buttons: {
                    confirm: {
                        label: game.i18n.localize('NAS.common.labels.confirm') ?? "Confirm",
                        callback: async (html) => {
                            const selected = html.find('input[name="blind-speed-type"]:checked');
                            const speedTotal = Number(selected.data("total"));
                            await handleSpeedSelection(speedTotal);
                        }
                    },
                    cancel: {
                        label: game.i18n.localize('NAS.common.buttons.cancel'),
                        callback: () => {
                        }
                    }
                },
                default: "confirm",
                render: (html) => {
                    html.find('input[name="blind-speed-type"]').on('change', ev => {
                        const speedTotal = Number(ev.currentTarget.dataset.total);
                        html.find('.blind-move-info').text(
                            game.i18n.format('NAS.conditions.sockets.BlindMovementSpeedInfo', { halfSpeed: speedTotal / 2 })
                        );
                    });
                }
            }).render(true);

            return false; 
        };
    };

    return true;
});

function sendNotification(type, message) {
    ui.notifications[type](message);
}

function sendNotificationToOwners(token, type, message) {
    if (token.actor.hasPlayerOwner) {
        const owners = game.users.filter(user => user.id == token.actor.activeOwner.id);
        for (let user of owners) {
            if (token.actor.isOwner) {
                socket.executeAsUser("sendNotification", user.id, type, message);
            }
        }
    } else {
        if (game.user.isGM) {
            socket.executeAsUser("sendNotification", game.user.id, type, message);
        }
    }
}

Hooks.on('pf1ToggleActorCondition', async (actor, condition, enabled) => {
    if (immobileConditionIds.has(condition)) {
        const tokens = actor.getActiveTokens();
        for (const token of tokens) {
            await socket.executeAsGM("applyImmobilize", token.id, enabled ? 0 : null);
        }
    }
});

async function handleFlatFootedRemoval(tokenId, round, turn) {
  const token = canvas.tokens.get(tokenId);
  if (!token || !token.actor) return;
  
  const actor = token.actor;
  
  if (!actor.statuses.has("flatFooted")) return;
  
  const combat = game.combat;
  if (!combat) return;
  
  const ffTracker = combat.getFlag(MODULE.ID, "flatFootedTracker") || {};
  const trackerData = ffTracker[tokenId];
  
  if (trackerData) {
    const targetRemovalRound = trackerData.targetRemovalRound || 1;
    
    if (round >= targetRemovalRound) {
      try {
        const isBeingProcessed = combat.getFlag(MODULE.ID, "processingFlatFooted") || {};
        
        if (isBeingProcessed[tokenId]) return;
        
        isBeingProcessed[tokenId] = true;
        await combat.setFlag(MODULE.ID, "processingFlatFooted", isBeingProcessed);
        
        if (actor.statuses.has("flatFooted")) {
          await actor.setCondition("flatFooted", false);
        }
        
        trackerData.removalInfo = {
          removedOnRound: round,
          removedOnTurn: turn
        };
        trackerData.wasFlatFooted = false;
        
        ffTracker[tokenId] = trackerData;
        await combat.setFlag(MODULE.ID, "flatFootedTracker", ffTracker);
        
        isBeingProcessed[tokenId] = false;
        await combat.setFlag(MODULE.ID, "processingFlatFooted", isBeingProcessed);
      } catch (error) {
        console.error("Error in handleFlatFootedRemoval:", error);
        const isBeingProcessed = combat.getFlag(MODULE.ID, "processingFlatFooted") || {};
        isBeingProcessed[tokenId] = false;
        await combat.setFlag(MODULE.ID, "processingFlatFooted", isBeingProcessed);
      }
    }
  }
}

async function applyBuffToTargetsSocket(buffData, targetIds, duration, casterLevel, options = {}) {
    let buffDoc = null;
    if (buffData.pack) {
        const pack = game.packs.get(buffData.pack);
        if (!pack) return;
        buffDoc = await pack.getDocument(buffData.id);
    } else {
        buffDoc = game.items.get(buffData.id);
    }
    if (!buffDoc) return;
    const targets = targetIds.map(id => canvas.tokens.get(id)).filter(Boolean);
    await applyBuffToTargets({ ...buffData, document: buffDoc }, targets, duration, casterLevel, options);
}







