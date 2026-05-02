import { MODULE } from '../common/module.js';
import { chatMessageStyle } from '../common/foundryCompat.js';
import { applyBuffToTargets } from '../features/automation/buffs/buffs.js';
import {
    applyMirrorImageStateSocket,
    undoMirrorImageOperationSocket
} from '../features/automation/buffs/mirrorImage.js';

export let socket;

function localizeSockets(path) {
    return game.i18n.localize(`NAS.integration.sockets.${path}`);
}

function formatSockets(path, data = {}) {
    return game.i18n.format(`NAS.integration.sockets.${path}`, data);
}

function isConditionAutomationEnabled() {
    const key = `${MODULE.ID}.automateConditions`;
    if (!globalThis.game?.settings?.settings?.has?.(key)) return true;
    return game.settings.get(MODULE.ID, "automateConditions");
}

async function applyPlayerDamageWithAutomation(actor, value, options = {}) {
    if (actor) {
        const { applyNasHeadlessDamage } = await import('../features/automation/damage/systemApplyDamage.js');
        const applied = await applyNasHeadlessDamage(value, { ...options, targets: [actor], forceDialog: false });
        if (applied?.handled) return applied.result;
    }

    const Ctor = pf1?.documents?.actor?.ActorPF;
    if (!actor || Ctor == null || typeof Ctor.applyDamage !== "function") {
        return actor?.applyDamage?.(value, options);
    }
    return Ctor.applyDamage(value, { ...options, targets: [actor], forceDialog: false });
}

export function initializeSockets() {
    if (!game.modules.get("socketlib")?.active) {
        if (game.settings.get(MODULE.ID, "massiveDamage")) {
            ui.notifications.warn(localizeSockets("warnings.socketLibRequired"));
        }
        return;
    }
    
    socket = socketlib.registerModule(MODULE.ID);
    
    socket.register("rollMassiveDamageSave", rollMassiveDamageSave);
    socket.register("applyImmobilize", applyImmobilize);
    socket.register("sendNotification", sendNotification);
    socket.register("promptHTKChoice", promptHTKChoice);
    socket.register("promptHTKUseFeature", promptHTKUseFeature);
    socket.register("useHtkItem", useHtkItem);
    socket.register("postHtkItemCard", postHtkItemCard);
    socket.register("applyGrappleToTarget", applyGrappleToTarget);
    socket.register("handleFlatFootedRemoval", handleFlatFootedRemoval);
    socket.register("applyBuffToTargetsSocket", applyBuffToTargetsSocket);
    socket.register("applyReactiveDamageToActorSocket", applyReactiveDamageToActorSocket);
    socket.register("toggleReactiveConditionSocket", toggleReactiveConditionSocket);
    socket.register("toggleReactiveBuffSocket", toggleReactiveBuffSocket);
    socket.register("setMirrorImageBuffStateSocket", applyMirrorImageStateSocket);
    socket.register("undoMirrorImageOperationSocket", undoMirrorImageOperationSocket);
}

function getHtkWhisperTargets() {
    const gmIds = (game.users?.filter?.((u) => u?.isGM)?.map?.((u) => u.id) ?? []).filter(Boolean);
    const base = game.user?.isGM ? gmIds : [game.user.id, ...gmIds];
    return Array.from(new Set(base.filter(Boolean)));
}

async function resolveTokenContext(tokenUuid) {
    if (!tokenUuid || typeof tokenUuid !== "string" || typeof fromUuid !== "function") {
        return { token: null, tokenDocument: null, tokenActor: null };
    }

    let tokenDoc = null;
    try {
        tokenDoc = await fromUuid(tokenUuid);
    } catch (_err) {
        tokenDoc = null;
    }

    const token =
        tokenDoc?.object ??
        (tokenDoc?.id ? canvas?.tokens?.get?.(tokenDoc.id) : null) ??
        null;
    const tokenDocument = tokenDoc ?? token?.document ?? null;
    const tokenActor = token?.actor ?? tokenDoc?.actor ?? null;

    return { token, tokenDocument, tokenActor };
}

async function resolveTokenOwnedItem(itemUuid, tokenActor) {
    if (!itemUuid || typeof itemUuid !== "string" || typeof fromUuid !== "function") return null;
    let item = null;
    try {
        item = await fromUuid(itemUuid);
    } catch (_err) {
        item = null;
    }
    if (!item) return null;
    if (!tokenActor) return item;

    if (item.actor && item.actor.uuid !== tokenActor.uuid) {
        const tokenItem = tokenActor.items?.get?.(item.id) ?? tokenActor.items?.find?.((i) => i.id === item.id) ?? null;
        if (tokenItem) return tokenItem;
    }

    return item;
}

async function postHtkItemCard(actorId, itemUuid, tokenUuid = null) {
    if (!actorId || !itemUuid) return false;
    const whisper = getHtkWhisperTargets();

    const { tokenDocument, tokenActor } = await resolveTokenContext(tokenUuid);
    const baseActor = game.actors.get(actorId) ?? null;
    const actor = tokenActor ?? baseActor;
    if (!actor) return false;

    const item = await resolveTokenOwnedItem(itemUuid, tokenActor);
    if (!item) return false;

    if (item.actor?.id && item.actor.id !== actor.id) return false;
    if (item.actor?.uuid && actor.uuid && item.actor.uuid !== actor.uuid) return false;

    try {
        await item.displayCard?.({ whisper, rollMode: "gmroll" }, { token: tokenDocument });
        return true;
    } catch (_err) {
        return false;
    }
}

async function applyGrappleToTarget(targetTokenUuid, attackerActorUuid, attackTotal, cmdValue) {
    if (!targetTokenUuid || typeof fromUuid !== "function") return false;
    let tokenDoc = null;
    try {
        tokenDoc = await fromUuid(targetTokenUuid);
    } catch (_err) {
        tokenDoc = null;
    }
    const token = tokenDoc?.object ?? (tokenDoc?.id ? canvas?.tokens?.get?.(tokenDoc.id) : null) ?? null;
    const targetActor = token?.actor ?? tokenDoc?.actor ?? null;
    if (!targetActor) return false;

    await targetActor.setCondition?.("grappled", true);

    await targetActor.setFlag(MODULE.ID, "grappleContext", {
        attacker: attackerActorUuid ?? null,
        attackTotal,
        cmd: cmdValue,
        timestamp: Date.now(),
    });

    return true;
}

export function getPreferredOwnerUserId(actor) {
    if (!actor) return null;
    try {
        const ownerUsers = game.users.filter(u => actor.testUserPermission(u, 'OWNER'));
        const nonGMOwners = ownerUsers.filter(u => !u.isGM && u.active);
        if (nonGMOwners.length > 0) return nonGMOwners[0].id;
        const activeOwners = ownerUsers.filter(u => u.active);
        if (activeOwners.length > 0) return activeOwners[0].id;
    } catch (_err) {
    }
    const activeGm = game.users.find(u => u.active && u.isGM);
    return activeGm?.id ?? null;
}

async function resolveActorByUuid(actorUuid) {
    if (!actorUuid) return null;
    try {
        if (typeof fromUuid === "function") {
            const actorDoc = await fromUuid(actorUuid);
            if (actorDoc) return actorDoc;
        }
    } catch (_err) {
    }
    if (String(actorUuid).includes(".Token.")) {
        return null;
    }
    const m = String(actorUuid).match(/^Actor\.([^.]+)$/i);
    if (m) {
        return game.actors?.get?.(m[1]) ?? null;
    }
    return null;
}

function findActorBuffBySource(actor, buffUuid) {
    if (!actor || !buffUuid) return null;
    return actor.items?.find?.((item) => {
        if (item.type !== "buff") return false;
        const source = item.flags?.[MODULE.ID]?.sourceId || item.flags?.core?.sourceId || item._stats?.compendiumSource;
        return source === buffUuid;
    }) ?? null;
}

async function resolveBuffDocumentByUuid(buffUuid) {
    if (!buffUuid) return null;
    try {
        const doc = await fromUuid(buffUuid);
        return doc?.type === "buff" ? doc : null;
    } catch (_err) {
        return null;
    }
}

async function setReactiveBuffState(actorUuid, buffUuid, enabled) {
    const actor = await resolveActorByUuid(actorUuid);
    if (!actor) return false;

    const existingBuff = findActorBuffBySource(actor, buffUuid);
    if (existingBuff) {
        await existingBuff.update({ "system.active": enabled === true });
        return true;
    }

    if (enabled !== true) return false;

    const buffDoc = await resolveBuffDocumentByUuid(buffUuid);
    if (!buffDoc) return false;

    let buffData;
    if (buffDoc.pack && typeof Item?.implementation?.fromCompendium === "function") {
        buffData = await Item.implementation.fromCompendium(buffDoc);
    } else if (buffDoc.pack && typeof Item?.fromCompendium === "function") {
        buffData = await Item.fromCompendium(buffDoc);
    } else {
        buffData = buffDoc.toObject();
    }

    buffData.flags = buffData.flags || {};
    buffData.flags[MODULE.ID] = buffData.flags[MODULE.ID] || {};
    buffData.flags[MODULE.ID].sourceId = buffUuid;
    buffData.system = buffData.system || {};
    buffData.system.active = true;

    await actor.createEmbeddedDocuments("Item", [buffData]);
    return true;
}

export async function applyReactiveDamageToActor(actor, value, options = {}) {
    if (!actor || !Number.isFinite(Number(value))) return;
    const amount = Number(value);
    if (amount === 0) return;

    const canModify = game.user?.isGM || actor?.isOwner;
    if (!canModify) {
        if (!socket) return;
        await socket.executeAsGM("applyReactiveDamageToActorSocket", actor.uuid, amount, options);
        return;
    }
    return applyPlayerDamageWithAutomation(actor, amount, options);
}

export async function toggleReactiveCondition(actor, conditionId, enabled) {
    if (!actor || !conditionId) return;
    const canModify = game.user?.isGM || actor?.isOwner;
    if (!canModify) {
        if (!socket) return;
        await socket.executeAsGM("toggleReactiveConditionSocket", actor.uuid, String(conditionId), enabled === true);
        return;
    }
    await actor.setCondition?.(String(conditionId), enabled === true);
}

export async function toggleReactiveBuff(actor, buffUuid, enabled) {
    if (!actor || !buffUuid) return;
    const canModify = game.user?.isGM || actor?.isOwner;
    if (!canModify) {
        if (!socket) return;
        await socket.executeAsGM("toggleReactiveBuffSocket", actor.uuid, String(buffUuid), enabled === true);
        return;
    }
    await setReactiveBuffState(actor.uuid, String(buffUuid), enabled === true);
}

async function rollMassiveDamageSave(actorId, damageAmount, threshold) {
    const actor = game.actors.get(actorId);
    if (!actor) {
        return {
            name: localizeSockets("results.unknownName"),
            result: localizeSockets("results.noActorFound")
        };
    }

    const roll = await actor.rollSavingThrow("fort", { dc: 15 });
    
    let total = 0;
    if (roll.rolls && roll.rolls.length > 0) {
        total = roll.rolls[0].total || 0;
    }
    
    console.log("Massive damage save roll result:", total);
    const success = total >= 15;
    
    if (!success) {
        await actor.setCondition('dead', {overlay: true});
    }
    
    let content = formatSockets("chat.massiveDamageSummary", {
        name: actor.name,
        damage: damageAmount,
        threshold
    });
    content += success
        ? formatSockets("chat.fortitudeSuccess", { name: actor.name, total })
        : formatSockets("chat.fortitudeFailure", { name: actor.name, total });
    
        ChatMessage.create({
            ...chatMessageStyle("OTHER"),
            user: game.user.id,
            speaker: ChatMessage.getSpeaker({actor}),
            content: content
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
        console.log(`Massive damage threshold met: ${damage} damage >= ${damageThreshold} threshold for ${token.name}`);
        
        if (!game.modules.get("socketlib")?.active) {
            ui.notifications.warn(formatSockets("warnings.massiveDamageNoSocketLib", {
                name: token.name,
                damage
            }));
            if (game.user.isGM) {
                const actor = token.actor;
                ChatMessage.create({
                    content: formatSockets("chat.gmFallbackPrompt", {
                        name: actor.name,
                        damage
                    }),
                    speaker: ChatMessage.getSpeaker({token})
                });
            }
            return;
        }
        
        if (typeof socketlib === 'undefined' || !socket) {
            ui.notifications.warn(formatSockets("warnings.massiveDamageSocketUninitialized", {
                name: token.name,
                damage
            }));
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
                .then(result => {
                    console.log("Massive damage save result:", result);
                })
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
                ui.notifications.warn(formatSockets("warnings.massiveDamageNoUser", {
                    name: token.name,
                    damage
                }));
            }
        }
    }
} 
let immobileConditionIds = new Set();

export function initializeConditionIds() {
    const immobileConditions = ["anchored", "cowering", "dazed", "grappled", "helpless", "paralyzed", "petrified", "pinned"];

    pf1.registry.conditions.forEach(condition => {
        if (immobileConditions.includes(condition._id)) {
            immobileConditionIds.add(condition._id);
        }
    });
}

function hasImmobileCondition(token) {
    return token.actor.statuses?.some(status => immobileConditionIds.has(status)) ?? false;
}

function getActorSpeedFromSystemData(actor, movementType = "walk") {
    const speedData = actor?.system?.attributes?.speed ?? {};
    const normalized = String(movementType ?? "walk").toLowerCase();
    const mappedKey = normalized === "walk" ? "land" : normalized;
    const speedValue = Number(speedData?.[mappedKey]?.total ?? speedData?.[mappedKey]);
    return Number.isFinite(speedValue) && speedValue > 0 ? speedValue : 0;
}

function getActorMovementSpeed(actor, movementType = "walk") {
    const normalized = String(movementType ?? "walk").toLowerCase();
    const movementInfo = actor?.getMovement?.(normalized);
    const movementSpeed = Number(movementInfo?.speed);
    if (Number.isFinite(movementSpeed) && movementSpeed > 0) return movementSpeed;
    return getActorSpeedFromSystemData(actor, normalized);
}

function getLegacySpeedOptions(actor) {
    const speedData = actor?.system?.attributes?.speed ?? {};
    return Object.entries(speedData)
        .map(([speedType, data]) => ({
            speedType,
            total: Number(data?.total ?? data)
        }))
        .filter(({ total }) => Number.isFinite(total) && total > 0);
}

async function applyImmobilize(tokenId, limit) {
    if (!isConditionAutomationEnabled()) {
        return true;
    }
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

async function promptHTKUseFeature(actorId, titleKey, bodyKey, bodyData = {}) {
    const actor = game.actors.get(actorId);
    if (!actor) return false;

    const title = game.i18n.localize(titleKey);
    const content = `<p>${game.i18n.format(bodyKey, { name: actor.name, ...bodyData })}</p>`;

    const yes = game.i18n.localize('NAS.conditions.sockets.UseFeatureYes');
    const no = game.i18n.localize('NAS.conditions.sockets.UseFeatureNo');

    const choice = await new Promise(resolve => {
        new Dialog({
            title,
            content,
            buttons: {
                yes: { label: yes, callback: () => resolve(true) },
                no: { label: no, callback: () => resolve(false) }
            },
            default: "yes"
        }).render(true);
    });
    return Boolean(choice);
}

async function useHtkItem(actorId, itemUuid, tokenUuid = null) {
    if (!actorId || !itemUuid) return false;
    if (typeof fromUuid !== "function") return false;

    const whisper = getHtkWhisperTargets();
    const { tokenDocument, tokenActor } = await resolveTokenContext(tokenUuid);

    const baseActor = game.actors.get(actorId) ?? null;
    const actor = tokenActor ?? baseActor;
    if (!actor) return false;

    let item = await resolveTokenOwnedItem(itemUuid, tokenActor);
    if (!item) return false;

    if (item.actor?.id && item.actor.id !== actor.id) return false;
    if (item.actor?.uuid && actor.uuid && item.actor.uuid !== actor.uuid) return false;

    try {
        await item.use?.({ skipDialog: true, chatMessage: false, token: tokenDocument });

        await item.displayCard?.({ whisper, rollMode: "gmroll" }, { token: tokenDocument });
        return true;
    } catch (_err) {
        return false;
    }
}

let updatingToken = false;

const BLIND_MOVEMENT_ACROBATICS_DC = 10;

Hooks.on('preUpdateToken', (tokenDocument, updateData, options, userId) => {
    if (!isConditionAutomationEnabled()) return true;
    if (updatingToken) return true; 

    if (!globalThis.game?.settings?.settings?.has?.(`${MODULE.ID}.restrictMovement`)) {
        return true;
    }
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
                                const roll = await token.actor.rollSkill("acr", {
                                    dc: BLIND_MOVEMENT_ACROBATICS_DC,
                                    reason: "blindMovement"
                                });
                                const total =
                                    roll?.rolls?.[0]?.total
                                    ?? roll?.total
                                    ?? 0;
                                if (total >= BLIND_MOVEMENT_ACROBATICS_DC) {
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

            const currentMovementType = tokenDocument.movementAction;
            const hasTokenMovementMode = typeof currentMovementType === "string" && currentMovementType.length > 0;
            if (hasTokenMovementMode) {
                const speedTotal = getActorMovementSpeed(token.actor, currentMovementType);
                if (speedTotal > 0) {
                    void handleSpeedSelection(speedTotal);
                    return false; 
                }
            }

            const availableSpeeds = getLegacySpeedOptions(token.actor);
            if (!availableSpeeds.length) return true;

            const defaultSpeedTotal = availableSpeeds[0].total;
            const speedOptions = availableSpeeds.map(({ speedType, total }, index) => {
                const label = `${speedType.charAt(0).toUpperCase()}${speedType.slice(1)} (${total})`;
                const checked = index === 0 ? "checked" : "";
                return `<label><input type="radio" name="blind-speed-type" value="${speedType}" data-total="${total}" ${checked}/> ${label}</label>`;
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
    if (!isConditionAutomationEnabled()) return;
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

async function applyReactiveDamageToActorSocket(actorUuid, value, options = {}) {
    const actor = await resolveActorByUuid(actorUuid);
    if (!actor) return false;
    await applyPlayerDamageWithAutomation(actor, Number(value) || 0, options ?? {});
    return true;
}

async function toggleReactiveConditionSocket(actorUuid, conditionId, enabled) {
    const actor = await resolveActorByUuid(actorUuid);
    if (!actor || !conditionId) return false;
    await actor.setCondition?.(String(conditionId), enabled === true);
    return true;
}

async function toggleReactiveBuffSocket(actorUuid, buffUuid, enabled) {
    return setReactiveBuffState(actorUuid, buffUuid, enabled === true);
}
