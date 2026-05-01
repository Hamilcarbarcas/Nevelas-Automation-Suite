import { MODULE } from '../../../common/module.js';

const REACTIVE_FLAG_KEY = "itemReactiveEffects";

function remapItemReactiveOnHitIdKeyedByActionId(source, idMap) {
    if (!idMap || idMap.size === 0) return {};
    if (!source || typeof source !== "object" || Array.isArray(source)) {
        return {};
    }
    const out = {};
    for (const [oldId, value] of Object.entries(source)) {
        const newId = idMap.get(String(oldId));
        if (newId) {
            out[newId] = foundry.utils.deepClone(value);
        }
    }
    return out;
}

function itemAttackFromItemFlagCopy(wrapped, item) {
    const data = wrapped(item);

    const flags = item.flags?.[MODULE.ID] || {};
    data.flags ||= {};
    data.flags[MODULE.ID] = foundry.utils.mergeObject(flags, data.flags[MODULE.ID] || {});

    const newActions = data.system?.actions;
    const nasFlags = data.flags[MODULE.ID];
    const copiedActions = nasFlags?.itemActionSettings?.actions;
    if (Array.isArray(newActions) && Array.isArray(copiedActions) && newActions.length === copiedActions.length) {
        const idMap = new Map();
        for (let i = 0; i < copiedActions.length; i++) {
            const oldId = copiedActions[i]?.id;
            const newId = newActions[i]?._id;
            if (oldId != null && newId != null) {
                idMap.set(String(oldId), String(newId));
            }
            copiedActions[i].id = newId;
        }
        const reactive = nasFlags?.[REACTIVE_FLAG_KEY];
        if (reactive) {
            if (reactive.onHitByAction != null && typeof reactive.onHitByAction === "object" && !Array.isArray(reactive.onHitByAction)) {
                reactive.onHitByAction = remapItemReactiveOnHitIdKeyedByActionId(reactive.onHitByAction, idMap);
            }
            if (reactive.onHitByActionOverride != null && typeof reactive.onHitByActionOverride === "object" && !Array.isArray(reactive.onHitByActionOverride)) {
                reactive.onHitByActionOverride = remapItemReactiveOnHitIdKeyedByActionId(
                    reactive.onHitByActionOverride,
                    idMap
                );
            }
        }
    }

    return data;
}

export function initItemAttackFlagCopy() {
    if (!globalThis.libWrapper) {
        console.error(`[${MODULE.ID}] libWrapper is required for item attack flag copying.`);
        return;
    }
    globalThis.libWrapper.register(
        MODULE.ID,
        'pf1.documents.item.ItemAttackPF.fromItem',
        itemAttackFromItemFlagCopy,
        'WRAPPER'
    );
} 
