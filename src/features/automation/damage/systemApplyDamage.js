import { MODULE } from '../../../common/module.js';
import { checkMassiveDamage } from '../../../integration/moduleSockets.js';
import { abilityDeltaCalculation } from './ability.js';
import { buildAbilityDmgEntries, splitAbilityInstances } from './abilityTags.js';
import { applyNasDefenseBypass } from './nasBypass.js';
import { applyLegacyPriorityTypes } from './priorityTypes.js';
import { applyManualVulnerability, buildNumericInstances, sumInstanceValues } from './instances.js';
import { normalizeTargets } from '../utils/targeting.js';
import { isWoundsVigorActive, isWoundsVigorAutomationEnabled, isWvNoWoundsActor } from '../utils/woundsVigor.js';
import { getCasterLevelEquivalentFromFormula, getDiceCountFromFormula, getDiceCountFromRoll } from '../utils/formulaUtils.js';
import { applyReactiveEffectsForHit, resolveSourceActorFromOptions } from './reactiveTriggers.js';
import { resolveMirrorImageForApplyDamage } from '../buffs/mirrorImage.js';
import { recordCombatTextContext } from '../utils/combatTextContext.js';

function getConfiguredWoundDamageTypeIds() {
    const fallback = ['negative', 'positive'];
    try {
        const raw = game.settings.get(MODULE.ID, 'woundsVigorWoundDamageTypeIds');
        if (!Array.isArray(raw)) return new Set(fallback);
        return new Set(raw.map((id) => String(id).toLowerCase()).filter(Boolean));
    } catch (_err) {
        return new Set(fallback);
    }
}

function getMessageDamageParts(options) {
    const msg = options?.message;
    const rolls =
        (msg?.systemRolls && Object.keys(msg.systemRolls).length) ? msg.systemRolls :
        msg?.rolls;
    const idx = Number(options?.attackIndex);
    if (!Number.isInteger(idx)) return [];
    const attack = rolls?.attacks?.[idx];
    if (!attack) return [];
    const baseParts = Array.isArray(attack.damage) ? attack.damage : [];
    const critParts = Array.isArray(attack.critDamage) ? attack.critDamage : [];
    return options?.isCritical ? [...baseParts, ...critParts] : baseParts;
}

function countDiceFromMessageParts(options) {
    const parts = getMessageDamageParts(options);
    if (!parts.length) return 0;
    const rollData = options?.action?.getRollData?.()
        ?? options?.item?.getRollData?.()
        ?? {};
    let total = 0;
    for (const part of parts) {
        const fromRoll = getDiceCountFromRoll(part?.roll);
        if (fromRoll > 0) {
            total += fromRoll;
            continue;
        }
        total += getDiceCountFromFormula(part?.formula ?? part?.roll?.formula, rollData);
    }
    return total;
}

function countDiceFromInstances(instances, rollData = {}) {
    if (!Array.isArray(instances) || !instances.length) return 0;
    let total = 0;
    for (const inst of instances) {
        const fromRoll = getDiceCountFromRoll(inst?.roll);
        if (fromRoll > 0) {
            total += fromRoll;
            continue;
        }
        total += getDiceCountFromFormula(inst?.formula ?? inst?.roll?.formula, rollData);
    }
    return total;
}

function getTypeIdsFromInstance(inst) {
    const raw = inst?.typeIds ?? inst?.types ?? [];
    return Array.from(raw).map((t) => String(t).toLowerCase());
}

function resolveDamageType(typeId) {
    const reg = pf1?.registry?.damageTypes;
    if (!reg?.get) return null;
    const direct = reg.get(typeId);
    if (direct) return direct;
    const needle = String(typeId).toLowerCase();
    for (const [, value] of reg.entries()) {
        const id = String(value?.id ?? '').toLowerCase();
        const name = String(value?.name ?? '').toLowerCase();
        const shortName = String(value?.shortName ?? '').toLowerCase();
        if (needle && (needle === id || needle === name || needle === shortName)) return value;
    }
    return null;
}

function matchesConfiguredWoundType(typeId, configuredTypeIds) {
    const entry = resolveDamageType(typeId);
    const raw = String(typeId).toLowerCase();
    if (configuredTypeIds.has(raw)) return true;
    const id = String(entry?.id ?? raw).toLowerCase();
    const name = String(entry?.name ?? '').toLowerCase();
    const shortName = String(entry?.shortName ?? '').toLowerCase();
    return configuredTypeIds.has(id) || configuredTypeIds.has(name) || configuredTypeIds.has(shortName);
}

function isEnergyTypeId(typeId) {
    const entry = resolveDamageType(typeId);
    return Boolean(entry?.isEnergy || entry?.category === 'energy');
}

function gatherMessagePartTypeIds(options) {
    const out = [];
    const parts = getMessageDamageParts(options);
    for (const part of parts) {
        const dt = part?.options?.damageType;
        if (Array.isArray(dt)) out.push(...dt.map((t) => String(t).toLowerCase()));
    }
    return out;
}

function isLifeEnergyContext(options, instances) {
    const typeIds = [];
    for (const inst of instances ?? []) typeIds.push(...getTypeIdsFromInstance(inst));
    if (!typeIds.length) typeIds.push(...gatherMessagePartTypeIds(options));
    if (!typeIds.length) return false;

    const configuredTypeIds = getConfiguredWoundDamageTypeIds();
    const hasLifeType = typeIds.some((t) => matchesConfiguredWoundType(t, configuredTypeIds));
    if (!hasLifeType) return false;

    const hasEnergy = typeIds.some((t) => isEnergyTypeId(t)) || hasLifeType;
    return hasLifeType && hasEnergy;
}

function inferCasterLevel(options, actor) {
    const fromAction = Number(options?.action?.getRollData?.()?.cl);
    if (Number.isFinite(fromAction) && fromAction > 0) return Math.floor(fromAction);

    const fromItem = Number(options?.item?.getRollData?.()?.cl);
    if (Number.isFinite(fromItem) && fromItem > 0) return Math.floor(fromItem);

    const bookKey = options?.action?.item?.system?.spellbook ?? options?.item?.system?.spellbook;
    const fromBook = Number(actor?.system?.attributes?.spells?.spellbooks?.[bookKey]?.cl?.total);
    if (Number.isFinite(fromBook) && fromBook > 0) return Math.floor(fromBook);

    return 0;
}

function looksCasterLevelScaled(options, instances) {
    const item = options?.item ?? options?.action?.item ?? options?.message?.itemSource;
    if (item?.type === 'spell') return true;

    const formulas = [];
    for (const inst of instances ?? []) {
        if (typeof inst?.formula === 'string') formulas.push(inst.formula);
    }
    for (const part of getMessageDamageParts(options)) {
        if (typeof part?.formula === 'string') formulas.push(part.formula);
        if (typeof part?.roll?.formula === 'string') formulas.push(part.roll.formula);
    }
    return formulas.some((f) => /@cl|\/\s*level|caster\s*level|per\s*level/i.test(f));
}

function getRuleWoundMagnitude(options, instances, actor, fallbackValue) {
    const rollData = options?.action?.getRollData?.()
        ?? options?.item?.getRollData?.()
        ?? actor?.getRollData?.()
        ?? {};
    const casterLevel = inferCasterLevel(options, actor);
    const diceCount = countDiceFromMessageParts(options) || countDiceFromInstances(instances, rollData);
    if (diceCount > 0) return diceCount;

    const formulas = new Set();
    for (const part of getMessageDamageParts(options)) {
        if (typeof part?.formula === 'string') formulas.add(part.formula);
        if (typeof part?.roll?.formula === 'string') formulas.add(part.roll.formula);
    }
    for (const inst of instances ?? []) {
        if (typeof inst?.formula === 'string') formulas.add(inst.formula);
        if (typeof inst?.roll?.formula === 'string') formulas.add(inst.roll.formula);
    }

    let clEquivalentTotal = 0;
    for (const formula of formulas) {
        clEquivalentTotal += getCasterLevelEquivalentFromFormula(formula, { rollData, casterLevel });
    }
    if (clEquivalentTotal > 0) return clEquivalentTotal;

    if (looksCasterLevelScaled(options, instances) && casterLevel > 0) {
        return casterLevel;
    }

    return Math.max(0, Math.floor(Math.abs(Number(fallbackValue) || 0)));
}

function getApplyModeFromElement(options) {
    const mode = options?.element?.closest?.('.chat-attack')?.getAttribute?.('data-nas-health-mode');
    return mode === 'wounds' ? 'wounds' : 'vigor';
}

function resistanceCheckWouldForceDialog(options) {
    const moduleId = "pf1-resistance-check";
    if (!game.modules.get(moduleId)?.active) return false;
    if (options?._nasDamageDialog) return false;

    try {
        if (!game.settings.get(moduleId, "enableCheck")) return false;
    } catch (_err) {
        return false;
    }

    const relevantKeys = ["dr", "di", "dv", "eres", "cr", "ci"];
    const keysToCheck = [];
    for (const key of relevantKeys) {
        try {
            if (game.settings.get(moduleId, `check${key}`)) {
                keysToCheck.push(key);
            }
        } catch (_err) {
        }
    }
    if (!keysToCheck.length) return false;

    const evt = options?.event ?? globalThis.event;
    if (evt?.shiftKey || evt?.ctrlKey || evt?.metaKey) return false;

    const emptyIWR = {
        ci: [],
        cres: "",
        di: [],
        dr: { value: [], custom: "" },
        dv: [],
        eres: { value: [], custom: "" }
    };

    for (const token of canvas?.tokens?.controlled ?? []) {
        const traits = token?.actor?.toObject?.()?.system?.traits;
        if (!traits) continue;
        const traitDiffKeys = Object.keys(foundry.utils.diffObject(emptyIWR, traits));
        if (traitDiffKeys.some((k) => keysToCheck.includes(k))) return true;
    }

    return false;
}

function resolveCriticalWoundBonus({ options, actor, applyDamageOpts, appliedValue, preVigor, preTemp, isCritical }) {
    if (!isCritical) return 0;
    if (!isWoundsVigorActive(actor)) return 0;
    if (isWvNoWoundsActor(actor)) return 0;
    if (options?.asNonlethal) return 0;

    const critMult = Math.max(0, Number(options?.critMult ?? applyDamageOpts?.critMult ?? 0) || 0);
    if (critMult <= 0) return 0;

    if (options?.asWounds) return critMult;

    let finalDamage = Math.floor(Math.max(0, appliedValue) * (applyDamageOpts?.ratio ?? 1));
    finalDamage -= Math.min(finalDamage, applyDamageOpts?.reduction ?? 0);
    finalDamage = Math.floor(finalDamage);
    const vigorBuffer = Math.max(0, Number(preVigor) || 0) + Math.max(0, Number(preTemp) || 0);
    if (finalDamage > vigorBuffer) return 0;

    return critMult;
}

function recordDamageCombatTextContext(actor, options = {}) {
    if (options?.isCritical !== true) return;
    recordCombatTextContext(actor, {
        isCritical: true,
        critMult: options?.critMult,
        messageUuid: options?.message?.uuid ?? options?.reference ?? null,
        attackIndex: options?.attackIndex
    });
}

function tokenDocumentKey(tokenDocument) {
    return String(tokenDocument?.uuid ?? tokenDocument?.id ?? "");
}

function messageTargetTokenDocuments(message) {
    const ids = message?.system?.targets;
    if (!Array.isArray(ids) || !ids.length) return [];
    return ids.map((id) => {
        return canvas?.scene?.tokens?.get?.(id) ?? globalThis.fromUuidSync?.(id) ?? null;
    }).filter(Boolean);
}

function hasInvalidControlledTargets(options = {}) {
    const cardTargets = messageTargetTokenDocuments(options.message);
    const controlled = Array.from(canvas?.tokens?.controlled ?? []).map((token) => token.document).filter(Boolean);
    if (!cardTargets.length || !controlled.length) return false;
    const cardTargetKeys = new Set(cardTargets.map(tokenDocumentKey).filter(Boolean));
    return controlled.some((tokenDocument) => !cardTargetKeys.has(tokenDocumentKey(tokenDocument)));
}

function buildChatApplyDamageData(message, eventLike) {
    const button = eventLike?.currentTarget ?? eventLike?.target;
    if (button?.dataset?.action !== "applyDamage") return null;

    let asNonlethal;
    if (message?.system?.config?.nonlethal) asNonlethal = true;
    const tags = new Set(button.dataset.tags?.split(";") ?? []);
    if (tags.has("nonlethal")) asNonlethal = true;

    const value = parseInt(button.dataset.value);
    if (Number.isNaN(value)) return null;

    const attackIndex = parseInt(button.closest("[data-index]")?.dataset.index);
    const attackType = button.dataset.type;
    const attack = message?.systemRolls?.attacks?.[attackIndex];
    const isCritical = attackType === "critical";
    const instances = [];
    const addInstances = (damageRolls) => {
        if (!damageRolls) return;
        for (const dmg of damageRolls) {
            const d = new pf1.models.action.DamagePartModel(dmg.damageType.toObject());
            d.value = dmg.total;
            instances.push(d);
        }
    };

    if (attack) {
        addInstances(attack.damage);
        if (isCritical) addInstances(attack.critDamage);
    }

    const item = message.itemSource;
    const action = message.actionSource;
    return {
        value,
        options: {
            asNonlethal,
            event: eventLike,
            element: button,
            message,
            item,
            action,
            attackIndex,
            reference: message.uuid,
            isCritical,
            critMult: isCritical ? (message.system.config.critMult ?? 0) : 0,
            instances,
            interactive: true
        }
    };
}

export async function applyNasHeadlessDamage(value = 0, options = {}) {
    if (!game.settings.get(MODULE.ID, "enableDamageAutomation")) {
        return { handled: false };
    }

    if (value === 0 || !Number.isFinite(value)) return { handled: false };

    if (isWoundsVigorAutomationEnabled() && !options.asWounds && getApplyModeFromElement(options) === "wounds") {
        options = { ...options, asWounds: true };
    }

    let showDialog = (typeof options.dialog === "boolean") ? options.dialog : false;
    if (options.event?.shiftKey) showDialog = !showDialog;
    if (showDialog) return { handled: false };

    if (resistanceCheckWouldForceDialog(options)) {
        return { handled: false, options: { ...options, dialog: true } };
    }

    if (!pf1?.applications?.ApplyDamage) return { handled: false };

    if (hasInvalidControlledTargets(options)) {
        return { handled: false };
    }

    const targets = normalizeTargets(options);
    if (!targets.length) return { handled: true, result: false };

    const isHealing = (value < 0) || options.isHealing === true;
    const sourceActor = resolveSourceActorFromOptions(options);
    const sourceItem = options?.item ?? options?.action?.item ?? options?.message?.itemSource ?? null;

    const hasExplicitTargets = options?.targets != null;
    const hasClickContext = options?.element != null || options?.event != null || options?.message != null;
    if (isHealing && hasExplicitTargets && !hasClickContext) {
        return { handled: false };
    }

    let ratio = Number(options?.ratio);
    if (!(Number.isFinite(ratio) && ratio > 0)) {
        const elRatio = Number(options?.element?.dataset?.ratio);
        if (Number.isFinite(elRatio) && elRatio > 0 && Array.isArray(options?.instances) && options.instances.length) {
            ratio = elRatio;
        }
    }
    const hasRatio = Number.isFinite(ratio) && ratio > 0 && ratio !== 1;

    const calcOpts = {
        ...options,
        targets,
        value,          
    };

    const allInstances = buildNumericInstances(Math.abs(value), calcOpts);
    const { abilityInstances, hpInstances } = splitAbilityInstances(allInstances);
    calcOpts.instances = hpInstances;

    if (abilityInstances.length && !isHealing) {
        const abilityDmg = buildAbilityDmgEntries(abilityInstances, hasRatio ? ratio : 1, false);
        for (const target of targets) {
            const traits = target.system.traits;
            const abilities = foundry.utils.deepClone(target.system.abilities);
            abilityDeltaCalculation(traits.di, traits.ci, abilities, abilityDmg);
            const updates = {};
            for (const key in abilities) {
                updates[`system.abilities.${key}.damage`] = abilities[key].damage;
                updates[`system.abilities.${key}.drain`] = abilities[key].drain;
                updates[`system.abilities.${key}.userPenalty`] = abilities[key].userPenalty;
            }
            recordDamageCombatTextContext(target, options);
            await target.update(updates);
        }
    }

    const hpValue = sumInstanceValues(hpInstances);
    if (isHealing) {
        calcOpts.value = -Math.abs(hpValue);
    } else if (hpValue > 0) {
        calcOpts.value = Math.abs(hpValue);
    } else {
        calcOpts.value = Math.abs(Number(value) || 0);
    }

    if (isHealing) {
        const promises = [];
        for (const actor of targets) {
            const targetInstances = calcOpts.instances;
            let targetValue = Math.abs(calcOpts.value);
            const targetAsWounds = Boolean(options?.asWounds) && !isWvNoWoundsActor(actor);
            const targetOpts = {
                ...calcOpts,
                targets: [actor],
                instances: targetInstances,
                value: -Math.abs(targetValue),
                asWounds: targetAsWounds
            };

            const useRuleWounds = Boolean(targetOpts?.asWounds) && isWoundsVigorActive(actor);
            const useRuleMagnitude = useRuleWounds && (isHealing || isLifeEnergyContext(targetOpts, targetInstances));
            if (useRuleMagnitude) {
                targetValue = getRuleWoundMagnitude(targetOpts, targetInstances, actor, targetValue);
                targetOpts.value = -Math.abs(targetValue);
            }

            const app = new pf1.applications.ApplyDamage(targetOpts);

            if (hasRatio) {
                for (const target of app.targets) {
                    target.ratio = ratio;
                }
                if (typeof app._refreshTargets === "function") {
                    app._refreshTargets();
                }
            }

            applyLegacyPriorityTypes(app, options, targetOpts.instances, isHealing);
            applyNasDefenseBypass(app, options, { isHealing });

            if (options?.element?.dataset && hasRatio && ratio === 0.5) {
                options.element.dataset.tooltip = "PF1.ApplyHalf";
            }

            let appliedValue = Math.max(0, app.value);
            if (app.isHealing) appliedValue = -appliedValue;
            appliedValue += app.bonus || 0;

            const targetModel = app.targets.get(actor.uuid) ?? app.targets.first();
            const applyDamageOpts = app._getTargetDamageOptions(targetModel);
            applyDamageOpts._nasDamageDialog = true;

            recordDamageCombatTextContext(actor, options);
            promises.push(actor.applyDamage(appliedValue, applyDamageOpts));
        }

        return { handled: true, result: Promise.all(promises) };
    }

    const promises = [];
    for (const actor of targets) {
        const dv = actor?.system?.traits?.dv;
        const vulnInstances = applyManualVulnerability(calcOpts.instances, dv);
        const vulnTotal = sumInstanceValues(vulnInstances);
        let vulnValue;
        if (isHealing) {
            vulnValue = Math.abs(vulnTotal);
        } else if (vulnTotal > 0) {
            vulnValue = Math.abs(vulnTotal);
        } else {
            vulnValue = Math.abs(Number(calcOpts.value) || 0);
        }
        const targetAsWounds = Boolean(options?.asWounds) && !isWvNoWoundsActor(actor);
        const targetOpts = {
            ...calcOpts,
            targets: [actor],
            instances: vulnInstances,
            value: vulnValue,
            asWounds: targetAsWounds
        };

        const useRuleWounds = Boolean(targetOpts?.asWounds) && isWoundsVigorActive(actor);
        const useRuleMagnitude = useRuleWounds && (isHealing || isLifeEnergyContext(targetOpts, vulnInstances));
        if (useRuleMagnitude) {
            vulnValue = getRuleWoundMagnitude(targetOpts, vulnInstances, actor, vulnValue);
            targetOpts.value = isHealing ? -Math.abs(vulnValue) : Math.abs(vulnValue);
        }

        const app = new pf1.applications.ApplyDamage(targetOpts);

        if (hasRatio) {
            for (const target of app.targets) {
                target.ratio = ratio;
            }
            if (typeof app._refreshTargets === "function") {
                app._refreshTargets();
            }
        }

        applyLegacyPriorityTypes(app, options, targetOpts.instances, false);
        applyNasDefenseBypass(app, options, { isHealing: false });

        if (options?.element?.dataset && isHealing && hasRatio && ratio === 0.5) {
            options.element.dataset.tooltip = "PF1.ApplyHalf";
        }

        let appliedValue = Math.max(0, app.value);
        if (app.isHealing) appliedValue = -appliedValue;
        appliedValue += app.bonus || 0;

        const targetModel = app.targets.get(actor.uuid) ?? app.targets.first();
        const applyDamageOpts = app._getTargetDamageOptions(targetModel);
        applyDamageOpts._nasDamageDialog = true;

        promises.push((async () => {
            const mirrorImage = await resolveMirrorImageForApplyDamage({
                sourceActor,
                targetActor: actor,
                options
            });
            if (mirrorImage?.blockDamage) {
                return false;
            }

            const preVigor = Number(actor?.system?.attributes?.vigor?.value ?? 0) || 0;
            const preTemp = Number(actor?.system?.attributes?.vigor?.temp ?? 0) || 0;
            recordDamageCombatTextContext(actor, options);
            const result = await actor.applyDamage(appliedValue, applyDamageOpts);

            const critBonus = resolveCriticalWoundBonus({
                options,
                actor,
                applyDamageOpts,
                appliedValue,
                preVigor,
                preTemp,
                isCritical: Boolean(options?.isCritical)
            });
            if (result && critBonus > 0) {
                recordDamageCombatTextContext(actor, options);
                await actor.applyDamage(critBonus, {
                    asWounds: true,
                    _nasDamageDialog: true
                });
            }

            let finalDamage = Math.floor(Math.max(0, appliedValue) * (applyDamageOpts?.ratio ?? 1));
            finalDamage -= Math.min(finalDamage, applyDamageOpts?.reduction ?? 0);
            finalDamage = Math.floor(finalDamage);

            if (result && finalDamage >= 0 && !isHealing) {
                await applyReactiveEffectsForHit({
                    sourceActor,
                    sourceItem,
                    targetActor: actor,
                    options,
                    finalDamage
                });
            }

            if (result && game.settings.get(MODULE.ID, "massiveDamage")) {
                const token =
                    actor?.token?.object ??
                    actor?.token ??
                    actor?.getActiveTokens?.(true, true)?.[0] ??
                    actor?.getActiveTokens?.()?.[0] ??
                    null;

                if (token) {
                    if (finalDamage > 0) {
                        const maxHP = actor?.system?.attributes?.hp?.max ?? 0;
                        checkMassiveDamage(finalDamage, maxHP, token);
                    }
                }
            }

            return result;
        })());
    }

    return { handled: true, result: Promise.all(promises) };
}

export async function applyNasChatDamageButton(message, eventLike) {
    const button = eventLike?.currentTarget ?? eventLike?.target;
    if (button?.dataset?.action !== "applyDamage") return { handled: false };
    if (message?.flags?.[MODULE.ID]?.source === "command") return { handled: false };

    const data = buildChatApplyDamageData(message, eventLike);
    if (!data) {
        return { handled: false };
    }

    const applied = await applyNasHeadlessDamage(data.value, data.options);
    if (!applied?.handled) {
        if (applied?.options) {
            pf1?.documents?.actor?.ActorPF?.applyDamage?.(data.value, applied.options);
            button.disabled = false;
            return { handled: true, result: false };
        }
        return { handled: false };
    }

    eventLike?.preventDefault?.();
    button.disabled = false;
    return { handled: true, result: applied.result };
}

export function registerSystemApplyDamage() {
    if (!globalThis.libWrapper) return false;

    function markDialogUse(wrapped, ...args) {
        const rv = wrapped(...args);
        if (rv && typeof rv === "object") {
            rv._nasDamageDialog = true;
        }
        return rv;
    }

    libWrapper.register(
        MODULE.ID,
        "pf1.applications.ApplyDamage.prototype._getTargetDamageOptions",
        markDialogUse,
        libWrapper.WRAPPER
    );

    libWrapper.register(
        MODULE.ID,
        "pf1.documents.actor.ActorPF.prototype.applyDamage",
        async function (wrapped, value = 0, options = {}) {
            if (!options?._nasDamageDialog) return wrapped.call(this, value, options);
            if (!Boolean(options?.asWounds)) return wrapped.call(this, value, options);
            if (!isWoundsVigorActive(this)) return wrapped.call(this, value, options);

            if (isWvNoWoundsActor(this)) {
                return wrapped.call(this, value, { ...options, asWounds: false });
            }

            const isHealing = (value < 0) || options?.isHealing === true;
            const instances = Array.isArray(options?.instances) ? options.instances : [];
            const useRuleMagnitude = isHealing || isLifeEnergyContext(options, instances);
            if (!useRuleMagnitude) return wrapped.call(this, value, options);

            const adjustedMagnitude = getRuleWoundMagnitude(
                options,
                instances,
                this,
                Math.abs(Number(value) || 0)
            );
            const adjustedValue = isHealing
                ? -Math.abs(adjustedMagnitude)
                : Math.abs(adjustedMagnitude);

            return wrapped.call(this, adjustedValue, options);
        },
        libWrapper.WRAPPER
    );

    return true;
}
