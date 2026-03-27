import { MODULE } from '../../../common/module.js';
import { abilityDeltaCalculation } from './ability.js';
import { buildAbilityDmgEntries, splitAbilityInstances } from './abilityTags.js';
import { applyNasDefenseBypass } from './nasBypass.js';
import { applyLegacyPriorityTypes } from './priorityTypes.js';
import { applyManualVulnerability, buildNumericInstances, sumInstanceValues } from './instances.js';
import { normalizeTargets } from '../utils/targeting.js';

export function registerSystemApplyDamage() {
    if (!globalThis.libWrapper) return false;

    function markDialogUse(wrapped, ...args) {
        const rv = wrapped(...args);
        rv._nasDamageDialog = true;
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
        "pf1.documents.actor.ActorPF.applyDamage",
        async function (wrapped, value = 0, options = {}) {
            if (value === 0 || !Number.isFinite(value)) return wrapped(value, options);

            if (options._nasDamageDialog) return wrapped(value, options);

            let showDialog = (typeof options.dialog === "boolean") ? options.dialog : false;
            if (options.event?.shiftKey) showDialog = !showDialog;
            if (showDialog) return wrapped(value, options);

            if (!pf1?.applications?.ApplyDamage) return wrapped(value, options);

            const targets = normalizeTargets(options);
            if (!targets.length) return false;

            const isHealing = (value < 0) || options.isHealing === true;
            const ratio = Number(options?.ratio);
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
                    await target.update(updates);
                }
            }

            const hpValue = sumInstanceValues(hpInstances);
            calcOpts.value = isHealing ? -Math.abs(hpValue) : Math.abs(hpValue);

            if (isHealing) {
                const app = new pf1.applications.ApplyDamage(calcOpts);

                if (hasRatio) {
                    for (const target of app.targets) {
                        target.ratio = ratio;
                    }
                    if (typeof app._refreshTargets === "function") {
                        app._refreshTargets();
                    }
                }

                applyLegacyPriorityTypes(app, options, calcOpts.instances, isHealing);
                applyNasDefenseBypass(app, options, { isHealing });

                if (options?.element?.dataset && isHealing && hasRatio && ratio === 0.5) {
                    options.element.dataset.tooltip = "PF1.ApplyHalf";
                }

                let appliedValue = Math.max(0, app.value);
                if (app.isHealing) appliedValue = -appliedValue;
                appliedValue += app.bonus || 0;

                const promises = [];
                for (const target of app.targets) {
                    const applyDamageOpts = app._getTargetDamageOptions(target);
                    applyDamageOpts._nasDamageDialog = true;

                    promises.push(wrapped.call(target.actor, appliedValue, applyDamageOpts));
                }

                return Promise.all(promises);
            }

            const promises = [];
            for (const actor of targets) {
                const dv = actor?.system?.traits?.dv;
                const vulnInstances = applyManualVulnerability(calcOpts.instances, dv);
                const vulnTotal = sumInstanceValues(vulnInstances);
                const vulnValue = Math.abs(vulnTotal);
                const targetOpts = {
                    ...calcOpts,
                    targets: [actor],
                    instances: vulnInstances,
                    value: vulnValue
                };

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

                promises.push(wrapped.call(actor, appliedValue, applyDamageOpts));
            }

            return Promise.all(promises);
        },
        libWrapper.MIXED
    );

    return true;
}
