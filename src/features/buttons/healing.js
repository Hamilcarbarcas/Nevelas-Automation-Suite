import { MODULE } from '../../common/module.js';
import { elementFromHtmlLike } from '../../common/foundryCompat.js';
import { getHealthModeForElement } from './healthModeChip.js';

const MARK_ATTR = "data-nas-healing-extra";
const HOVER_ATTR = "data-nas-healing-hover";
const ROOT_HOVER_ATTR = "data-nas-healing-root-hover";

export function onRenderChatMessage(html) {
    const root = elementFromHtmlLike(html);
    if (!(root instanceof HTMLElement)) return;
    bindRootHoverHandlers(root);
    addInlineExtras(root);
    addSimpleDamageExtras(root);
}

function addInlineExtras(root) {
    const originals = root.querySelectorAll(`a.inline-action[data-action="applyDamage"]:not([${MARK_ATTR}])`);
    for (const orig of originals) {
        const kind = isHalfInline(orig) ? "half" : "apply";
        if (isNextExtra(orig, kind)) continue;

        const extra = createHealingClone(orig, kind);
        orig.insertAdjacentElement("afterend", extra);
    }
}

function addSimpleDamageExtras(root) {
    const containers = root.querySelectorAll("div.card-buttons");
    for (const container of containers) {
        const originals = Array.from(container.querySelectorAll(`button[data-action="applyDamage"]:not([${MARK_ATTR}])`));
        if (originals.length < 2) continue;

        const applyOrig = originals[0];
        const halfOrig = originals[1];

        if (!isNextExtra(applyOrig, "apply")) {
            const extraApply = createHealingClone(applyOrig, "apply");
            applyOrig.insertAdjacentElement("afterend", extraApply);
        }

        if (!isNextExtra(halfOrig, "half")) {
            const extraHalf = createHealingClone(halfOrig, "half");
            halfOrig.insertAdjacentElement("afterend", extraHalf);
        }
    }
}

function isHalfInline(a) {
    return a.dataset.ratio === "0.5" || a.dataset.tooltip === "PF1.ApplyHalf";
}

function isNextExtra(el, kind) {
    const next = el.nextElementSibling;
    return next?.getAttribute(MARK_ATTR) === MODULE.ID && next?.dataset?.nasHealingKind === kind;
}

function createHealingClone(orig, kind) {
    const extra = orig.cloneNode(true);
    markExtra(extra, kind);

    extra.innerHTML = kind === "half" ? "🩹" : "❤️";
    modifyElementStyles(extra, kind !== "half");
    modifyElementAttributes(extra, kind === "half" ? "Heal Half" : "Heal");

    extra.dataset.nasHealingBaseValue = orig.dataset.value ?? "";
    setNegatedValue(extra, extra.dataset.nasHealingBaseValue);
    extra.dataset.tooltip = "PF1.ApplyHealing";

    bindHealingClick(extra);
    bindHoverVisibility(extra);
    return extra;
}

function markExtra(el, kind) {
    el.setAttribute(MARK_ATTR, MODULE.ID);
    el.dataset.nasHealingKind = kind;
    el.dataset.nasHealingMode = "healing-clone";
    el.classList.add("nas-healing-applyDamage");
    el.setAttribute("aria-label", kind === "half" ? "Apply Half (Healing)" : "Apply (Healing)");
    if (el.tagName === "A") el.setAttribute("role", "button");
}

function bindHoverVisibility(extraEl) {
    const message = extraEl.closest(".chat-attack") ?? extraEl.closest(".chat-message");
    if (!message || message.getAttribute(HOVER_ATTR) === "1") return;
    message.setAttribute(HOVER_ATTR, "1");
}

function toggleHealingButtons(container, visible) {
    const buttons = container.querySelectorAll(`[${MARK_ATTR}]`);
    buttons.forEach((btn) => {
        if (visible) btn.style.visibility = "visible";
    });
}

function bindRootHoverHandlers(root) {
    if (root.getAttribute(ROOT_HOVER_ATTR) === "1") return;
    root.setAttribute(ROOT_HOVER_ATTR, "1");

    root.addEventListener("mouseover", (ev) => {
        const target = ev.target instanceof HTMLElement ? ev.target : null;
        const container = target?.closest?.(".chat-attack") ?? target?.closest?.(".chat-message");
        if (!container) return;
        toggleHealingButtons(container, true);
    }, true);

    root.addEventListener("mouseout", (ev) => {
        const target = ev.target instanceof HTMLElement ? ev.target : null;
        const container = target?.closest?.(".chat-attack") ?? target?.closest?.(".chat-message");
        if (!container) return;

        const related = ev.relatedTarget instanceof HTMLElement ? ev.relatedTarget : null;
        if (related && container.contains(related)) return;
        toggleHealingButtons(container, false);
    }, true);
}

function bindHealingClick(extraEl) {
    if (extraEl.dataset.nasHealingBound === "1") return;
    extraEl.dataset.nasHealingBound = "1";

    extraEl.addEventListener("click", async (ev) => {
        if (ev.button !== 0) return;

        ev.preventDefault();
        ev.stopImmediatePropagation();

        const action = extraEl.dataset.action;
        if (!action) return;

        const ratio = getRatio(extraEl);
        const context = getMessageContext(extraEl);
        const parts = getDamageParts(context);

        let hpTotal = null;
        if (parts.length) {
            hpTotal = await applyAbilityHealing(parts, ratio, context.targets);
        }

        if (hpTotal === null) {
            setNegatedValue(extraEl, extraEl.dataset.nasHealingBaseValue);
        } else if (hpTotal > 0) {
            extraEl.dataset.value = String(-Math.abs(hpTotal));
        } else {
            return;
        }

        const value = Number(extraEl.dataset.value);
        if (!Number.isFinite(value) || value === 0) return;

        try {
            const asWounds = getHealthModeForElement(extraEl) === "wounds";
            const applyOptions = {
                targets: context.targets,
                message: context.message,
                attackIndex: context.attackIndex,
                isCritical: context.isCritical,
                ratio,
                element: extraEl,
                event: ev,
                nasHealing: true,
                asWounds
            };
            await pf1?.documents?.actor?.ActorPF?.applyDamage?.(value, applyOptions);
        } catch (err) {
            console.error(`[${MODULE.ID}] Failed to apply healing`, err);
        }
    }, { capture: true });
}

function setNegatedValue(extraEl, rawValue) {
    const n = Number(rawValue);
    if (Number.isFinite(n)) {
        extraEl.dataset.value = String(-Math.abs(n));
    } else {
        extraEl.dataset.value = rawValue ?? "";
    }
    extraEl.dataset.tooltip = "PF1.ApplyHealing";
}

function getRatio(el) {
    const r = Number(el.dataset.ratio);
    return Number.isFinite(r) ? r : 1;
}

function getMessageContext(el) {
    const chatMessage = el.closest(".chat-message");
    const chatAttack = el.closest(".chat-attack");
    const messageId = chatMessage?.dataset?.messageId ?? null;
    const attackIndex = chatAttack?.dataset?.index ?? null;
    const isCritical = el.dataset.type === "critical" || el.closest('[data-damage-type="critical"]') != null;
    const message = messageId ? game.messages.get(messageId) : null;
    const targets = getTargets();

    return { message, chatAttack, attackIndex, isCritical, targets };
}

function getTargets() {
    let targets = canvas.tokens.controlled;
    if (!targets?.length && game.user.character) targets = [game.user.character];
    return targets.map((t) => t.actor || t).filter((t) => t instanceof Actor);
}

function getDamageParts(context) {
    const fromMessage = getMessageParts(context.message, context.attackIndex, context.isCritical);
    if (fromMessage.length) return fromMessage;

    if (!context.chatAttack) return [];
    const { normal, critical } = collectDomDamageInfo(context.chatAttack);
    const selected = context.isCritical ? [...normal, ...critical] : normal;
    return selected.map((entry) => ({ total: entry.totalDamage, types: entry.damageType }));
}

function getMessageParts(message, attackIndex, isCritical) {
    if (!message) return [];
    const rolls =
        (message.systemRolls && Object.keys(message.systemRolls).length) ? message.systemRolls :
        message.rolls;
    const idx = Number(attackIndex);
    if (!Number.isInteger(idx)) return [];
    const attack = rolls?.attacks?.[idx];
    if (!attack) return [];

    const baseParts = Array.isArray(attack.damage) ? attack.damage : [];
    const critParts = Array.isArray(attack.critDamage) ? attack.critDamage : [];
    const parts = isCritical ? [...baseParts, ...critParts] : baseParts;

    const out = [];
    for (const p of parts) {
        if (!Number.isFinite(p?.total)) continue;
        const dt = p?.options?.damageType;
        const types = (Array.isArray(dt) && dt.length) ? dt : ["untyped"];
        out.push({ total: p.total, types });
    }
    return out;
}

async function applyAbilityHealing(parts, ratio, targets) {
    if (!targets.length) return null;

    let hpTotal = 0;
    const updates = new Map();

    for (const actor of targets) {
        updates.set(actor, foundry.utils.deepClone(actor.system.abilities));
    }

    for (const part of parts) {
        const abilityTags = getAbilityTagsFromTypes(part.types);
        if (!abilityTags.length) {
            hpTotal += part.total;
            continue;
        }

        const amount = Math.floor(-Math.abs(part.total) * ratio);
        if (!Number.isFinite(amount) || amount === 0) continue;

        for (const actor of targets) {
            const abilities = updates.get(actor);
            if (!abilities) continue;

            for (const tag of abilityTags) {
                const abilityKey = tag.ability;
                if (!abilities[abilityKey]) continue;

                switch (tag.type) {
                    case "damage":
                        abilities[abilityKey].damage = Math.max(abilities[abilityKey].damage + amount, 0);
                        break;
                    case "drain":
                        abilities[abilityKey].drain = Math.max(abilities[abilityKey].drain + amount, 0);
                        break;
                    case "penalty":
                        abilities[abilityKey].userPenalty = Math.max(abilities[abilityKey].userPenalty + amount, 0);
                        break;
                }
            }
        }
    }

    for (const [actor, abilities] of updates.entries()) {
        const changes = buildAbilityUpdates(actor, abilities);
        if (Object.keys(changes).length) {
            await actor.update(changes);
        }
    }

    return hpTotal;
}

function getAbilityTagsFromTypes(types) {
    const tags = [];
    for (const typeRef of types ?? []) {
        const entry = resolveDamageType(typeRef);
        const flags = entry?.flags?.[MODULE.ID];
        if (!flags?.vsAbility) continue;

        const abilities = Array.isArray(flags.abilities)
            ? flags.abilities
            : (flags.abilities ? [flags.abilities] : []);
        for (const ability of abilities) {
            tags.push({ ability, type: flags.type });
        }
    }
    return tags;
}

function resolveDamageType(typeRef) {
    const reg = pf1?.registry?.damageTypes;
    if (!reg?.get) return null;
    const direct = reg.get(typeRef);
    if (direct) return direct;

    const needle = String(typeRef).toLowerCase();
    for (const [, value] of reg.entries()) {
        const name = value?.name?.toLowerCase();
        const shortName = value?.shortName?.toLowerCase();
        if (needle === name || needle === shortName) return value;
    }
    return null;
}

function buildAbilityUpdates(actor, newAbilities) {
    const updates = {};
    const current = actor.system.abilities;
    for (const key of Object.keys(newAbilities)) {
        const next = newAbilities[key];
        const cur = current?.[key];
        if (!cur) continue;
        if (next.damage !== cur.damage) updates[`system.abilities.${key}.damage`] = next.damage;
        if (next.drain !== cur.drain) updates[`system.abilities.${key}.drain`] = next.drain;
        if (next.userPenalty !== cur.userPenalty) updates[`system.abilities.${key}.userPenalty`] = next.userPenalty;
    }
    return updates;
}

function collectDomDamageInfo(message) {
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
                let value = null;
                let types = [];
                if (rollCell) value = parseInt(rollCell.textContent.trim(), 10);
                if (typeCell) {
                    types = Array.from(typeCell.querySelectorAll('.damage-type, .custom')).map((dt) =>
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
                    normalDamageTypes = Array.from(normalTDs).flatMap((td) =>
                        Array.from(td.querySelectorAll('.damage-type, .custom')).map((dt) =>
                            dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim()
                        )
                    );
                } else {
                    normalDamageTypes = Array.from(row.querySelectorAll('td.damage-types .damage-type, td.damage-types .custom'))
                        .map((dt) => dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim());
                }
                if (normalRollElement && normalDamageTypes.length > 0) {
                    const totalDamage = parseInt(normalRollElement.textContent.trim(), 10);
                    normalDamageInfo.push({ damageType: normalDamageTypes, totalDamage });
                }
                const criticalRollElement = row.querySelector('td.roll.damage.critical a[data-tooltip]');
                let criticalDamageTypes = [];
                const criticalTDs = row.querySelectorAll('td.damage-type');
                if (criticalTDs.length > 0) {
                    criticalDamageTypes = Array.from(criticalTDs).flatMap((td) =>
                        Array.from(td.querySelectorAll('.damage-type, .custom')).map((dt) =>
                            dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim()
                        )
                    );
                } else {
                    criticalDamageTypes = Array.from(row.querySelectorAll('td.damage-type .damage-type, td.damage-type .custom'))
                        .map((dt) => dt.getAttribute('data-tooltip')?.trim() || dt.textContent.trim());
                }
                if (criticalRollElement && criticalDamageTypes.length > 0) {
                    const totalDamage = parseInt(criticalRollElement.textContent.trim(), 10);
                    criticalDamageInfo.push({ damageType: criticalDamageTypes, totalDamage });
                }
            }
        }
    }
    return { normal: normalDamageInfo, critical: criticalDamageInfo };
}

function modifyElementStyles(element, pulsating = false) {
    element.style.visibility = "visible";
    element.style.display = "inline-block";

    if (pulsating) {
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
        animation.pause();

        element.addEventListener('mouseenter', () => { animation.play(); });
        element.addEventListener('mouseleave', () => { animation.pause(); });

    } else {
        element.addEventListener('mouseenter', () => { element.style.transform = 'scale(1.1)'; });
        element.style.transition = "transform 150ms";
        element.addEventListener('mouseleave', () => { element.style.transform = 'scale(1)'; });
    }
}

function modifyElementAttributes(element, tooltipText) {
    element.setAttribute("data-tooltip", tooltipText);
}
