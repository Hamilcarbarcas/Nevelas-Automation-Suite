import { MODULE } from "../../../common/module.js";
import { elementFromHtmlLike } from "../../../common/foundryCompat.js";

function isSpellActor(actor) {
  return actor && actor.documentName === "Actor";
}

function getSpellbook(actor, spellbookKey) {
  if (!isSpellActor(actor) || !spellbookKey) return null;
  return actor.system?.attributes?.spells?.spellbooks?.[spellbookKey] ?? null;
}

function normalizeSpellbookKey(actor, spellbookKey) {
  const key = (spellbookKey ?? "").toString().trim();
  if (key) return key;
  if (actor?.system?.attributes?.spells?.spellbooks?.primary) return "primary";
  return "";
}

function getAbilityTotal(actor, abilityKey) {
  const ability = actor?.system?.abilities?.[abilityKey];
  const total = Number(ability?.total ?? ability?.value ?? 0);
  return Number.isFinite(total) ? total : 0;
}

function getAbilityShortLabel(abilityKey) {
  const key = (abilityKey ?? "").toString();
  return (
    pf1?.config?.abilitiesShort?.[key]
    ?? pf1?.config?.abilities?.[key]
    ?? key.toUpperCase()
  );
}

export function isSpellAbilityMinimumEnabled() {
  try {
    return game.settings.get(MODULE.ID, "enforceSpellAbilityMinimum") === true;
  } catch (_err) {
    return false;
  }
}

export function evaluateSpellAbilityEligibility(actor, {
  spellbookKey = "",
  spellLevel = 0,
  honorNoAbilityLimit = true
} = {}) {
  const level = Number(spellLevel ?? 0);
  if (!Number.isFinite(level) || level < 0) {
    return { allowed: true, reason: "invalidLevel" };
  }

  const book = getSpellbook(actor, spellbookKey);
  if (!book) {
    return { allowed: true, reason: "missingSpellbook" };
  }

  if (honorNoAbilityLimit && book.noAbilityLimit === true) {
    return { allowed: true, reason: "noAbilityLimit" };
  }

  const abilityKey = (book.ability ?? "").toString();
  const abilityTotal = getAbilityTotal(actor, abilityKey);
  const required = 10 + level;
  const levelData = book.spells?.[`spell${level}`] ?? null;
  const lowAbilityFlag = levelData?.lowAbilityScore === true;
  const allowed = lowAbilityFlag ? false : abilityTotal >= required;

  return {
    allowed,
    reason: allowed ? "ok" : (lowAbilityFlag ? "lowAbilityFlag" : "insufficientAbility"),
    spellbookKey,
    spellLevel: level,
    abilityKey,
    abilityLabel: getAbilityShortLabel(abilityKey),
    abilityTotal,
    required,
    noAbilityLimit: book.noAbilityLimit === true
  };
}

export function evaluateSpellItemAbilityEligibility(actor, spellItem, {
  spellbookKey,
  spellLevel,
  honorNoAbilityLimit = true
} = {}) {
  if (!spellItem || spellItem.type !== "spell" || !isSpellActor(actor)) {
    return { allowed: true, reason: "notSpellItem" };
  }

  const nextBook = normalizeSpellbookKey(actor, spellbookKey ?? spellItem.system?.spellbook ?? "");
  const nextLevelRaw = spellLevel ?? spellItem.system?.level ?? 0;
  const nextLevel = Number(nextLevelRaw);
  if (!nextBook || !Number.isFinite(nextLevel)) {
    return { allowed: true, reason: "missingBookOrLevel" };
  }

  return evaluateSpellAbilityEligibility(actor, {
    spellbookKey: nextBook,
    spellLevel: nextLevel,
    honorNoAbilityLimit
  });
}

function notifyBlocked(actionKey, spellName, eligibility) {
  ui.notifications.warn(game.i18n.format("NAS.spellcasting.abilityMinimumBlocked", {
    action: game.i18n.localize(actionKey),
    spell: spellName,
    level: Number(eligibility?.spellLevel ?? 0),
    required: Number(eligibility?.required ?? 0),
    current: Number(eligibility?.abilityTotal ?? 0),
    ability: (eligibility?.abilityLabel ?? "").toString()
  }));
}

export function enforceSpellAbilityMinimumOnActionUse(actionUse) {
  if (!isSpellAbilityMinimumEnabled()) return true;

  const item = actionUse?.item ?? null;
  if (!item || item.type !== "spell") return true;
  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? null;
  const eligibility = evaluateSpellItemAbilityEligibility(actor, item, { honorNoAbilityLimit: true });
  if (eligibility.allowed) return true;

  notifyBlocked("NAS.spellcasting.actions.cast", item.name ?? game.i18n.localize("PF1.Spell"), eligibility);
  actionUse.shared ??= {};
  actionUse.shared.reject = true;
  actionUse.shared.scriptData ??= {};
  actionUse.shared.scriptData.reject = true;
  return false;
}

export function enforceSpellAbilityMinimumOnPreCreateItem(item, data, userId) {
  if (!isSpellAbilityMinimumEnabled()) return true;
  const itemType = data?.type ?? item?.type;
  if (itemType !== "spell") return true;

  const actor = item?.actor ?? item?.parent ?? null;
  if (!isSpellActor(actor)) return true;

  const spellbookKey = normalizeSpellbookKey(actor, foundry.utils.getProperty(data, "system.spellbook"));
  const spellLevel = Number(foundry.utils.getProperty(data, "system.level") ?? 0);
  const eligibility = evaluateSpellAbilityEligibility(actor, {
    spellbookKey,
    spellLevel,
    honorNoAbilityLimit: true
  });
  if (eligibility.allowed) return true;

  if (game.user?.id === userId) {
    notifyBlocked("NAS.spellcasting.actions.add", data?.name ?? game.i18n.localize("PF1.Spell"), eligibility);
  }
  return false;
}

export function enforceSpellAbilityMinimumOnPreUpdateItem(item, change, userId) {
  if (!isSpellAbilityMinimumEnabled()) return true;
  if (item?.type !== "spell") return true;
  const actor = item?.actor ?? null;
  if (!isSpellActor(actor)) return true;

  const hasLevelChange = foundry.utils.getProperty(change, "system.level") !== undefined;
  const hasBookChange = foundry.utils.getProperty(change, "system.spellbook") !== undefined;
  if (!hasLevelChange && !hasBookChange) return true;

  const nextBook = normalizeSpellbookKey(actor, foundry.utils.getProperty(change, "system.spellbook") ?? item.system?.spellbook);
  const nextLevel = Number(foundry.utils.getProperty(change, "system.level") ?? item.system?.level ?? 0);
  const eligibility = evaluateSpellAbilityEligibility(actor, {
    spellbookKey: nextBook,
    spellLevel: nextLevel,
    honorNoAbilityLimit: true
  });
  if (eligibility.allowed) return true;

  if (game.user?.id === userId) {
    notifyBlocked("NAS.spellcasting.actions.update", item?.name ?? game.i18n.localize("PF1.Spell"), eligibility);
  }
  return false;
}

function parseSpellListDataKey(dataList) {
  const key = (dataList ?? "").toString();
  const match = /^spell-(.+)-(\d+)$/.exec(key);
  if (!match) return null;
  return {
    spellbookKey: match[1],
    spellLevel: Number(match[2])
  };
}

export function applySpellAbilityMinimumSheetVisibility(app, html) {
  const root = elementFromHtmlLike(html);
  if (!root) return;

  const actor = app?.actor ?? app?.object ?? null;
  if (!isSpellActor(actor)) return;

  const lists = root.querySelectorAll('ol.item-list[data-list^="spell-"][data-level]');
  if (!lists?.length) return;

  const shouldHideIneligible = isSpellAbilityMinimumEnabled();
  for (const listEl of lists) {
    const key = parseSpellListDataKey(listEl.dataset.list);
    if (!key) continue;

    let hide = false;
    if (shouldHideIneligible) {
      const eligibility = evaluateSpellAbilityEligibility(actor, {
        spellbookKey: key.spellbookKey,
        spellLevel: key.spellLevel,
        honorNoAbilityLimit: true
      });
      hide = !eligibility.allowed;
    }

    listEl.classList.toggle("nas-hidden-by-ability-limit", hide);
    listEl.style.display = hide ? "none" : "";
  }
}
