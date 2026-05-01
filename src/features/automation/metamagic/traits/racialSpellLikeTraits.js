import { resolveEnglishName } from "../../utils/compendiumNameResolver.js";

const ARCANE_PRODIGY_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-traits.Item.Yyc3F8awZSrjjcBd";
const ARCANE_PRODIGY_ENGLISH_NAME = "Arcane Prodigy (Drow)";
const ARCANE_PRODIGY_ID = "arcaneProdigyDrow";

const PROLONG_MAGIC_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-traits.Item.gyZegJ9ctVXfzfXq";
const PROLONG_MAGIC_ENGLISH_NAME = "Prolong Magic (Tiefling)";
const PROLONG_MAGIC_ID = "prolongMagicTiefling";

const CURATOR_MYSTIC_SECRETS_COMPENDIUM_SOURCE = "Compendium.pf-content.pf-traits.Item.3ryjl46Ui9o1uqvR";
const CURATOR_MYSTIC_SECRETS_ENGLISH_NAME = "Curator of Mystic Secrets";
const CURATOR_MYSTIC_SECRETS_ID = "curatorOfMysticSecrets";

function normalizeKey(value) {
  return (value ?? "").toString().trim().toLowerCase();
}

function isEnglishLanguage() {
  return (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
}

function isFeatLikeItem(item) {
  const subType = item?.subType ?? item?.system?.subType;
  return item?.type === "feat" && (subType === "trait" || subType === "feat" || subType === "classFeat");
}

function getBabeleOriginalName(item) {
  const direct = item?.flags?.babele?.originalName;
  const canUseBabele = game?.modules?.get?.("babele")?.active === true;
  const hasGetFlag = canUseBabele && typeof item?.getFlag === "function";

  let viaGetFlag = null;
  try {
    viaGetFlag = hasGetFlag ? item.getFlag("babele", "originalName") : null;
  } catch (_err) {
    viaGetFlag = null;
  }

  return direct ?? viaGetFlag ?? null;
}

function getFastMatchReason(item, { compendiumSource, englishName }) {
  const source = item?._stats?.compendiumSource ?? "";
  if (source && source === compendiumSource) return "compendiumSource";
  const originalName = getBabeleOriginalName(item);
  if (originalName && originalName === englishName) return "babele.originalName";
  if (isEnglishLanguage() && (item?.name ?? "") === englishName) return "name";
  return null;
}

async function findActorTrait(actor, { compendiumSource, englishName }) {
  const items = actor?.items;
  if (!items) return null;

  const feats = Array.from(items).filter(isFeatLikeItem);
  for (const feat of feats) {
    const reason = getFastMatchReason(feat, { compendiumSource, englishName });
    if (reason) return feat;
  }

  if (isEnglishLanguage()) return null;
  for (const feat of feats) {
    const resolved = await resolveEnglishName(feat?.name, { documentName: "Item", deepScanMode: "off" });
    if ((resolved ?? "") === englishName) return feat;
  }

  return null;
}

function getUseState(item) {
  const uses = item?.system?.uses;
  if (!uses || !uses.per) {
    return {
      limited: false,
      hasUsesData: false,
      remaining: null,
      max: null,
      hasRemaining: true
    };
  }
  const remaining = Number(uses.value ?? 0);
  const max = Number(uses.max ?? 0);
  return {
    limited: true,
    hasUsesData: true,
    remaining: Number.isFinite(remaining) ? remaining : 0,
    max: Number.isFinite(max) ? max : 0,
    hasRemaining: Number.isFinite(remaining) && remaining > 0
  };
}

function isSpellLikeSpellbookItem(actor, spellItem) {
  if (!actor || !spellItem) return false;
  if (spellItem.type !== "spell") return false;
  const spellbookKey = normalizeKey(spellItem?.system?.spellbook);
  if (spellbookKey !== "spelllike") return false;
  const abilityType = normalizeKey(spellItem?.system?.abilityType);
  // Spellbook "spelllike" is the authoritative PF1 discriminator for these spell entries.
  // Some spelllike entries omit abilityType at item level, so only reject explicit non-sp values.
  if (abilityType && abilityType !== "sp") return false;
  const actorBook = actor?.system?.attributes?.spells?.spellbooks?.[spellbookKey];
  return Boolean(actorBook);
}

async function resolveSpellNameCandidates(spellItem) {
  const localized = spellItem?.name ?? "";
  const babeleOriginal = getBabeleOriginalName(spellItem);
  const resolvedEnglish = await resolveEnglishName(localized, { documentName: "Item", deepScanMode: "off" });
  return [localized, babeleOriginal, resolvedEnglish]
    .map(normalizeKey)
    .filter(Boolean);
}

const ARCANE_PRODIGY_SPELLS = {
  uuids: new Set([
    "Compendium.pf1.spells.Item.tjog6bufg5b08lvq",
    "Compendium.pf1.spells.Item.vr35zocojcg7gei8",
    "Compendium.pf1.spells.Item.zymaptg3vmnvfvxl",
    "Compendium.pf1.spells.Item.tsndfcfijmgxs37p",
    "Compendium.pf1.spells.Item.jdsvncnna6oy189a",
    "Compendium.pf1.spells.Item.uqh87jz757r2cb7r",
    "Compendium.pf1.spells.Item.gmgwyjfpeuuc4t4o",
    "Compendium.pf1.spells.Item.21bxbzdaawjrnvyo",
    "Compendium.pf1.spells.Item.ech2cibcsmsms9s7",
    "Compendium.pf1.spells.Item.bl71og1gklwncmt7",
    "Compendium.pf1.spells.Item.68m9du7zw2di7ew6",
    "Compendium.pf1.spells.Item.mtxqp85izkb20djq",
    "Compendium.pf1.spells.Item.plou8h168bfn5hq6",
    "Compendium.pf1.spells.Item.3mfmhx8avu7h3iom",
    "Compendium.pf1.spells.Item.zqj5qzyl46af27v0"
  ]),
  names: new Set([
    "charm person",
    "cloak of shade",
    "dancing lights",
    "darkness",
    "deeper darkness",
    "detect magic",
    "dispel magic",
    "divine favor",
    "dust of twilight",
    "faerie fire",
    "feather fall",
    "ghost sound",
    "levitate",
    "spider climb",
    "suggestion"
  ])
};

const PROLONG_MAGIC_SPELLS = {
  uuids: new Set([
    "Compendium.pf1.spells.Item.kouqz0pm1xl8xilm",
    "Compendium.pf1.spells.Item.8uwmrygxgih1fb57",
    "Compendium.pf1.spells.Item.k9iu3d82hlo7coct",
    "Compendium.pf1.spells.Item.usdv1eqvibmxun6x",
    "Compendium.pf1.spells.Item.18kryvj2dfuymk9d",
    "Compendium.pf1.spells.Item.p2kosvizylhy8vfa",
    "Compendium.pf1.spells.Item.c0hvjnqzry149ugs",
    "Compendium.pf1.spells.Item.tsndfcfijmgxs37p",
    "Compendium.pf1.spells.Item.k0k8gvqyy3rd5gke",
    "Compendium.pf1.spells.Item.1hjxxr3k62rcpb5c",
    "Compendium.pf1.spells.Item.tr7m97npkbgm4wp7",
    "Compendium.pf1.spells.Item.xllxylvvqr82o2d5",
    "Compendium.pf1.spells.Item.g33euis7yi9pwddy",
    "Compendium.pf1.spells.Item.5nr9o7o0it6ewf17",
    "Compendium.pf1.spells.Item.qooag8g9dxck4lik",
    "Compendium.pf1.spells.Item.p8lm7khq3ynyif21",
    "Compendium.pf1.spells.Item.27o3msobhyghmfid",
    "Compendium.pf1.spells.Item.8nrt26t37v7koqr8",
    "Compendium.pf1.spells.Item.0d756xyqhinsstas",
    "Compendium.pf1.spells.Item.8u1xa5javcxc6szk",
    "Compendium.pf1.spells.Item.q63jn0kbrrb6t3ea",
    "Compendium.pf1.spells.Item.byn5q46du0ck592a",
    "Compendium.pf1.spells.Item.0onqjy8gfgop1xsi",
    "Compendium.pf1.spells.Item.j9qpy0d585f8lpm4"
  ]),
  names: new Set([
    "alter self",
    "animate dead",
    "animate objects",
    "bear's endurance",
    "blur",
    "commune",
    "curse water",
    "darkness",
    "death knell",
    "deathwatch",
    "detect evil",
    "detect thoughts",
    "fog cloud",
    "hideous laughter",
    "inflict light wounds",
    "mage hand",
    "minor image",
    "misdirection",
    "pyrotechnics",
    "rage",
    "shatter",
    "speak with dead",
    "ventriloquism",
    "web"
  ])
};

async function isSpellAllowedForTrait(spellItem, allowed) {
  const compendiumSource = spellItem?._stats?.compendiumSource ?? "";
  if (compendiumSource && allowed.uuids.has(compendiumSource)) return true;
  const candidates = await resolveSpellNameCandidates(spellItem);
  return candidates.some((name) => allowed.names.has(name));
}

export async function getRacialSpellLikeTraitSources(actor, spellItem) {
  if (!actor || !spellItem) return [];

  const [curatorItem, arcaneProdigyItem, prolongMagicItem] = await Promise.all([
    findActorTrait(actor, {
      compendiumSource: CURATOR_MYSTIC_SECRETS_COMPENDIUM_SOURCE,
      englishName: CURATOR_MYSTIC_SECRETS_ENGLISH_NAME
    }),
    findActorTrait(actor, {
      compendiumSource: ARCANE_PRODIGY_COMPENDIUM_SOURCE,
      englishName: ARCANE_PRODIGY_ENGLISH_NAME
    }),
    findActorTrait(actor, {
      compendiumSource: PROLONG_MAGIC_COMPENDIUM_SOURCE,
      englishName: PROLONG_MAGIC_ENGLISH_NAME
    })
  ]);

  const sources = [];
  if (curatorItem) {
    const use = getUseState(curatorItem);
    if (!use.limited || use.hasRemaining) {
      sources.push({
        id: CURATOR_MYSTIC_SECRETS_ID,
        label: curatorItem?.name ?? CURATOR_MYSTIC_SECRETS_ENGLISH_NAME,
        itemUuid: curatorItem?.uuid ?? null,
        limited: use.limited,
        hasUsesData: use.hasUsesData,
        hasRemaining: use.hasRemaining,
        usesValue: use.remaining,
        usesMax: use.max,
        effectType: "castTimeBypass",
        requiresMetamagicIntent: true
      });
    }
  }

  const eligibleSpelllike = isSpellLikeSpellbookItem(actor, spellItem);
  if (!eligibleSpelllike) return sources;

  const allowedArcane = arcaneProdigyItem ? await isSpellAllowedForTrait(spellItem, ARCANE_PRODIGY_SPELLS) : false;
  const allowedProlong = prolongMagicItem ? await isSpellAllowedForTrait(spellItem, PROLONG_MAGIC_SPELLS) : false;

  if (arcaneProdigyItem && allowedArcane) {
    const use = getUseState(arcaneProdigyItem);
    if (!use.limited || use.hasRemaining) {
      sources.push({
        id: ARCANE_PRODIGY_ID,
        label: arcaneProdigyItem?.name ?? ARCANE_PRODIGY_ENGLISH_NAME,
        itemUuid: arcaneProdigyItem?.uuid ?? null,
        limited: use.limited,
        hasUsesData: use.hasUsesData,
        hasRemaining: use.hasRemaining,
        usesValue: use.remaining,
        usesMax: use.max,
        effectType: "durationExtension",
        requiresMetamagicIntent: false
      });
    }
  }

  if (prolongMagicItem && allowedProlong) {
    const use = getUseState(prolongMagicItem);
    if (!use.limited || use.hasRemaining) {
      sources.push({
        id: PROLONG_MAGIC_ID,
        label: prolongMagicItem?.name ?? PROLONG_MAGIC_ENGLISH_NAME,
        itemUuid: prolongMagicItem?.uuid ?? null,
        limited: use.limited,
        hasUsesData: use.hasUsesData,
        hasRemaining: use.hasRemaining,
        usesValue: use.remaining,
        usesMax: use.max,
        effectType: "durationExtension",
        requiresMetamagicIntent: false
      });
    }
  }

  return sources;
}

