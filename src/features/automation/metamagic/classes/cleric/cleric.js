import { resolveEnglishName } from "../../../utils/compendiumNameResolver.js";

const HEALERS_BLESSING_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.a7yviTyteK4X7A0c";
const HEALERS_BLESSING_ENGLISH_NAME = "Healer's Blessing";
const HEALING_DOMAIN_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.lqCm6Fy7JVezprqU";
const HEALING_DOMAIN_ENGLISH_NAME = "Healing Domain";
const HEALERS_BLESSING_REQUIRED_LEVEL = 6;
const HEALERS_BLESSING_CLASS_NAMES = ["cleric", "inquisitor"];
const INTENSE_CELEBRATION_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.1BuSzuaMzcrxfwmn";
const INTENSE_CELEBRATION_ENGLISH_NAME = "Intense Celebration";
const REVELRY_SUBDOMAIN_COMPENDIUM_SOURCE = "Compendium.pf1.class-abilities.Item.LdTgQd53LA9JUQaj";
const REVELRY_SUBDOMAIN_ENGLISH_NAME = "Revelry Subdomain";
const INTENSE_CELEBRATION_REQUIRED_LEVEL = 8;
const INTENSE_CELEBRATION_CLASS_NAMES = ["cleric"];
const INTENSE_CELEBRATION_ELIGIBLE_UUIDS = new Set([
  "Compendium.pf1.spells.Item.h2i514vklrry38fj",
  "Compendium.pf1.spells.Item.ibk7jrc5rwubpia6",
  "Compendium.pf1.spells.Item.gt6958kgaauhi0dc",
  "Compendium.pf1.spells.Item.etng7hnjsqqyoy8i",
  "Compendium.pf1.spells.Item.w5ctdt1qctqt5let",
  "Compendium.pf1.spells.Item.wa0zb2pncesmm9lz",
  "Compendium.pf1.spells.Item.272szq5d556toqgc",
  "Compendium.pf1.spells.Item.m8nue3zocbg7c5o1",
  "Compendium.pf1.spells.Item.g09p4w3w8enp5tki",
  "Compendium.pf1.spells.Item.6fbwb3gq0g36n7r9",
  "Compendium.pf1.spells.Item.MyyQSa5H4qYabx18",
  "Compendium.pf1.spells.Item.c9a27yvlb5mrow83",
  "Compendium.pf-content.pf-occult-rituals.Item.8uYPKrlfEN3KT6bo",
  "Compendium.pf1.spells.Item.3wcvosazm31i8758",
  "Compendium.pf1.spells.Item.dirqq4kmbkt9gc1m",
  "Compendium.pf1.spells.Item.qvcbd606owi2zeix",
  "Compendium.pf1.spells.Item.mkn0jklea8knq1lz",
  "Compendium.pf1.spells.Item.e0rlki4hwu76l5ya",
  "Compendium.pf1.spells.Item.vln322jeis7ghbfi",
  "Compendium.pf1.spells.Item.jiu8a7raurkroyfx",
  "Compendium.pf1.spells.Item.x2j14gyvqbpmwrqb",
  "Compendium.pf1.spells.Item.cfztz4xbdlo5yqjd",
  "Compendium.pf1.spells.Item.rgd5v3llwoptwcdb",
  "Compendium.pf-content.pf-occult-rituals.Item.jNYEoaGTONqXu90z",
  "Compendium.pf1.spells.Item.yx70o17kxbcut6ni",
  "Compendium.pf1.spells.Item.op40qjf9oohlx5nu",
  "Compendium.pf1.spells.Item.ueuz3ymuz8pxpzr6",
  "Compendium.pf1.spells.Item.vqfrp8t0c1lw1jna",
  "Compendium.pf1.spells.Item.z0duc2v2n3ioynta",
  "Compendium.pf1.spells.Item.htpbmbmgzito5qx2",
  "Compendium.pf1.spells.Item.429wzxjzm1chnsq7",
  "Compendium.pf-content.pf-occult-rituals.Item.Rpf905wlSL8HXgHV",
  "Compendium.pf1.spells.Item.6ux76jy9wbi88br0",
  "Compendium.pf1.spells.Item.c7j4fxzlxyvll2t4",
  "Compendium.pf1.spells.Item.cXuzSqklJEBXhZqg",
  "Compendium.pf1.spells.Item.md6ltvqh9paadsm4",
  "Compendium.pf1.spells.Item.71g6nmvy8qnodwyz",
  "Compendium.pf1.spells.Item.ekhygw5fzfe52n35",
  "Compendium.pf1.spells.Item.skgjjub1hng29nwo",
  "Compendium.pf1.spells.Item.yarcdn0xxlarbii5",
  "Compendium.pf1.spells.Item.8kd54d2ho1fywj2u",
  "Compendium.pf1.spells.Item.opdh00yu3ud5z867",
  "Compendium.pf1.spells.Item.6sc2szgfkr6xb71l",
  "Compendium.pf1.spells.Item.nw7jzpmn7rh7qlp5",
  "Compendium.pf1.spells.Item.lygujduijp64f0wj",
  "Compendium.pf1.spells.Item.8u1xa5javcxc6szk",
  "Compendium.pf1.spells.Item.i973i4haczi7e0xp",
  "Compendium.pf1.spells.Item.eu4u4mr1naoiqohz",
  "Compendium.pf1.spells.Item.770nro22rs4fxz7y",
  "Compendium.pf1.spells.Item.dm52jk1e37b05v9d",
  "Compendium.pf1.spells.Item.ggc6xepnsthjsmth",
  "Compendium.pf1.spells.Item.l1emjlFvcCg4Ebnb",
  "Compendium.pf1.spells.Item.8237jjhtwfqje4n6",
  "Compendium.pf1.spells.Item.4uzseabeqamqnmb6",
  "Compendium.pf1.spells.Item.eimp8bzk9oabtqpj",
  "Compendium.pf1.spells.Item.p4h6ocn0axh59xrc",
  "Compendium.pf1.spells.Item.UkSdEOtHdzeuRiUu",
  "Compendium.pf1.spells.Item.suilv5ug3tvbwuh2",
  "Compendium.pf1.spells.Item.dw5k7pl4sz07pxke",
  "Compendium.pf1.spells.Item.5zrqn8bj3jxm05pl"
].map((value) => normalizeKey(value)));
const INTENSE_CELEBRATION_ELIGIBLE_IDS = new Set(
  Array.from(INTENSE_CELEBRATION_ELIGIBLE_UUIDS)
    .map((value) => value.split(".").pop())
    .filter(Boolean)
);

function normalizeKey(value) {
  return (value ?? "").toString().trim().toLowerCase();
}

function isEnglishLanguage() {
  return (game?.i18n?.lang ?? "en").toLowerCase().startsWith("en");
}

function isFeatItem(item) {
  const subType = item?.subType ?? item?.system?.subType;
  return item?.type === "feat" && (subType === "feat" || subType === "classFeat");
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

function getFeatFastMatchReason(item, { compendiumSource, englishName }) {
  const source = item?._stats?.compendiumSource ?? "";
  if (source && source === compendiumSource) return "compendiumSource";

  const originalName = getBabeleOriginalName(item);
  if (originalName && originalName === englishName) return "babele.originalName";

  if (isEnglishLanguage() && (item?.name ?? "") === englishName) return "name";
  return null;
}

async function findFeat(actor, { compendiumSource, englishName }) {
  const items = actor?.items;
  if (!items) return null;

  const feats = Array.from(items).filter(isFeatItem);
  if (!feats.length) return null;

  for (const feat of feats) {
    const reason = getFeatFastMatchReason(feat, { compendiumSource, englishName });
    if (reason) return feat;
  }

  if (isEnglishLanguage()) return null;
  for (const feat of feats) {
    const resolved = await resolveEnglishName(feat?.name, { documentName: "Item", deepScanMode: "off" });
    if ((resolved ?? "") === englishName) return feat;
  }

  return null;
}

async function resolveSpellEnglishName(spellItem) {
  if (!spellItem) return "";
  const babeleOriginal = getBabeleOriginalName(spellItem);
  if (typeof babeleOriginal === "string" && babeleOriginal.trim().length > 0) {
    return babeleOriginal.trim();
  }
  if (isEnglishLanguage()) {
    return (spellItem?.name ?? "").toString().trim();
  }
  const resolved = await resolveEnglishName(spellItem?.name, { documentName: "Item", deepScanMode: "off" });
  return (resolved ?? spellItem?.name ?? "").toString().trim();
}

async function isHealersBlessingEligibleSpell(spellItem) {
  if (!spellItem || spellItem.type !== "spell") return false;
  const englishName = (await resolveSpellEnglishName(spellItem)).toLowerCase();
  if (!englishName) return false;
  return englishName.startsWith("cure ");
}

function isIntenseCelebrationEligibleSpell(spellItem) {
  if (!spellItem) return false;
  const compendiumSource = normalizeKey(spellItem?._stats?.compendiumSource ?? "");
  const sourceId = normalizeKey(spellItem?.flags?.core?.sourceId ?? "");
  const uuid = normalizeKey(spellItem?.uuid ?? "");
  const id = normalizeKey(spellItem?.id ?? "");
  if (INTENSE_CELEBRATION_ELIGIBLE_UUIDS.has(compendiumSource)) return true;
  if (INTENSE_CELEBRATION_ELIGIBLE_UUIDS.has(sourceId)) return true;
  if (INTENSE_CELEBRATION_ELIGIBLE_UUIDS.has(uuid)) return true;
  if (INTENSE_CELEBRATION_ELIGIBLE_IDS.has(id)) return true;
  return false;
}

function getHighestClassLevel(actor, classNames = []) {
  const targetNames = new Set((classNames ?? []).map((value) => normalizeKey(value)).filter(Boolean));
  if (!targetNames.size) return 0;

  const levels = [];
  for (const [classKey, classData] of Object.entries(actor?.classes ?? {})) {
    const candidates = [
      classKey,
      classData?.tag,
      classData?.name,
      classData?._id,
      classData?.id
    ]
      .map(normalizeKey)
      .filter(Boolean);

    const matches = candidates.some((candidate) => {
      if (targetNames.has(candidate)) return true;
      for (const target of targetNames) {
        if (candidate.includes(target)) return true;
      }
      return false;
    });
    if (!matches) continue;

    const level = Number(classData?.level ?? 0);
    if (Number.isFinite(level) && level > 0) levels.push(level);
  }

  return levels.length ? Math.max(...levels) : 0;
}

export async function getHealersBlessingFeatureSources(actor, spellItem) {
  if (!actor || !spellItem || spellItem.type !== "spell") return [];
  const classLevel = getHighestClassLevel(actor, HEALERS_BLESSING_CLASS_NAMES);
  if (classLevel < HEALERS_BLESSING_REQUIRED_LEVEL) return [];

  const [healersBlessingItem, healingDomainItem, isEligibleSpell] = await Promise.all([
    findFeat(actor, {
      compendiumSource: HEALERS_BLESSING_COMPENDIUM_SOURCE,
      englishName: HEALERS_BLESSING_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: HEALING_DOMAIN_COMPENDIUM_SOURCE,
      englishName: HEALING_DOMAIN_ENGLISH_NAME
    }),
    isHealersBlessingEligibleSpell(spellItem)
  ]);

  if (!isEligibleSpell || !healersBlessingItem || !healingDomainItem) return [];

  return [{
    id: "healersBlessing",
    label: healersBlessingItem?.name ?? HEALERS_BLESSING_ENGLISH_NAME,
    itemUuid: healersBlessingItem?.uuid ?? null,
    limited: false,
    persistent: true,
    defaultEnabled: true,
    hasUsesData: true,
    hasRemaining: true,
    dependencies: {
      healingDomainItemUuid: healingDomainItem?.uuid ?? null
    }
  }];
}

export async function getIntenseCelebrationFeatureSources(actor, spellItem) {
  if (!actor || !spellItem) return [];
  const classLevel = getHighestClassLevel(actor, INTENSE_CELEBRATION_CLASS_NAMES);
  if (classLevel < INTENSE_CELEBRATION_REQUIRED_LEVEL) return [];
  if (!isIntenseCelebrationEligibleSpell(spellItem)) return [];

  const [intenseCelebrationItem, revelrySubdomainItem] = await Promise.all([
    findFeat(actor, {
      compendiumSource: INTENSE_CELEBRATION_COMPENDIUM_SOURCE,
      englishName: INTENSE_CELEBRATION_ENGLISH_NAME
    }),
    findFeat(actor, {
      compendiumSource: REVELRY_SUBDOMAIN_COMPENDIUM_SOURCE,
      englishName: REVELRY_SUBDOMAIN_ENGLISH_NAME
    })
  ]);
  if (!intenseCelebrationItem || !revelrySubdomainItem) return [];

  return [{
    id: "intenseCelebration",
    label: intenseCelebrationItem?.name ?? INTENSE_CELEBRATION_ENGLISH_NAME,
    itemUuid: intenseCelebrationItem?.uuid ?? null,
    limited: false,
    persistent: true,
    defaultEnabled: true,
    hasUsesData: true,
    hasRemaining: true,
    dependencies: {
      revelrySubdomainItemUuid: revelrySubdomainItem?.uuid ?? null
    }
  }];
}

