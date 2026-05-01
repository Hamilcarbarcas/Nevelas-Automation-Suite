import { MODULE } from '../../../common/module.js';

function getPf1HealthConfig() {
  return game.settings.get('pf1', 'healthConfig');
}

export function isWoundsVigorAutomationEnabled() {
  return Boolean(game.settings.get(MODULE.ID, 'enableWoundsVigorAutomation'));
}

export function isWvNoWoundsConstructUndeadEnabled() {
  return Boolean(game.settings.get(MODULE.ID, 'woundsVigorNoWoundsConstructUndead'));
}

export function actorUsesWoundsVigor(actor) {
  if (!actor) return false;
  try {
    const cfg = getPf1HealthConfig();
    return Boolean(cfg?.getActorConfig?.(actor)?.rules?.useWoundsAndVigor);
  } catch (_err) {
    return false;
  }
}

export function isWoundsVigorActive(actor) {
  if (!isWoundsVigorAutomationEnabled()) return false;
  return actorUsesWoundsVigor(actor);
}

function parseAliasList(value) {
  if (Array.isArray(value)) {
    return value
      .map((entry) => String(entry ?? '').trim().toLowerCase())
      .filter(Boolean);
  }
  const text = String(value ?? '');
  return text
    .split(/[,;|]/)
    .map((entry) => entry.trim().toLowerCase())
    .filter(Boolean);
}

function addCollectionValues(out, source) {
  if (!source) return;
  if (typeof source === 'string') {
    parseAliasList(source).forEach((entry) => out.add(entry));
    return;
  }
  if (Array.isArray(source)) {
    source.forEach((entry) => {
      const text = String(entry ?? '').trim().toLowerCase();
      if (text) out.add(text);
    });
    return;
  }
  if (typeof source.forEach === 'function') {
    source.forEach((entry) => {
      const text = String(entry ?? '').trim().toLowerCase();
      if (text) out.add(text);
    });
    return;
  }
  if (typeof source === 'object') {
    Object.values(source).forEach((entry) => addCollectionValues(out, entry));
  }
}

function collectCreatureTypeEvidence(actor) {
  const out = new Set();
  const traits = actor?.system?.traits ?? {};

  // Trait-based evidence (condition immunities + PF1 type/subtype sets).
  addCollectionValues(out, traits?.ci?.names);
  addCollectionValues(out, traits?.ci?.custom);
  addCollectionValues(out, traits?.creatureTypes?.standard);
  addCollectionValues(out, traits?.creatureTypes?.total);
  addCollectionValues(out, traits?.creatureTypes?.names);
  addCollectionValues(out, traits?.creatureSubtypes?.standard);
  addCollectionValues(out, traits?.creatureSubtypes?.total);
  addCollectionValues(out, traits?.creatureSubtypes?.names);

  // Class/race evidence for edge cases where PF1 booleans are stale or wrong.
  const classes = actor?.classes ?? {};
  Object.entries(classes).forEach(([key, cls]) => {
    const classKey = String(key ?? '').trim().toLowerCase();
    if (classKey) out.add(classKey);
    const className = String(cls?.name ?? '').trim().toLowerCase();
    if (className) out.add(className);
  });

  const raceName = String(actor?.race?.name ?? actor?.system?.details?.race ?? '').trim().toLowerCase();
  if (raceName) out.add(raceName);

  return Array.from(out);
}

function matchNeedles(evidence, needles) {
  if (!evidence.length || !needles.length) return false;
  return evidence.some((entry) => needles.some((needle) => entry.includes(needle)));
}

function getTypeNeedles(kind) {
  const translations = game.settings.get(MODULE.ID, 'translations') || {};
  const translatedTrait = String(translations?.[kind] ?? `${kind} traits`).trim().toLowerCase();

  const classAliasKey = kind === 'construct' ? 'constructClassNames' : 'undeadClassNames';
  const raceAliasKey = kind === 'construct' ? 'constructRaceNames' : 'undeadRaceNames';
  const classAliases = parseAliasList(translations?.[classAliasKey]);
  const raceAliases = parseAliasList(translations?.[raceAliasKey]);

  const needles = new Set([
    kind,
    `${kind} traits`,
    translatedTrait,
    ...classAliases,
    ...raceAliases
  ].map((entry) => String(entry ?? '').trim().toLowerCase()).filter(Boolean));

  return Array.from(needles);
}

function resolveTypeByBooleanAndEvidence({ boolValue, evidence, needles }) {
  const hasEvidence = evidence.length > 0;
  const fromEvidence = matchNeedles(evidence, needles);
  if (!hasEvidence) return Boolean(boolValue);
  if (fromEvidence !== Boolean(boolValue)) return fromEvidence;
  return fromEvidence;
}

export function getWvCreatureTypeState(actor) {
  const traits = actor?.system?.traits ?? {};
  const boolConstruct = Boolean(traits?.construct);
  const boolUndead = Boolean(traits?.undead);
  const boolLiving = traits?.living;
  const evidence = collectCreatureTypeEvidence(actor);

  // String evidence is preferred whenever it contradicts booleans.
  const isConstruct = resolveTypeByBooleanAndEvidence({
    boolValue: boolConstruct,
    evidence,
    needles: getTypeNeedles('construct')
  });
  const isUndead = resolveTypeByBooleanAndEvidence({
    boolValue: boolUndead,
    evidence,
    needles: getTypeNeedles('undead')
  });

  const livingFromTypes = !(isConstruct || isUndead);
  const isLiving = (typeof boolLiving === 'boolean' && boolLiving === livingFromTypes)
    ? boolLiving
    : livingFromTypes;

  return { isConstruct, isUndead, isLiving };
}

export function isWvNoWoundsActor(actor) {
  if (!isWoundsVigorActive(actor)) return false;
  if (!isWvNoWoundsConstructUndeadEnabled()) return false;
  const { isConstruct, isUndead } = getWvCreatureTypeState(actor);
  return isConstruct || isUndead;
}

export function getWvState(actor) {
  const attrs = actor?.system?.attributes ?? {};
  const wounds = attrs.wounds ?? {};
  const vigor = attrs.vigor ?? {};

  const woundsValue = Number(wounds.value ?? 0) || 0;
  const woundsMax = Number(wounds.max ?? 0) || 0;
  const woundThreshold = Number(wounds.threshold ?? 0) || 0;
  const vigorValue = Number(vigor.value ?? 0) || 0;
  const vigorMax = Number(vigor.max ?? 0) || 0;

  const noWounds = isWvNoWoundsActor(actor);
  // House rule: undead/construct use vigor only and are destroyed at 0 vigor.
  const isDead = noWounds ? vigorValue <= 0 : woundsValue <= 0;
  const isWounded = noWounds ? false : (!isDead && woundsValue <= woundThreshold);

  return {
    vigorValue,
    vigorMax,
    woundsValue,
    woundsMax,
    woundThreshold,
    noWounds,
    isDead,
    isWounded
  };
}

export function isWvDead(actor) {
  return getWvState(actor).isDead;
}

export function shouldBeStaggeredFromWv(actor) {
  return getWvState(actor).isWounded;
}

export function isWvWounded(actor) {
  return getWvState(actor).isWounded;
}

export function hasWvHealthUpdate(change) {
  const attrs = change?.system?.attributes;
  if (!attrs) return false;
  const hasWounds =
    attrs.wounds &&
    (attrs.wounds.value !== undefined ||
      attrs.wounds.offset !== undefined ||
      attrs.wounds.threshold !== undefined);
  const hasVigor =
    attrs.vigor &&
    (attrs.vigor.value !== undefined ||
      attrs.vigor.offset !== undefined ||
      attrs.vigor.temp !== undefined);
  return Boolean(hasWounds || hasVigor);
}
