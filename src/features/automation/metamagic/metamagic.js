import { METAMAGIC_DEFINITION as DazingSpell } from "./dazingSpell.js";
import { METAMAGIC_DEFINITION as EmpowerSpell } from "./empowerSpell.js";
import { METAMAGIC_DEFINITION as EnlargeSpell } from "./enlargeSpell.js";
import { METAMAGIC_DEFINITION as ExtendSpell } from "./extendSpell.js";
import { METAMAGIC_DEFINITION as HeightenSpell } from "./heightenSpell.js";
import { METAMAGIC_DEFINITION as IntensifiedSpell } from "./intensifiedSpell.js";
import { METAMAGIC_DEFINITION as MaximizeSpell } from "./maximizeSpell.js";
import { METAMAGIC_DEFINITION as PersistentSpell } from "./persistentSpell.js";
import { METAMAGIC_DEFINITION as QuickenSpell } from "./quickenSpell.js";
import { METAMAGIC_DEFINITION as ReachSpell } from "./reachSpell.js";
import { METAMAGIC_DEFINITION as SelectiveSpell } from "./selectiveSpell.js";
import { METAMAGIC_DEFINITION as StillSpell } from "./stillSpell.js";
import { METAMAGIC_DEFINITION as SilentSpell } from "./silentSpell.js";

export const METAMAGIC_DEFINITIONS = [
  DazingSpell,
  EmpowerSpell,
  EnlargeSpell,
  ExtendSpell,
  HeightenSpell,
  IntensifiedSpell,
  MaximizeSpell,
  PersistentSpell,
  QuickenSpell,
  ReachSpell,
  SelectiveSpell,
  StillSpell,
  SilentSpell,
];

const NAME_LOOKUP = new Map(
  METAMAGIC_DEFINITIONS.map((definition) => [normalizeMetamagicKey(definition.name), definition.name])
);
const PREFIX_LOOKUP = new Map(
  METAMAGIC_DEFINITIONS.map((definition) => [normalizeMetamagicKey(definition.prefix), definition.name])
);

function normalizeMetamagicKey(value) {
  return (value ?? "")
    .toString()
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractRodPrefixFromName(name) {
  const label = (name ?? "").toString();
  if (!label) return "";
  const lower = label.toLowerCase();
  if (!lower.includes("rod") || !lower.includes("metamagic")) return "";

  const ofMatch = label.match(/rod\s+of\s+(.+?)\s+metamagic/i);
  if (ofMatch?.[1]) return normalizeMetamagicKey(ofMatch[1]);

  const rodMatch = label.match(/(.+?)\s+metamagic\s+rod/i);
  if (rodMatch?.[1]) return normalizeMetamagicKey(rodMatch[1]);

  return "";
}

export function resolveMetamagicNameFromDatabase(name) {
  const normalized = normalizeMetamagicKey(name);
  if (!normalized) return null;

  if (NAME_LOOKUP.has(normalized)) return NAME_LOOKUP.get(normalized);

  const rodPrefix = extractRodPrefixFromName(name);
  if (rodPrefix && PREFIX_LOOKUP.has(rodPrefix)) return PREFIX_LOOKUP.get(rodPrefix);

  return null;
}
