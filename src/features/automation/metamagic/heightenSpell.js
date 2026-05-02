export const METAMAGIC_DEFINITION = {
  key: "heightenSpell",
  name: "Heighten Spell",
  get prefix() { return globalThis.game?.i18n?.localize?.("NAS.metamagic.prefixes.heightenSpell") ?? "Heighten"; },
};

export function applyHeightenSpell(context, options = {}) {
  const originalRaw = options.originalLevel ?? context?.spellLevel?.original;
  const targetRaw = options.targetLevel ?? context?.spellLevel?.effective;
  const original = Number(originalRaw ?? 0);
  const target = Number(targetRaw ?? 0);

  if (!Number.isFinite(original) || !Number.isFinite(target)) return false;

  const capped = Math.min(9, Math.max(original, target));
  if (capped <= original) return false;

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  context.metamagic.heightenLevel = capped;
  context.spellLevel = { original, effective: capped };

  if (!context.metamagic.applied.includes(METAMAGIC_DEFINITION.name)) {
    context.metamagic.applied.push(METAMAGIC_DEFINITION.name);
  }

  return true;
}
