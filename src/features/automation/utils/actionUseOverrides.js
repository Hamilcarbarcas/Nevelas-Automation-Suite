import { contextHasSavingThrow } from "./saveDcUtils.js";

function setOverride(overrides, obj, path, value) {
  if (!obj || value === undefined) return;
  overrides.push({
    obj,
    path,
    value: foundry.utils.getProperty(obj, path),
  });
  foundry.utils.setProperty(obj, path, value);
}

export function applyActionUseOverrides(actionUse, context) {
  if (!actionUse || !context) return () => {};

  const overrides = [];

  const item = actionUse.item;
  const action = actionUse.action;

  if (context.components) {
    Object.entries(context.components).forEach(([key, value]) => {
      // Use document path to ensure downstream chat rendering reads the projected component values.
      const path = `system.components.${key}`;
      setOverride(overrides, item, path, value);
    });
  }

  if (context.range) {
    setOverride(overrides, action, "touch", context.range.touch);
    setOverride(overrides, action, "hasRange", context.range.hasRange);
    if (context.range.range) {
      setOverride(overrides, action, "range.value", context.range.range.value);
      setOverride(overrides, action, "range.minUnits", context.range.range.minUnits);
      setOverride(overrides, action, "range.minValue", context.range.range.minValue);
      setOverride(overrides, action, "range.units", context.range.range.units);
      setOverride(overrides, action, "range.maxIncrements", context.range.range.maxIncrements);
    }
  }

  if (context.save && contextHasSavingThrow(context, actionUse)) {
    setOverride(overrides, action, "save.dc", context.save.dc);
    setOverride(overrides, action, "save.type", context.save.type);
    setOverride(overrides, action, "save.description", context.save.description);
    setOverride(overrides, action, "save.harmless", context.save.harmless);
  }

  if (context.duration) {
    setOverride(overrides, action, "duration.value", context.duration.value);
    setOverride(overrides, action, "duration.units", context.duration.units);
    setOverride(overrides, action, "duration.dismiss", context.duration.dismiss);
    setOverride(overrides, action, "duration.concentration", context.duration.concentration);
  }

  if (context.damageOverrides?.parts && Array.isArray(context.damageOverrides.parts)) {
    context.damageOverrides.parts.forEach((entry) => {
      if (!entry || typeof entry.index !== "number") return;
      const path = entry.isArray
        ? `damage.parts.${entry.index}.0`
        : `damage.parts.${entry.index}.formula`;
      setOverride(overrides, action, path, entry.formula);
    });
  }

  if (context.damage) {
    if (!context.damageOverrides?.parts) {
      setOverride(overrides, action, "damage.parts", context.damage.parts);
    }
    setOverride(overrides, action, "damage.critParts", context.damage.critParts);
    setOverride(overrides, action, "damage.nonCritParts", context.damage.nonCritParts);
  }

  if (context.alignments) {
    if (context.alignments.item) {
      Object.entries(context.alignments.item).forEach(([key, value]) => {
        setOverride(overrides, item, `system.alignments.${key}`, value);
      });
    }
    if (context.alignments.resolved) {
      Object.entries(context.alignments.resolved).forEach(([key, value]) => {
        setOverride(overrides, action, `alignments.${key}`, value);
      });
    }
  }

  if (context.actionType !== undefined) {
    setOverride(overrides, action, "actionType", context.actionType);
  }

  if (context.spellLevel?.effective !== undefined) {
    setOverride(overrides, item, "system.level", context.spellLevel.effective);
  }

  if (context.notes?.footer && Array.isArray(context.notes.footer)) {
    setOverride(overrides, action, "notes.footer", context.notes.footer);
  }

  return () => {
    for (let i = overrides.length - 1; i >= 0; i -= 1) {
      const { obj, path, value } = overrides[i];
      foundry.utils.setProperty(obj, path, value);
    }
  };
}
