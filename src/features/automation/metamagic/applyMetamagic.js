import { MODULE } from "../../../common/module.js";
import { METAMAGIC_DEFINITION as StillSpell, applyStillSpell } from "./stillSpell.js";
import { METAMAGIC_DEFINITION as SilentSpell, applySilentSpell } from "./silentSpell.js";
import {
  METAMAGIC_DEFINITION as ExtendSpell,
  applyExtendDurationOnly,
  applyExtendSpell,
  applyExtendSpellWithMaskFocus,
  isDurationEligibleForExtendSpell
} from "./extendSpell.js";
import { METAMAGIC_DEFINITION as EnlargeSpell, applyEnlargeSpell } from "./enlargeSpell.js";
import { METAMAGIC_DEFINITION as ReachSpell, applyReachSpell } from "./reachSpell.js";
import { METAMAGIC_DEFINITION as QuickenSpell, QUICKEN_SPELL_NAME, applyQuickenSpell } from "./quickenSpell.js";
import { METAMAGIC_DEFINITION as SelectiveSpell, applySelectiveSpell } from "./selectiveSpell.js";
import { METAMAGIC_DEFINITION as DazingSpell, applyDazingSpell } from "./dazingSpell.js";
import { METAMAGIC_DEFINITION as HeightenSpell, applyHeightenSpell } from "./heightenSpell.js";
import { METAMAGIC_DEFINITION as PersistentSpell, applyPersistentSpell } from "./persistentSpell.js";
import { METAMAGIC_DEFINITION as EmpowerSpell } from "./empowerSpell.js";
import { applyEmpowerToFormula } from "./empowerSpell.js";
import { METAMAGIC_DEFINITION as IntensifiedSpell, canIntensifyAnyDamagePart } from "./intensifiedSpell.js";
import { METAMAGIC_DEFINITION as MaximizeSpell } from "./maximizeSpell.js";
import { resolveMetamagicNameFromDatabase } from "./metamagic.js";
import {
  getHealersBlessingFeatureSources,
  getIntenseCelebrationFeatureSources,
  getNaniteBloodlineArcanaFeatureSources,
  getOneBodyTwoMindsFeatureSources,
  getOracleSeekerFeatureSources,
  getOracleSuccorFinalRevelationFeatureSources,
  getPeerlessSpeedFeatureSources,
  getPsychicMimicMetamagicState,
  getPsychicSpellbookMaxCastableSpellLevel,
  getTimelessSoulFeatureSources,
  getMetamixingState,
  getSorcererArcaneMetamagicState,
  getSpellbookMaxCastableSpellLevel,
  getWizardMetamagicMasteryState,
  MIMIC_METAMAGIC_FEATURE_ID,
  METAMAGIC_MASTERY_FEATURE_ID,
  METAMIXING_FEATURE_ID
} from "./classes/index.js";
import {
  getMagicalLineageSlotAdjustment,
  getRacialSpellLikeTraitSources,
  getTransmuterOfKoradaSource,
  getWayangSpellhunterSlotAdjustment,
  TRANSMUTER_OF_KORADA_ID
} from "./traits/index.js";
import {
  getEldritchResearcherSlotAdjustment,
  getMaleficiumMinimumConsumedSlotLevel,
  getMaleficiumSlotAdjustment,
  getSpontaneousMetafocusStatus,
  getSpellPerfectionStatus,
  MASK_FOCUS_FEATURE_ID,
  MASK_FOCUS_ID
} from "./feats/index.js";
import { appendDamagePartOverrides, mapDamagePartFormulas } from "../utils/formulaUtils.js";
import {
  contextHasSavingThrow,
  ensureSpellSaveBaseDcSnapshot,
  resolveFeatSaveDcBase,
} from "../utils/saveDcUtils.js";

const SHORT_CAST_TYPES = new Set(["swift", "immediate", "move"]);
const STANDARD_CAST_TYPE = "standard";
const INSTANT_DURATION_UNITS = new Set(["inst", "instantaneous"]);
const CLASS_FEATURE_OPTIONS_KEY = "classFeatures";
const ARCANE_BLOODLINE_FEATURE_ID = "arcaneBloodline";
const METAMAGIC_ADEPT_FEATURE_ID = "metamagicAdept";
const ARCANE_APOTHEOSIS_FEATURE_ID = "arcaneApotheosis";
const GRAND_MAESTRO_FEATURE_ID = "grandMaestro";
const SEEKER_ETERNAL_EMPEROR_FEATURE_ID = "seekerOfTheEternalEmperor";
const RETRIBUTION_FEATURE_ID = "retribution";
const HEALERS_BLESSING_FEATURE_ID = "healersBlessing";
const INTENSE_CELEBRATION_FEATURE_ID = "intenseCelebration";
const NANITE_BLOODLINE_ARCANA_FEATURE_ID = "naniteBloodlineArcana";
const ONE_BODY_TWO_MINDS_FEATURE_ID = "oneBodyTwoMinds";
const PEERLESS_SPEED_FEATURE_ID = "peerlessSpeed";
const TIMELESS_SOUL_FEATURE_ID = "timelessSoul";
const SUCCOR_FINAL_REVELATION_FEATURE_ID = "succorFinalRevelation";
const SPELL_PERFECTION_FEATURE_ID = "spellPerfection";
const SPONTANEOUS_METAFOCUS_FEATURE_ID = "spontaneousMetafocus";
const CURATOR_MYSTIC_SECRETS_ID = "curatorOfMysticSecrets";
const SUCCOR_ELIGIBLE_METAMAGIC = new Set(["Enlarge Spell", "Extend Spell", "Silent Spell", "Still Spell"]);

function localizeMetamagic(path) {
  return game.i18n.localize(`NAS.metamagic.${path}`);
}

function formatMetamagic(path, data = {}) {
  return game.i18n.format(`NAS.metamagic.${path}`, data);
}

function canApplySelectiveSpell(context) {
  if (!context?.area?.isAreaOfEffect) return false;
  const units = (context.duration?.units ?? "").toString().toLowerCase();
  if (!INSTANT_DURATION_UNITS.has(units)) return false;
  return true;
}

function getSelectiveMaxExclusions(action) {
  const spellbook = action?.item?.system?.spellbook;
  const mod = action?.shared?.rollData?.attributes?.spells?.spellbooks?.[spellbook]?.abilityMod;
  const value = Number(mod ?? 0);
  return Number.isFinite(value) ? Math.max(0, value) : 0;
}

function getTargetId(target) {
  const token = target?.object ?? target?.token ?? target;
  return (
    token?.document?.uuid ??
    token?.uuid ??
    token?.id ??
    target?.uuid ??
    target?.id ??
    null
  );
}

function getTargetLabel(target, fallback) {
  const token = target?.object ?? target?.token ?? target;
  return token?.name ?? token?.actor?.name ?? target?.name ?? fallback ?? game.i18n.localize("NAS.common.labels.target");
}

function normalizeDescriptorToken(value) {
  return (value ?? "").toString().toLowerCase().replace(/[^a-z0-9]/g, "");
}

function descriptorLooksMindAffecting(value) {
  const normalized = normalizeDescriptorToken(value);
  return normalized === "mindaffecting";
}

function collectDescriptorLabelStrings(descriptors) {
  const out = [];
  if (Array.isArray(descriptors?.names)) out.push(...descriptors.names);
  if (Array.isArray(descriptors?.base)) out.push(...descriptors.base);
  const value = descriptors?.value;
  if (typeof value === "string" && value.trim()) out.push(value);
  return out;
}

function spellItemHasMindAffectingDescriptor(item) {
  if (item?.type !== "spell") return false;
  const legacy = item?.system?.descriptor;
  if (typeof legacy === "string") {
    const parts = legacy.split(",");
    if (parts.some((entry) => descriptorLooksMindAffecting(entry))) return true;
  }

  const descriptors = item?.system?.descriptors;
  if (!descriptors || typeof descriptors !== "object") return false;
  for (const entry of collectDescriptorLabelStrings(descriptors)) {
    if (descriptorLooksMindAffecting(entry)) return true;
  }

  const total = descriptors.total;
  if (total instanceof Set) {
    for (const key of total) {
      if (descriptorLooksMindAffecting(key)) return true;
    }
  } else if (total && typeof total === "object" && !Array.isArray(total)) {
    for (const [key, enabled] of Object.entries(total)) {
      if (!enabled) continue;
      if (descriptorLooksMindAffecting(key)) return true;
    }
  }
  return false;
}

function isCasterOnlyTargetSelection(action) {
  const targets = Array.isArray(action?.shared?.targets) ? action.shared.targets : [];
  if (targets.length > 1) return false;
  if (targets.length === 0) return true;
  const casterId = getTargetId(action?.token);
  if (!casterId) return false;
  return getTargetId(targets[0]) === casterId;
}

function isDivinationSpell(item) {
  const school = (item?.system?.school ?? "").toString().trim().toLowerCase();
  return school === "div" || school === "divination";
}

function promptSelectiveExclusions(targets, maxSelections) {
  return new Promise((resolve) => {
    const cards = targets
      .map((entry, index) => {
        const token = entry.target?.object ?? entry.target?.token ?? entry.target;
        const tokenImg = token?.document?.texture?.src || token?.texture?.src || "";
        const tokenName = entry.label || token?.name || token?.actor?.name || `Target ${index + 1}`;
        const disposition = token?.document?.disposition ?? token?.disposition;
        const isSameDisposition = disposition === entry?.actionTokenDisposition;
        const borderColor = isSameDisposition ? "green" : "red";
        const tag = isSameDisposition
          ? game.i18n.localize("NAS.common.labels.ally")
          : game.i18n.localize("NAS.common.labels.foe");
        return `
          <div class="selective-target-card" style="display: flex; flex-direction: column; align-items: center; width: 170px; border: 1px solid #ccc; border-radius: 6px; padding: 6px;">
            <div style="align-self: flex-start; margin-bottom: 4px;">
              <input type="checkbox" name="selectiveTarget" value="${entry.id}" />
            </div>
            <img src="${tokenImg}" style="width: 64px; height: 64px; border: 2px solid ${borderColor}; border-radius: 5px;" />
            <label style="margin: 4px 0; font-weight: 600; text-align: center;">${tokenName}</label>
            <small style="color: ${borderColor};">${tag}</small>
          </div>
        `;
      })
      .join("");
    const content = `
      <form>
        <p style="margin: 0 0 8px 0;">${foundry.utils.escapeHTML(localizeMetamagic("selectiveSpell.excludeTargetsLabel"))} <b><span class="nas-selective-count">0</span>/${maxSelections}</b></p>
        <div class="nas-selective-options" style="max-height: 340px; overflow-y: auto; border: 1px solid #ccc; border-radius: 6px; padding: 8px;">
          <div style="display: flex; flex-wrap: wrap; gap: 12px;">
            ${cards}
          </div>
        </div>
      </form>
    `;
    new Dialog({
      title: localizeMetamagic("selectiveSpell.title"),
      content,
      buttons: {
        ok: {
          label: game.i18n.localize("NAS.common.buttons.apply"),
          callback: (html) => {
            const checked = html.find('input[name="selectiveTarget"]:checked').toArray();
            resolve(checked.map((input) => input.value));
          },
        },
        cancel: {
          label: game.i18n.localize("NAS.common.buttons.cancel"),
          callback: () => resolve(null),
        },
      },
      default: "ok",
      close: () => resolve(null),
      render: (html) => {
        const app = html.closest('.app');
        if (app?.length) {
          app.css('width', '800px');
        }
        const dlg = ui.windows?.[Object.keys(ui.windows).find((id) => ui.windows[id]?.element?.[0] === app?.[0])];
        if (dlg?.setPosition) {
          dlg.setPosition({ width: 800 });
        }
        const updateCount = () => {
          const count = html.find('input[name="selectiveTarget"]:checked').length;
          html.find('.nas-selective-count').text(count);
        };
        html.find('input[name="selectiveTarget"]').on('change', (event) => {
          const checked = html.find('input[name="selectiveTarget"]:checked').length;
          if (checked > maxSelections) {
            event.currentTarget.checked = false;
          }
          updateCount();
        });
        updateCount();
      },
    }).render(true);
  });
}

function resolveTargetToken(target) {
  return target?.object ?? target?.token ?? target ?? null;
}

function isAllyOrSelfTarget(action, target) {
  const casterToken = action?.token ?? null;
  const casterDisposition = casterToken?.document?.disposition ?? casterToken?.disposition ?? null;
  const casterActorId = (action?.actor ?? casterToken?.actor ?? null)?.id ?? null;
  const token = resolveTargetToken(target);
  if (!token) return false;
  const targetActorId = token?.actor?.id ?? null;
  if (casterActorId && targetActorId && casterActorId === targetActorId) return true;
  const disposition = token?.document?.disposition ?? token?.disposition ?? null;
  if (casterDisposition === null || casterDisposition === undefined) return false;
  return disposition === casterDisposition;
}

function promptSuccorFinalRevelationNonAllies(nonAllyEntries = []) {
  const items = nonAllyEntries
    .map((entry) => `<li>${foundry.utils.escapeHTML(entry?.label ?? game.i18n.localize("NAS.common.labels.target"))}</li>`)
    .join("");
  const content = `
    <form>
      <p style="margin:0 0 8px 0;">
        ${foundry.utils.escapeHTML(localizeMetamagic("succorFinalRevelation.nonAlliesPrompt"))}
      </p>
      <ul style="margin:0 0 8px 18px;max-height:160px;overflow:auto;">
        ${items}
      </ul>
    </form>
  `;
  return new Promise((resolve) => {
    new Dialog({
      title: game.i18n.localize("NAS.metamagic.featureNames.succorFinalRevelation"),
      content,
      buttons: {
        ignore: {
          label: localizeMetamagic("succorFinalRevelation.ignoreNonAllies"),
          callback: () => resolve("ignore")
        },
        reject: {
          label: localizeMetamagic("succorFinalRevelation.cancelAction"),
          callback: () => resolve("reject")
        },
        continue: {
          label: localizeMetamagic("succorFinalRevelation.continueAnyway"),
          callback: () => resolve("continue")
        }
      },
      default: "ignore",
      close: () => resolve(null)
    }).render(true);
  });
}

function getSpellbookData(action) {
  const actor = action?.token?.actor ?? action?.actor;
  const spellbook = action?.item?.system?.spellbook;
  if (!actor || !spellbook) return null;
  return actor.system?.attributes?.spells?.spellbooks?.[spellbook] ?? null;
}

function applyMetamagicCastTime(action, context) {
  if (!context?.activation) return;
  const spellbookData = getSpellbookData(action);
  if (!spellbookData?.spontaneous) return;
  const hasAppliedMetamagic = Array.isArray(context?.metamagic?.applied) && context.metamagic.applied.length > 0;
  const hasOneBodyTwoMindsSurcharge = context?.metamagic?.oneBodyTwoMindsCastTimeSurcharge === true;
  if (!hasAppliedMetamagic && !hasOneBodyTwoMindsSurcharge) return;
  if (context?.metamagic?.bypassCastTimeIncrease === true) return;
  if (context.metamagic.applied.includes(QUICKEN_SPELL_NAME)) return;

  const rule = game.settings.get(MODULE.ID, "metamagicCastTimeRule");
  const type = (context.activation.type ?? "").toString().toLowerCase();
  const costValue = Number(context.activation.cost ?? 1);
  const cost = Number.isFinite(costValue) && costValue > 0 ? costValue : 1;

  if (SHORT_CAST_TYPES.has(type) && rule === "standard") {
    return;
  }

  if (type === STANDARD_CAST_TYPE) {
    context.activation.type = "full";
    context.activation.cost = 1;
    if (context.activation.unchained) {
      context.activation.unchained.type = "full";
      context.activation.unchained.cost = 1;
    }
    return;
  }

  if (SHORT_CAST_TYPES.has(type)) {
    context.activation.type = "full";
    context.activation.cost = 1;
    if (context.activation.unchained) {
      context.activation.unchained.type = "full";
      context.activation.unchained.cost = 1;
    }
    return;
  }

  context.activationExtraFullRound = true;
}

function getClassFeatureOptions(context) {
  const options = context?.metamagicOptions?.[CLASS_FEATURE_OPTIONS_KEY];
  if (!options || typeof options !== "object") return {};
  return options;
}

function isClassFeatureSelected(context, featureId) {
  const features = getClassFeatureOptions(context);
  return features?.[featureId] === true;
}

function isClassFeatureEnabledOrDefault(context, featureId, defaultEnabled = false) {
  const features = getClassFeatureOptions(context);
  if (Object.prototype.hasOwnProperty.call(features, featureId)) {
    return features?.[featureId] === true;
  }
  return Boolean(defaultEnabled);
}

function hasAnySlotIncreasingMetamagic(context, baseSpellLevel) {
  const heightenLevel = Number(context?.metamagic?.heightenLevel ?? 0);
  const heightenIncrease = Number.isFinite(baseSpellLevel)
    ? Math.max(0, heightenLevel - baseSpellLevel)
    : 0;
  const slotIncrease = Number(
    context?.metamagic?.otherSlotIncrease
    ?? context?.metamagic?.slotIncrease
    ?? 0
  );
  const total = (Number.isFinite(slotIncrease) ? slotIncrease : 0) + heightenIncrease;
  return total > 0;
}

function resolveBaseSaveDc(action, context) {
  const contextDcRaw = context?.save?.dc;
  if (typeof contextDcRaw === "number" && Number.isFinite(contextDcRaw)) {
    return { value: contextDcRaw, source: "context.save.dc" };
  }
  if (typeof contextDcRaw === "string" && contextDcRaw.trim().length > 0) {
    const parsed = Number(contextDcRaw.trim());
    if (Number.isFinite(parsed)) {
      return { value: parsed, source: "context.save.dc" };
    }
  }

  const sharedDcRaw = action?.shared?.saveDC;
  if (typeof sharedDcRaw === "number" && Number.isFinite(sharedDcRaw)) {
    return { value: sharedDcRaw, source: "action.shared.saveDC" };
  }
  if (typeof sharedDcRaw === "string" && sharedDcRaw.trim().length > 0) {
    const parsed = Number(sharedDcRaw.trim());
    if (Number.isFinite(parsed)) {
      return { value: parsed, source: "action.shared.saveDC" };
    }
  }

  if (typeof action?.action?.getDC === "function" && action?.shared?.rollData) {
    const liveDcRaw = action.action.getDC(action.shared.rollData);
    const liveDc = Number(liveDcRaw);
    if (Number.isFinite(liveDc)) {
      return { value: liveDc, source: "action.action.getDC" };
    }
  }

  return { value: null, source: "unresolved" };
}

function applyArcaneBloodlineDcBonus(action, context) {
  if (!contextHasSavingThrow(context, action)) return;
  const baseDc = resolveBaseSaveDc(action, context);
  if (!Number.isFinite(baseDc?.value)) {
    return;
  }
  context.save ??= {};
  context.save.dc = baseDc.value + 1;

  const evaluatedTotal = Number(context?.save?.evaluated?.total);
  if (Number.isFinite(evaluatedTotal) && evaluatedTotal > 0) {
    context.save.evaluated.total = evaluatedTotal + 1;
  } else {
    context.save.evaluated ??= {};
    context.save.evaluated.total = baseDc.value + 1;
  }

  action.shared ??= {};
  action.shared.saveDC = baseDc.value + 1;

  context.metamagic ??= { applied: [], slotIncrease: 0 };
  context.metamagic.arcaneBloodlineDcBonus = 1;
}

function applyGrandMaestroComponents(context) {
  if (!context?.components) return false;
  if (context.components.verbal !== true) return false;
  if (context.components.somatic === true) {
    context.components.somatic = false;
    return true;
  }
  return false;
}

function applyHealersBlessingDamageBonus(context) {
  // TODO(NAS-HealersBlessing): Add branch-aware damage handling for spells with
  // multiple target-type outcomes (living/undead/construct/light-vulnerable).
  // Cure/Inflict and spells like Searing Light need separate formula branches.
  //
  // TODO(NAS-HealersBlessing): Add chat-card multi-section damage blocks with
  // independent apply buttons, then route final branch selection through actor/token
  // creature-type data before applying healing/damage.
  //
  // TODO(NAS-HealersBlessing): When branch-aware handling exists, enforce RAW
  // exclusion so this bonus does not empower cure-spell damage dealt to undead.
  const overrides = mapDamagePartFormulas(context, (formula) => applyEmpowerToFormula(formula));
  if (!appendDamagePartOverrides(context, overrides)) return false;
  return true;
}

function applyNaniteBloodlineArcanaDuration(context) {
  if (!context?.duration) return false;
  if (!isDurationEligibleForExtendSpell(context.duration)) return false;

  const units = (context.duration.units ?? "").toString();
  const evaluatedBase = Number(context.duration.evaluated?.total);
  let baseTotal = Number.isFinite(evaluatedBase) ? evaluatedBase : Number(context.duration.value ?? 0);
  if (!Number.isFinite(baseTotal) || baseTotal <= 0) return false;

  const extendedTotal = baseTotal * 2;
  context.duration.value = String(extendedTotal);
  context.duration.evaluated = {
    ...(context.duration.evaluated ?? {}),
    total: extendedTotal
  };
  context.duration.naniteBloodlineArcana = {
    units,
    baseTotal,
    extendedTotal
  };
  return true;
}

function applyOneBodyTwoMindsEffects(context, action, { allowMindAffectingExtend = true } = {}) {
  if (!context) return { applied: false, appliedExtend: false, mindAffecting: false };

  let touched = false;
  if (context?.components?.verbal === true) {
    context.components.verbal = false;
    touched = true;
  }
  if (context?.components?.somatic === true) {
    context.components.somatic = false;
    touched = true;
  }

  const spellItem = action?.item ?? context?.item ?? null;
  const mindAffecting = spellItemHasMindAffectingDescriptor(spellItem);
  let appliedExtend = false;
  if (allowMindAffectingExtend && mindAffecting) {
    const hasRealExtend = Array.isArray(context?.metamagic?.applied)
      && context.metamagic.applied.includes(ExtendSpell.name);
    if (!hasRealExtend) {
      appliedExtend = applyExtendDurationOnly(context);
    }
  }

  context.metamagic ??= { applied: [], slotIncrease: 0 };
  context.metamagic.oneBodyTwoMindsCastTimeSurcharge = true;
  // TODO(NAS-OneBodyTwoMinds): House rule currently keeps spontaneous metamagic cast-time
  // increase for this ability; revisit if table ruling changes.
  if (touched || appliedExtend) {
    context.metamagic.oneBodyTwoMindsApplied = true;
  }

  return {
    applied: touched || appliedExtend,
    appliedExtend,
    mindAffecting
  };
}

function validateLimitedFeatureUse(item) {
  const uses = item?.system?.uses;
  if (!uses || !uses.per) {
    return { ok: false, reason: "missingUsesData" };
  }
  const remaining = Number(uses.value ?? 0);
  if (!Number.isFinite(remaining) || remaining <= 0) {
    return { ok: false, reason: "noRemainingUses" };
  }
  return { ok: true };
}

function validateTraitSourceUse(source) {
  if (!source?.limited) return { ok: true };
  if (source?.hasUsesData !== true) return { ok: false, reason: "missingUsesData" };
  if (source?.hasRemaining !== true) return { ok: false, reason: "noRemainingUses" };
  return { ok: true };
}

function getMimicMetamagicSelection(context) {
  const raw = context?.metamagicOptions?.mimicMetamagic;
  if (!raw || typeof raw !== "object") return null;
  if (raw?.enabled !== true) return null;
  const chosenMetaName = (raw?.chosenMetaName ?? "").toString().trim();
  if (!chosenMetaName) return null;
  const heightenSpellLevel = Number(raw?.heightenSpellLevel ?? 0);
  return {
    chosenMetaName,
    heightenSpellLevel: Number.isFinite(heightenSpellLevel) ? heightenSpellLevel : 0
  };
}

function getPeerlessSpeedSelection(context) {
  const raw = context?.metamagicOptions?.peerlessSpeed;
  if (!raw || typeof raw !== "object") return null;
  if (raw?.enabled !== true) return null;
  const chosenMetaName = (raw?.chosenMetaName ?? "").toString().trim();
  if (!chosenMetaName) return null;
  return { chosenMetaName };
}

function getSuccorFinalRevelationSelection(context) {
  const raw = context?.metamagicOptions?.succorFinalRevelation;
  if (!raw || typeof raw !== "object") return null;
  if (raw?.enabled !== true) return null;
  const chosenMetaName = (raw?.chosenMetaName ?? "").toString().trim();
  if (!chosenMetaName) return null;
  return { chosenMetaName };
}

function isTimelessSoulEnabled(context) {
  if (typeof context?.metamagic?.timelessSoulActive === "boolean") {
    return context.metamagic.timelessSoulActive;
  }
  const selected = isClassFeatureSelected(context, TIMELESS_SOUL_FEATURE_ID);
  if (!selected) return false;
  const applied = Array.isArray(context?.metamagic?.applied) ? context.metamagic.applied : [];
  if (!applied.includes(QuickenSpell.name)) return false;
  const peerlessQuickened = context?.metamagic?.peerlessSpeedApplied === true
    && context?.metamagic?.peerlessSpeedMetaName === QuickenSpell.name;
  return !peerlessQuickened;
}

function getNominalMetamagicLevelIncrease(metaName, context, baseSpellLevel, mimicSelection = null) {
  switch ((metaName ?? "").toString()) {
    case "Still Spell":
    case "Silent Spell":
    case "Enlarge Spell":
    case "Extend Spell":
    case "Reach Spell":
    case "Selective Spell":
    case "Intensified Spell":
      return 1;
    case "Persistent Spell":
    case "Empower Spell":
      return 2;
    case "Dazing Spell":
    case "Maximize Spell":
      return 3;
    case "Quicken Spell":
      return isTimelessSoulEnabled(context) ? 3 : 4;
    case "Heighten Spell": {
      const targetFromMimic = Number(mimicSelection?.heightenSpellLevel ?? 0);
      const targetFromReal = Number(context?.metamagicOptions?.heightenSpellLevel ?? 0);
      const target = Number.isFinite(targetFromMimic) && targetFromMimic > 0
        ? targetFromMimic
        : targetFromReal;
      const base = Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0;
      return Math.max(0, target - base);
    }
    default:
      return 0;
  }
}

function computeConsumedSlotIncreaseForMetamagic(context, baseSpellLevel, hasAppliedMetamagic) {
  const otherIncrease = Number(context?.metamagic?.slotIncrease ?? 0);
  const heightenLevel = Number(context?.metamagic?.heightenLevel ?? 0);
  const heightenDelta = Number.isFinite(baseSpellLevel)
    ? Math.max(0, heightenLevel - baseSpellLevel)
    : 0;
  const normalizedOtherIncrease = Number.isFinite(otherIncrease) ? otherIncrease : 0;
  const timelessSoulActive = isTimelessSoulEnabled(context);
  const quickenNominalIncrease = timelessSoulActive ? 4 : 0;
  const quickenBaseIncrease = Math.min(
    Math.max(0, normalizedOtherIncrease),
    Math.max(0, quickenNominalIncrease)
  );
  const timelessSoulReduction = timelessSoulActive ? Math.min(1, quickenBaseIncrease) : 0;
  const nonQuickenIncrease = timelessSoulActive
    ? Math.max(0, normalizedOtherIncrease - quickenBaseIncrease)
    : normalizedOtherIncrease;
  const reducerInputIncrease = timelessSoulActive
    ? nonQuickenIncrease
    : normalizedOtherIncrease;
  const hasSlotSurchargeMetamagic = reducerInputIncrease > 0;
  const slotAdjustmentRaw =
    getEldritchResearcherSlotAdjustment(context, { hasAppliedMetamagic: hasSlotSurchargeMetamagic }) +
    getMagicalLineageSlotAdjustment(context, { hasAppliedMetamagic: hasSlotSurchargeMetamagic }) +
    getWayangSpellhunterSlotAdjustment(context, {
      hasAppliedMetamagic: hasSlotSurchargeMetamagic,
      timelessSoulActive
    }) +
    getMaleficiumSlotAdjustment(context, { hasAppliedMetamagic: hasSlotSurchargeMetamagic }) +
    getSuccorFinalRevelationSlotAdjustment(context, { hasAppliedMetamagic: hasSlotSurchargeMetamagic }) +
    getRetributionSlotAdjustment(context, { hasAppliedMetamagic });
  const slotAdjustment = Number.isFinite(slotAdjustmentRaw) ? slotAdjustmentRaw : 0;
  const reducedOtherSlotIncrease = timelessSoulActive
    ? (quickenBaseIncrease - timelessSoulReduction) + Math.max(0, nonQuickenIncrease + slotAdjustment)
    : Math.max(0, normalizedOtherIncrease + slotAdjustment);
  const preWaiverConsumedSlotIncrease = (Number.isFinite(heightenDelta) ? heightenDelta : 0) + reducedOtherSlotIncrease;
  const spellPerfectionWaiver = getSpellPerfectionMetamagicWaiver(context, baseSpellLevel, { hasAppliedMetamagic });
  let consumedSlotIncrease = Math.max(0, preWaiverConsumedSlotIncrease - spellPerfectionWaiver);
  const minimumConsumedSlotLevel = getMaleficiumMinimumConsumedSlotLevel(context, { hasAppliedMetamagic });
  const baseLevel = Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0;
  if (Number.isFinite(minimumConsumedSlotLevel)) {
    const consumedSlotLevel = baseLevel + consumedSlotIncrease;
    if (consumedSlotLevel < minimumConsumedSlotLevel) {
      consumedSlotIncrease = Math.max(0, minimumConsumedSlotLevel - baseLevel);
    }
  }
  return {
    normalizedOtherIncrease,
    heightenDelta,
    slotAdjustment,
    timelessSoulReduction,
    timelessSoulActive,
    reducedOtherSlotIncrease,
    preWaiverConsumedSlotIncrease,
    spellPerfectionWaiver,
    consumedSlotIncrease
  };
}

function getSpellPerfectionMetamagicWaiver(context, baseSpellLevel, { hasAppliedMetamagic = false } = {}) {
  if (!hasAppliedMetamagic) return 0;
  const status = getSpellPerfectionStatus(context);
  if (!status.enabled) return 0;
  const applied = Array.isArray(context?.metamagic?.applied) ? context.metamagic.applied : [];
  if (!applied.length) return 0;
  return applied.reduce((max, metaName) => {
    if ((metaName ?? "").toString() === QuickenSpell.name && isTimelessSoulEnabled(context)) {
      return max;
    }
    const increase = Number(getNominalMetamagicLevelIncrease(metaName, context, baseSpellLevel, null));
    if (!Number.isFinite(increase) || increase <= 0) return max;
    return Math.max(max, increase);
  }, 0);
}

function getRetributionSlotAdjustment(context, { hasAppliedMetamagic = false } = {}) {
  if (!hasAppliedMetamagic) return 0;
  const enabled = isClassFeatureEnabledOrDefault(context, RETRIBUTION_FEATURE_ID, false);
  return enabled ? -1 : 0;
}

function getSuccorFinalRevelationSlotAdjustment(context, { hasAppliedMetamagic = false } = {}) {
  if (!hasAppliedMetamagic) return 0;
  const raw = Number(context?.metamagic?.succorFinalRevelationSlotAdjustment ?? 0);
  if (!Number.isFinite(raw)) return 0;
  return raw;
}

function getSelectedTraitIds(context) {
  if (!Array.isArray(context?.traitOptions)) return [];
  return context.traitOptions.map((value) => `${value}`).filter(Boolean);
}

function appendPendingTraitUse(action, entry) {
  if (!entry || !entry.itemUuid) return;
  action.shared ??= {};
  const current = Array.isArray(action.shared.nasPendingTraitUses) ? action.shared.nasPendingTraitUses : [];
  current.push(entry);
  action.shared.nasPendingTraitUses = current;
}

function applyTraitDurationExtension(context) {
  if (!context?.duration) return false;
  if (!isDurationEligibleForExtendSpell(context.duration)) return false;

  const evaluatedBase = Number(context.duration.evaluated?.total);
  if (Number.isFinite(evaluatedBase)) {
    const total = evaluatedBase * 2;
    context.duration.value = String(total);
    context.duration.evaluated = {
      ...(context.duration.evaluated ?? {}),
      total
    };
    return true;
  }

  const rawValue = Number(context.duration.value ?? 0);
  if (!Number.isFinite(rawValue) || rawValue <= 0) return false;
  const doubled = rawValue * 2;
  context.duration.value = String(doubled);
  context.duration.evaluated = {
    ...(context.duration.evaluated ?? {}),
    total: doubled
  };
  return true;
}

async function recomputeDurationEvaluationFromRollData(action, context) {
  const duration = context?.duration;
  const formula = (duration?.value ?? "").toString().trim();
  if (!formula) return false;
  const rollData = action?.shared?.rollData ?? {};
  try {
    const total = Roll?.defaultImplementation?.safeRollSync?.(formula, rollData)?.total;
    if (!Number.isFinite(total)) return false;
    duration.evaluated = {
      ...(duration.evaluated ?? {}),
      total: Number(total)
    };
    return true;
  } catch (_error) {
    return false;
  }
}

async function applyRacialTraitSelections(action, context, traitSources = null) {
  const selectedIds = getSelectedTraitIds(context);
  if (!selectedIds.length) return;
  if (context?.metamagic?.transmuterOfKoradaActive === true) return;

  const availableTraitSources = Array.isArray(traitSources)
    ? traitSources
    : await getRacialSpellLikeTraitSources(context?.actor ?? action?.actor, action?.item ?? null);
  if (!Array.isArray(availableTraitSources) || !availableTraitSources.length) return;

  const sourceById = new Map(availableTraitSources.map((source) => [source.id, source]));
  const selectedSources = selectedIds
    .map((id) => sourceById.get(id))
    .filter((source) => source?.effectType === "durationExtension");
  if (!selectedSources.length) return;

  const hasExtendSpellAlready = Array.isArray(context?.metamagic?.applied)
    && context.metamagic.applied.includes(ExtendSpell.name);
  if (hasExtendSpellAlready) return;

  const activeLabels = [];
  const pending = [];

  for (const source of selectedSources) {
    if (source.limited) {
      if (source.hasUsesData !== true || source.hasRemaining !== true) {
        ui.notifications.warn(`${source.label} has no remaining uses.`);
        continue;
      }
    }

    const applied = applyTraitDurationExtension(context);
    if (!applied) continue;

    activeLabels.push(source.label);
    if (source.hasUsesData === true) {
      pending.push({
        itemUuid: source.itemUuid ?? null,
        mode: "use"
      });
    } else {
      pending.push({
        itemUuid: source.itemUuid ?? null,
        mode: "displayCard"
      });
    }
    break;
  }

  if (!activeLabels.length) return;
  context.metamagic ??= { applied: [], slotIncrease: 0 };
  const current = Array.isArray(context.metamagic.activeFeatureLabels) ? context.metamagic.activeFeatureLabels : [];
  context.metamagic.activeFeatureLabels = Array.from(new Set([...current, ...activeLabels]));

  pending.forEach((entry) => appendPendingTraitUse(action, entry));
}

function getSpellSlotData(action, slotIncrease) {
  const actor = action.token?.actor ?? action.actor;
  const spellbook = action.item?.system?.spellbook;
  const baseLevel = action.item?.system?.level ?? action.shared?.rollData?.sl ?? 0;
  const targetLevel = baseLevel + slotIncrease;

  if (!actor || !spellbook || targetLevel <= 0) {
    return null;
  }

  const spellbookData = actor.system?.attributes?.spells?.spellbooks?.[spellbook];
  const spellLevelKey = `spell${targetLevel}`;
  const spellLevelData = spellbookData?.spells?.[spellLevelKey];

  if (!spellbookData || !spellLevelData) {
    return null;
  }

  return {
    actor,
    spellbook,
    spellbookData,
    spellLevelKey,
    spellLevelData,
    targetLevel,
  };
}

async function consumeHigherSpellSlot(action, slotIncrease) {
  if (!action?.item || slotIncrease <= 0) return { skipped: true };
  if (action.item.type !== "spell") return { skipped: true };

  const slotData = getSpellSlotData(action, slotIncrease);
  if (!slotData) {
    action.shared.reject = true;
    ui.notifications.warn(
      game.i18n.format("NAS.buffs.NotEnoughSpellSlots", {
        remaining: 0,
        needed: 1,
      })
    );
    return { rejected: true };
  }

  if (!slotData.spellbookData?.spontaneous) return { skipped: true, prepared: true };

  const remainingSlots = Number(slotData.spellLevelData.value ?? 0);
  if (remainingSlots < 1) {
    action.shared.reject = true;
    ui.notifications.warn(
      game.i18n.format("NAS.buffs.NotEnoughSpellSlots", {
        remaining: remainingSlots,
        needed: 1,
      })
    );
    return { rejected: true };
  }

  const updatePath = `system.attributes.spells.spellbooks.${slotData.spellbook}.spells.${slotData.spellLevelKey}.value`;
  await slotData.actor.update({ [updatePath]: remainingSlots - 1 });

  if (typeof action.shared?.rollData?.chargeCost === "number") {
    action.shared.rollData.chargeCost = 0;
  }
  // We already consumed the higher-level slot manually above.
  // Prevent PF1's default charge pipeline from also deducting charges from the base spell level.
  action.shared.cost = 0;
  action.shared.rollData.chargeCost = 0;
  action.shared.rollData.chargeCostBonus = 0;

  return {
    spellbook: slotData.spellbook,
    spellLevelKey: slotData.spellLevelKey,
    targetLevel: slotData.targetLevel,
    consumed: true,
  };
}

export async function applyMetamagicSelections(action, context) {
  if (action?.shared && Object.prototype.hasOwnProperty.call(action.shared, "nasPendingMetamagicUse")) {
    delete action.shared.nasPendingMetamagicUse;
  }
  if (action?.shared && Object.prototype.hasOwnProperty.call(action.shared, "nasPendingTraitUses")) {
    delete action.shared.nasPendingTraitUses;
  }
  if (action?.shared && Object.prototype.hasOwnProperty.call(action.shared, "nasPendingMaskFocusUse")) {
    delete action.shared.nasPendingMaskFocusUse;
  }
  if (action?.shared && Object.prototype.hasOwnProperty.call(action.shared, "nasPendingArcaneReservoirSpend")) {
    delete action.shared.nasPendingArcaneReservoirSpend;
  }
  if (action?.shared && Object.prototype.hasOwnProperty.call(action.shared, "nasPendingPhrenicPoolSpend")) {
    delete action.shared.nasPendingPhrenicPoolSpend;
  }
  const [
    traitSources,
    healersBlessingSources,
    intenseCelebrationSources,
    naniteBloodlineArcanaSources,
    oneBodyTwoMindsSources,
    oracleSeekerSources,
    oracleSuccorSources,
    peerlessSpeedSources,
    timelessSoulSources,
    transmuterStandaloneSource
  ] = await Promise.all([
    getRacialSpellLikeTraitSources(context?.actor ?? action?.actor, action?.item ?? null),
    getHealersBlessingFeatureSources(context?.actor ?? action?.actor, action?.item ?? null),
    getIntenseCelebrationFeatureSources(context?.actor ?? action?.actor, action?.item ?? null),
    getNaniteBloodlineArcanaFeatureSources(context?.actor ?? action?.actor, action?.item ?? null, {
      durationOverride: context?.duration ?? null
    }),
    getOneBodyTwoMindsFeatureSources(context?.actor ?? action?.actor, action?.item ?? null),
    getOracleSeekerFeatureSources(context?.actor ?? action?.actor, action?.item ?? null),
    getOracleSuccorFinalRevelationFeatureSources(context?.actor ?? action?.actor, action?.item ?? null),
    getPeerlessSpeedFeatureSources(context?.actor ?? action?.actor, action?.item ?? null),
    getTimelessSoulFeatureSources(context?.actor ?? action?.actor, action?.item ?? null),
    getTransmuterOfKoradaSource(context?.actor ?? action?.actor, action?.item ?? null)
  ]);
  const allTraitSources = [
    ...(Array.isArray(traitSources) ? traitSources : []),
    ...(transmuterStandaloneSource ? [transmuterStandaloneSource] : [])
  ];
  const healersBlessingSource = Array.isArray(healersBlessingSources) ? (healersBlessingSources[0] ?? null) : null;
  const intenseCelebrationSource = Array.isArray(intenseCelebrationSources) ? (intenseCelebrationSources[0] ?? null) : null;
  const naniteBloodlineArcanaSource = Array.isArray(naniteBloodlineArcanaSources) ? (naniteBloodlineArcanaSources[0] ?? null) : null;
  const oneBodyTwoMindsSource = Array.isArray(oneBodyTwoMindsSources) ? (oneBodyTwoMindsSources[0] ?? null) : null;
  const oracleSeekerSource = Array.isArray(oracleSeekerSources) ? (oracleSeekerSources[0] ?? null) : null;
  const oracleSuccorSource = Array.isArray(oracleSuccorSources) ? (oracleSuccorSources[0] ?? null) : null;
  const peerlessSpeedSource = Array.isArray(peerlessSpeedSources) ? (peerlessSpeedSources[0] ?? null) : null;
  const timelessSoulSource = Array.isArray(timelessSoulSources) ? (timelessSoulSources[0] ?? null) : null;
  const selectedTraitIds = new Set(getSelectedTraitIds(context));
  const transmuterSource = Array.isArray(allTraitSources)
    ? (allTraitSources.find((source) => source?.id === TRANSMUTER_OF_KORADA_ID) ?? null)
    : null;
  const transmuterCanApplyExtendLike =
    Boolean(transmuterSource)
    && transmuterSource?.selectedSpellMatches === true
    && transmuterSource?.requiresSpellSelection !== true;
  const transmuterRequestedForExtendLike =
    selectedTraitIds.has(TRANSMUTER_OF_KORADA_ID)
    && transmuterCanApplyExtendLike;
  action.shared ??= {};
  action.shared.rollData ??= {};
  const passiveTransmuterCasterLevelBonus = Number(transmuterSource?.casterLevelBonus ?? 0);
  let transmuterCasterLevelBonusApplied = false;
  if (Number.isFinite(passiveTransmuterCasterLevelBonus) && passiveTransmuterCasterLevelBonus > 0) {
    const currentCl = Number(action.shared.rollData.cl ?? 0);
    if (Number.isFinite(currentCl)) {
      action.shared.rollData.cl = currentCl + passiveTransmuterCasterLevelBonus;
      transmuterCasterLevelBonusApplied = true;
      await recomputeDurationEvaluationFromRollData(action, context);
    }
  }
  let transmuterUsed = false;
  if (transmuterRequestedForExtendLike) {
    const useValidation = validateTraitSourceUse(transmuterSource);
    if (!useValidation.ok) {
      if (useValidation.reason === "missingUsesData") {
        ui.notifications.warn(`${transmuterSource?.label ?? localizeMetamagic("featureNames.transmuterOfKorada")} has no uses metadata.`);
      } else {
        ui.notifications.warn(`${transmuterSource?.label ?? localizeMetamagic("featureNames.transmuterOfKorada")} has no remaining uses.`);
      }
    } else {
      transmuterUsed = true;
      appendPendingTraitUse(action, {
        itemUuid: transmuterSource?.itemUuid ?? null,
        mode: "use"
      });
    }
  }
  if (transmuterUsed) {
    applyExtendDurationOnly(context);
  }
  context.metamagic ??= { applied: [], slotIncrease: 0 };
  context.metamagic.transmuterOfKoradaActive = transmuterUsed;
  context.metamagic.transmuterOfKoradaCasterLevelBonusApplied = transmuterCasterLevelBonusApplied;
  const shouldApplyHealersBlessing =
    Boolean(healersBlessingSource) &&
    isClassFeatureEnabledOrDefault(context, HEALERS_BLESSING_FEATURE_ID, true);
  const shouldApplyIntenseCelebration =
    Boolean(intenseCelebrationSource) &&
    isClassFeatureEnabledOrDefault(context, INTENSE_CELEBRATION_FEATURE_ID, true);
  const shouldApplyNaniteBloodlineArcana =
    Boolean(naniteBloodlineArcanaSource) &&
    isClassFeatureEnabledOrDefault(context, NANITE_BLOODLINE_ARCANA_FEATURE_ID, true);
  const shouldApplyOneBodyTwoMinds =
    Boolean(oneBodyTwoMindsSource) &&
    isClassFeatureEnabledOrDefault(context, ONE_BODY_TWO_MINDS_FEATURE_ID, false);
  const shouldApplyPeerlessSpeed =
    Boolean(peerlessSpeedSource) &&
    isClassFeatureEnabledOrDefault(context, PEERLESS_SPEED_FEATURE_ID, false);
  const shouldApplyTimelessSoul =
    Boolean(timelessSoulSource) &&
    isClassFeatureEnabledOrDefault(context, TIMELESS_SOUL_FEATURE_ID, false);
  const shouldApplyOracleSeeker =
    Boolean(oracleSeekerSource) &&
    isClassFeatureEnabledOrDefault(context, SEEKER_ETERNAL_EMPEROR_FEATURE_ID, false);
  const shouldApplyOracleSuccor =
    Boolean(oracleSuccorSource) &&
    isClassFeatureEnabledOrDefault(context, SUCCOR_FINAL_REVELATION_FEATURE_ID, false);
  const healerBlessingApplied = shouldApplyHealersBlessing && applyHealersBlessingDamageBonus(context);
  const intenseCelebrationApplied =
    !transmuterUsed && shouldApplyIntenseCelebration && applyExtendDurationOnly(context);
  const naniteArcanaRequested = shouldApplyNaniteBloodlineArcana && isCasterOnlyTargetSelection(action);
  let naniteBloodlineArcanaApplied = false;
  let oneBodyTwoMindsUsed = false;
  if (shouldApplyOneBodyTwoMinds) {
    const useValidation = validateTraitSourceUse(oneBodyTwoMindsSource);
    if (!useValidation.ok) {
      if (useValidation.reason === "missingUsesData") {
        ui.notifications.warn(`${oneBodyTwoMindsSource?.label ?? localizeMetamagic("featureNames.oneBodyTwoMinds")} has no uses metadata.`);
      } else {
        ui.notifications.warn(`${oneBodyTwoMindsSource?.label ?? localizeMetamagic("featureNames.oneBodyTwoMinds")} has no remaining uses.`);
      }
    } else {
      applyOneBodyTwoMindsEffects(context, action, { allowMindAffectingExtend: !transmuterUsed });
      oneBodyTwoMindsUsed = true;
      appendPendingTraitUse(action, {
        itemUuid: oneBodyTwoMindsSource?.itemUuid ?? null,
        mode: "use"
      });
    }
  }
  const selections = Array.isArray(context?.metamagicNames) ? context.metamagicNames : [];
  const selectionsWithoutExtend = transmuterUsed
    ? selections
      .map((name) => resolveMetamagicNameFromDatabase(name) ?? name)
      .filter((name) => (name ?? "").toString().trim() !== ExtendSpell.name)
    : selections;
  const mimicSelection = getMimicMetamagicSelection(context);
  const peerlessSelection = getPeerlessSpeedSelection(context);
  const succorSelection = shouldApplyOracleSuccor ? getSuccorFinalRevelationSelection(context) : null;
  const eldritchResearcherEffect = context?.featEffects?.eldritchResearcher;
  const magicalLineageEffect = context?.featEffects?.magicalLineage;
  if (!selectionsWithoutExtend.length && !mimicSelection && !peerlessSelection && !succorSelection) {
    if (naniteArcanaRequested && !transmuterUsed) {
      naniteBloodlineArcanaApplied = applyNaniteBloodlineArcanaDuration(context);
    }
    if (
      healerBlessingApplied
      || intenseCelebrationApplied
      || naniteBloodlineArcanaApplied
      || oneBodyTwoMindsUsed
      || transmuterUsed
      || eldritchResearcherEffect?.active
    ) {
      context.metamagic ??= { applied: [], slotIncrease: 0 };
      const currentLabels = Array.isArray(context.metamagic.activeFeatureLabels) ? context.metamagic.activeFeatureLabels : [];
      const nextLabels = [...currentLabels];
      if (healerBlessingApplied) {
        nextLabels.push(healersBlessingSource?.label ?? localizeMetamagic("featureNames.healersBlessing"));
      }
      if (intenseCelebrationApplied) {
        nextLabels.push(intenseCelebrationSource?.label ?? localizeMetamagic("featureNames.intenseCelebration"));
      }
      if (naniteBloodlineArcanaApplied) {
        nextLabels.push(naniteBloodlineArcanaSource?.label ?? localizeMetamagic("featureNames.naniteBloodlineArcana"));
      }
      if (oneBodyTwoMindsUsed) {
        nextLabels.push(oneBodyTwoMindsSource?.label ?? localizeMetamagic("featureNames.oneBodyTwoMinds"));
      }
      if (transmuterUsed) {
        nextLabels.push(transmuterSource?.label ?? localizeMetamagic("featureNames.transmuterOfKorada"));
      }
      if (eldritchResearcherEffect?.active) {
        nextLabels.push(eldritchResearcherEffect.label ?? localizeMetamagic("featureNames.eldritchResearcher"));
      }
      context.metamagic.activeFeatureLabels = Array.from(new Set(nextLabels));
    }
    ensureSpellSaveBaseDcSnapshot(action, context);
    await applyRacialTraitSelections(action, context, allTraitSources);
    return;
  }

  const normalizedRealSelections = selectionsWithoutExtend
    .map((name) => resolveMetamagicNameFromDatabase(name) ?? name)
    .filter(Boolean);

  const baseSpellLevelRaw = action.item?.system?.level ?? action.shared?.rollData?.sl ?? 0;
  const baseSpellLevel = Number(baseSpellLevelRaw ?? 0);

  let mimicMetaName = "";
  let mimicMetaIncrease = 0;
  let mimicMetaDebitCount = 0;
  let mimicState = null;
  let mimicEnabled = false;
  if (mimicSelection) {
    mimicMetaName = (resolveMetamagicNameFromDatabase(mimicSelection.chosenMetaName) ?? mimicSelection.chosenMetaName ?? "").toString().trim();
    if (!mimicMetaName) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.missingChoice"));
      return;
    }
    if (normalizedRealSelections.includes(mimicMetaName)) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.duplicateWithReal"));
      return;
    }
    mimicState = await getPsychicMimicMetamagicState(context?.actor ?? action?.actor, action?.item ?? null);
    if (!mimicState?.eligible || !mimicState?.mimicMetamagicItem || !mimicState?.phrenicPoolItem) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.notEligible"));
      return;
    }
    const configured = Array.isArray(mimicState?.config?.configuredChoices)
      ? mimicState.config.configuredChoices.map((name) => resolveMetamagicNameFromDatabase(name) ?? name)
      : [];
    if (!configured.includes(mimicMetaName)) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.choiceNotConfigured"));
      return;
    }
    mimicMetaIncrease = getNominalMetamagicLevelIncrease(mimicMetaName, context, baseSpellLevel, mimicSelection);
    const spellbookKey = (action?.item?.system?.spellbook ?? "").toString();
    const maxCastable = getPsychicSpellbookMaxCastableSpellLevel(context?.actor ?? action?.actor, spellbookKey);
    const safeBase = Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0;
    if (maxCastable <= 0 || safeBase + mimicMetaIncrease > maxCastable) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.maxLevelExceeded"));
      return;
    }
    if (!mimicState.hasUsesData) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.missingPhrenicUsesData"));
      return;
    }
    mimicMetaDebitCount = Math.max(2, mimicMetaIncrease * 2);
    if (mimicState.usesRemaining < mimicMetaDebitCount) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.mimicMetamagic.insufficientPhrenicPoints"));
      return;
    }
    mimicEnabled = true;
  }

  let peerlessMetaName = "";
  let peerlessEnabled = false;
  if (peerlessSelection && !shouldApplyPeerlessSpeed) {
    action.shared.reject = true;
    ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.notEligible"));
    return;
  }
  if (peerlessSelection && shouldApplyPeerlessSpeed) {
    peerlessMetaName = (resolveMetamagicNameFromDatabase(peerlessSelection.chosenMetaName) ?? peerlessSelection.chosenMetaName ?? "").toString().trim();
    if (!peerlessMetaName) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.missingChoice"));
      return;
    }
    if (![QuickenSpell.name, EmpowerSpell.name, MaximizeSpell.name].includes(peerlessMetaName)) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.invalidChoice"));
      return;
    }
    if (normalizedRealSelections.includes(peerlessMetaName)) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.duplicateWithReal"));
      return;
    }
    if (mimicMetaName && peerlessMetaName === mimicMetaName) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.duplicateWithMimic"));
      return;
    }
    const useValidation = validateTraitSourceUse(peerlessSpeedSource);
    if (!useValidation.ok) {
      action.shared.reject = true;
      if (useValidation.reason === "missingUsesData") {
        ui.notifications.warn(game.i18n.format("NAS.metamagic.peerlessSpeed.missingUsesData", {
          name: peerlessSpeedSource?.label ?? game.i18n.localize("NAS.metamagic.featureNames.peerlessSpeed")
        }));
      } else {
        ui.notifications.warn(game.i18n.format("NAS.metamagic.peerlessSpeed.noRemainingUses", {
          name: peerlessSpeedSource?.label ?? game.i18n.localize("NAS.metamagic.featureNames.peerlessSpeed")
        }));
      }
      return;
    }
    if (
      [EmpowerSpell.name, MaximizeSpell.name].includes(peerlessMetaName) &&
      !(Array.isArray(context?.damage?.parts) && context.damage.parts.length > 0)
    ) {
      action.shared.reject = true;
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.peerlessSpeed.requiresDamageForEmpowerOrMaximize"));
      return;
    }
    peerlessEnabled = true;
  }

  let succorMetaName = "";
  if (succorSelection && !shouldApplyOracleSuccor) {
    action.shared.reject = true;
    ui.notifications.warn(localizeMetamagic("succorFinalRevelation.notEligible"));
    return;
  }
  if (succorSelection && shouldApplyOracleSuccor) {
    succorMetaName = (resolveMetamagicNameFromDatabase(succorSelection.chosenMetaName) ?? succorSelection.chosenMetaName ?? "").toString().trim();
    if (!succorMetaName || !SUCCOR_ELIGIBLE_METAMAGIC.has(succorMetaName)) {
      action.shared.reject = true;
      ui.notifications.warn(localizeMetamagic("succorFinalRevelation.invalidChoice"));
      return;
    }
    if (normalizedRealSelections.includes(succorMetaName)) {
      action.shared.reject = true;
      ui.notifications.warn(localizeMetamagic("succorFinalRevelation.duplicateWithReal"));
      return;
    }
    if (mimicMetaName && succorMetaName === mimicMetaName) {
      action.shared.reject = true;
      ui.notifications.warn(localizeMetamagic("succorFinalRevelation.duplicateWithMimic"));
      return;
    }
    if (peerlessMetaName && succorMetaName === peerlessMetaName) {
      action.shared.reject = true;
      ui.notifications.warn(localizeMetamagic("succorFinalRevelation.duplicateWithPeerless"));
      return;
    }
  }

  const selectionsForApplication = mimicMetaName
    ? [...normalizedRealSelections, mimicMetaName]
    : [...normalizedRealSelections];
  if (peerlessMetaName) {
    selectionsForApplication.push(peerlessMetaName);
  }
  if (succorMetaName) {
    selectionsForApplication.push(succorMetaName);
  }

  let normalizedSelections = selectionsForApplication
    .map((name) => resolveMetamagicNameFromDatabase(name) ?? name)
    .filter(Boolean);
  if (transmuterUsed) {
    normalizedSelections = normalizedSelections.filter((name) => (name ?? "").toString().trim() !== ExtendSpell.name);
  }

  context.metamagicNames = normalizedSelections;
  context.rollPatch ??= {};
  context.rollPatch.damage ??= {};
  context.rollPatch.damage.transforms ??= [];

  const applied = [];
  const appliedNames = context.metamagic?.applied ?? [];

  if (normalizedSelections.includes(StillSpell.name)) {
    applyStillSpell(context);
    applied.push({ name: StillSpell.name });
  }

  if (normalizedSelections.includes(ReachSpell.name)) {
    const rawTopLevelReachSteps = context?.metamagicOptions?.reachSpellSteps;
    const rawMimicReachSteps = context?.metamagicOptions?.mimicMetamagic?.reachSpellSteps;
    const steps = Number(rawTopLevelReachSteps ?? rawMimicReachSteps ?? 1);
    const didApply = applyReachSpell(context, Number.isFinite(steps) ? steps : 1);
    if (didApply) {
      applied.push({ name: ReachSpell.name });
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== ReachSpell.name);
    }
  }

  if (normalizedSelections.includes(QuickenSpell.name)) {
    const didApply = applyQuickenSpell(context);
    if (didApply) {
      applied.push({ name: QuickenSpell.name });
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== QuickenSpell.name);
    }
  }

  if (normalizedSelections.includes(SelectiveSpell.name)) {
    if (!canApplySelectiveSpell(context)) {
      context.metamagicNames = normalizedSelections.filter((name) => name !== SelectiveSpell.name);
    } else {
      const targets = Array.isArray(action?.shared?.targets) ? action.shared.targets : [];
      const maxExclusions = getSelectiveMaxExclusions(action);
      let excludedIds = [];
      if (maxExclusions > 0 && targets.length) {
        const actionTokenDisposition = action?.token?.disposition;
        const entries = targets
          .map((target, index) => ({
            id: getTargetId(target) ?? `target-${index}`,
            label: getTargetLabel(target, `Target ${index + 1}`),
            target,
            actionTokenDisposition,
          }))
          .filter((entry) => entry.id);
        if (entries.length) {
          const selected = await promptSelectiveExclusions(entries, maxExclusions);
          if (selected === null) {
            action.shared.reject = true;
            context.metamagicNames = normalizedSelections.filter((name) => name !== SelectiveSpell.name);
          } else {
            excludedIds = selected;
          }
        }
      }

      if (context.metamagicNames.includes(SelectiveSpell.name)) {
        applySelectiveSpell(context);
        applied.push({ name: SelectiveSpell.name });
        if (excludedIds.length) {
          const excludedSet = new Set(excludedIds);
          action.shared.targets = targets.filter((target) => !excludedSet.has(getTargetId(target)));
          context.metamagicOptions = {
            ...(context.metamagicOptions ?? {}),
            selectiveExcluded: excludedIds,
          };
        }
      }
    }
  }

  if (normalizedSelections.includes(DazingSpell.name)) {
    const didApply = await applyDazingSpell(context, action, {
      rounds: Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0,
    });
    if (didApply) {
      applied.push({ name: DazingSpell.name });
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== DazingSpell.name);
    }
  }

  if (normalizedSelections.includes(HeightenSpell.name)) {
    const selectedLevel = Number(
      context.metamagicOptions?.heightenSpellLevel
      ?? context.metamagicOptions?.mimicMetamagic?.heightenSpellLevel
      ?? 0
    );
    const didApply = applyHeightenSpell(context, {
      originalLevel: Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0,
      targetLevel: Number.isFinite(selectedLevel) ? selectedLevel : 0,
    });
    if (didApply) {
      applied.push({ name: HeightenSpell.name });
      if (action.shared?.rollData) {
        action.shared.rollData.sl = context.spellLevel?.effective ?? action.shared.rollData.sl;
      }
      if (typeof action?.action?.getDC === "function") {
        action.shared.saveDC = action.action.getDC(action.shared.rollData);
      }
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== HeightenSpell.name);
    }
  }

  if (normalizedSelections.includes(PersistentSpell.name)) {
    const didApply = applyPersistentSpell(context);
    if (didApply) {
      applied.push({ name: PersistentSpell.name });
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== PersistentSpell.name);
    }
  }

  if (normalizedSelections.includes(ExtendSpell.name)) {
    const maskFocusActive = context?.featEffects?.[MASK_FOCUS_ID]?.active === true;
    const didApply = maskFocusActive
      ? applyExtendSpellWithMaskFocus(context)
      : applyExtendSpell(context);
    if (didApply) {
      applied.push({ name: ExtendSpell.name });
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== ExtendSpell.name);
      if (maskFocusActive) {
        if (context.featEffects) {
          delete context.featEffects[MASK_FOCUS_ID];
        }
        if (context.metamagic) {
          delete context.metamagic.extendSlotWaivedByMaskFocus;
        }
        if (context.duration) {
          delete context.duration.maskFocusSelf;
        }
        ui.notifications.warn(game.i18n.localize("NAS.metamagic.maskFocus.extendNotApplicable"));
      }
    }
  }

  if (normalizedSelections.includes(EnlargeSpell.name)) {
    const didApply = applyEnlargeSpell(context, action);
    if (didApply) {
      applied.push({ name: EnlargeSpell.name });
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== EnlargeSpell.name);
    }
  }

  if (normalizedSelections.includes(IntensifiedSpell.name)) {
    const clValue = action.shared?.rollData?.cl;
    const clNum = Number(clValue ?? 0);
    const canIntensify = Number.isFinite(clNum)
      && clNum > 0
      && canIntensifyAnyDamagePart(context?.damage?.parts ?? [], clNum);
    if (canIntensify) {
      context.rollPatch.damage.transforms.push(IntensifiedSpell.key);
      applied.push({ name: IntensifiedSpell.name });
      context.metamagic ??= { applied: [], slotIncrease: 0 };
      context.metamagic.slotIncrease = Number(context.metamagic.slotIncrease ?? 0) + 1;
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== IntensifiedSpell.name);
    }
  }

  if (normalizedSelections.includes(MaximizeSpell.name)) {
    const hasDamageParts = Array.isArray(context?.damage?.parts) && context.damage.parts.length > 0;
    if (hasDamageParts) {
      context.rollPatch.damage.transforms.push(MaximizeSpell.key);
      applied.push({ name: MaximizeSpell.name });
      context.metamagic ??= { applied: [], slotIncrease: 0 };
      context.metamagic.slotIncrease = Number(context.metamagic.slotIncrease ?? 0) + 3;
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== MaximizeSpell.name);
    }
  }

  if (normalizedSelections.includes(EmpowerSpell.name)) {
    const hasDamageParts = Array.isArray(context?.damage?.parts) && context.damage.parts.length > 0;
    if (hasDamageParts) {
      context.rollPatch.damage.transforms.push(EmpowerSpell.key);
      applied.push({ name: EmpowerSpell.name });
      context.metamagic ??= { applied: [], slotIncrease: 0 };
      context.metamagic.slotIncrease = Number(context.metamagic.slotIncrease ?? 0) + 2;
    } else {
      context.metamagicNames = normalizedSelections.filter((name) => name !== EmpowerSpell.name);
    }
  }

  if (normalizedSelections.includes(SilentSpell.name)) {
    const learnedAt = action.item?.system?.learnedAt?.class;
    const learnedList = Array.isArray(learnedAt)
      ? learnedAt
      : learnedAt
        ? [learnedAt]
        : [];
    const hasBard = learnedList.some((entry) => entry?.toString?.().toLowerCase().includes("bard"));
    if (hasBard) {
      ui.notifications.warn(`${SilentSpell.name} cannot be applied to bard spells.`);
      context.metamagicNames = normalizedSelections.filter((name) => name !== SilentSpell.name);
    } else {
      applySilentSpell(context);
      applied.push({ name: SilentSpell.name });
    }
  }

  if (!context.metamagic) {
    context.metamagic = { applied: [], slotIncrease: 0 };
  }

  applied.forEach((entry) => {
    if (!appliedNames.includes(entry.name)) {
      appliedNames.push(entry.name);
    }
  });
  context.metamagic.applied = appliedNames;
  if (!transmuterUsed && naniteArcanaRequested && !appliedNames.includes(ExtendSpell.name)) {
    naniteBloodlineArcanaApplied = applyNaniteBloodlineArcanaDuration(context);
  }
  const mimicWasApplied = mimicEnabled && appliedNames.includes(mimicMetaName);
  const peerlessWasApplied = peerlessEnabled && appliedNames.includes(peerlessMetaName);
  if (peerlessEnabled && !peerlessWasApplied) {
    action.shared.reject = true;
    ui.notifications.warn(game.i18n.format("NAS.metamagic.peerlessSpeed.couldNotApply", { metamagic: peerlessMetaName }));
    return;
  }
  if (peerlessWasApplied) {
    const waivedIncrease = getNominalMetamagicLevelIncrease(peerlessMetaName, context, baseSpellLevel, null);
    const currentIncrease = Number(context?.metamagic?.slotIncrease ?? 0);
    if (Number.isFinite(currentIncrease) && Number.isFinite(waivedIncrease) && waivedIncrease > 0) {
      context.metamagic.slotIncrease = Math.max(0, currentIncrease - waivedIncrease);
    }
    context.metamagic.peerlessSpeedApplied = true;
    context.metamagic.peerlessSpeedMetaName = peerlessMetaName;
    appendPendingTraitUse(action, {
      itemUuid: peerlessSpeedSource?.itemUuid ?? null,
      mode: "use"
    });
  }
  const timelessSoulAvailableForCast =
    shouldApplyTimelessSoul &&
    appliedNames.includes(QuickenSpell.name) &&
    !(peerlessWasApplied && peerlessMetaName === QuickenSpell.name);
  context.metamagic.timelessSoulActive = timelessSoulAvailableForCast;
  let baseSaveDc = resolveBaseSaveDc(action, context);
  if (!Number.isFinite(baseSaveDc?.value) && contextHasSavingThrow(context, action)) {
    const fallback = resolveFeatSaveDcBase(action, context);
    if (Number.isFinite(fallback)) {
      baseSaveDc = { value: fallback, source: "resolveFeatSaveDcBase" };
    }
  }
  if (Number.isFinite(baseSaveDc?.value)) {
    context.save ??= {};
    context.save.baseDc = baseSaveDc.value;
    context.metamagic.baseSaveDc = baseSaveDc.value;
  }

  const sorcererState = await getSorcererArcaneMetamagicState(context?.actor ?? action?.actor, action?.item ?? null);
  const hasAppliedMetamagic = Array.isArray(context?.metamagic?.applied) && context.metamagic.applied.length > 0;
  const spellPerfectionStatus = getSpellPerfectionStatus(context);
  const spontaneousMetafocusStatus = getSpontaneousMetafocusStatus(context);
  let succorWaiverActive = false;
  context.metamagic.succorFinalRevelationSlotAdjustment = 0;
  if (hasAppliedMetamagic && shouldApplyOracleSuccor) {
    const selectedSuccorMeta = (
      resolveMetamagicNameFromDatabase(context?.metamagicOptions?.succorFinalRevelation?.chosenMetaName)
      ?? context?.metamagicOptions?.succorFinalRevelation?.chosenMetaName
      ?? ""
    ).toString().trim();
    const hasSuccorEligibleMetaApplied = selectedSuccorMeta
      ? context.metamagic.applied.includes(selectedSuccorMeta)
      : false;
    if (hasSuccorEligibleMetaApplied && SUCCOR_ELIGIBLE_METAMAGIC.has(selectedSuccorMeta)) {
      const currentTargets = Array.isArray(action?.shared?.targets) ? action.shared.targets : [];
      const allyTargets = currentTargets.filter((target) => isAllyOrSelfTarget(action, target));
      let nonAllyTargets = currentTargets.filter((target) => !isAllyOrSelfTarget(action, target));
      let proceedWithWaiver = true;
      if (nonAllyTargets.length > 0) {
        const response = await promptSuccorFinalRevelationNonAllies(
          nonAllyTargets.map((target, index) => ({
            id: getTargetId(target) ?? `target-${index}`,
            label: getTargetLabel(target, `Target ${index + 1}`)
          }))
        );
        if (response === "reject" || response === null) {
          action.shared.reject = true;
          return;
        }
        if (response === "ignore") {
          action.shared.targets = allyTargets;
          nonAllyTargets = [];
        } else if (response === "continue") {
          proceedWithWaiver = false;
        }
      }
      succorWaiverActive = proceedWithWaiver && nonAllyTargets.length === 0;
      if (succorWaiverActive) {
        context.metamagic.succorFinalRevelationSlotAdjustment = -1;
      }
    }
  }
  const slotMath = computeConsumedSlotIncreaseForMetamagic(
    context,
    baseSpellLevel,
    hasAppliedMetamagic
  );
  if (spellPerfectionStatus.enabled && hasAppliedMetamagic) {
    const baseLevel = Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0;
    const rawModifiedLevel = baseLevel + Number(slotMath.preWaiverConsumedSlotIncrease ?? 0);
    if (rawModifiedLevel > 9) {
      ui.notifications.warn(game.i18n.localize("NAS.metamagic.spellPerfection.maxLevelExceeded"));
      action.shared.reject = true;
      return;
    }
  }
  let bypassCastTimeIncrease = false;
  const activeFeatureLabels = [];
  const pushActiveFeatureLabel = (label) => {
    if (typeof label !== "string" || label.trim().length === 0) return;
    if (!activeFeatureLabels.includes(label)) activeFeatureLabels.push(label);
  };
  if (healerBlessingApplied) {
    pushActiveFeatureLabel(healersBlessingSource?.label ?? localizeMetamagic("featureNames.healersBlessing"));
  }
  if (intenseCelebrationApplied) {
    pushActiveFeatureLabel(intenseCelebrationSource?.label ?? localizeMetamagic("featureNames.intenseCelebration"));
  }
  if (naniteBloodlineArcanaApplied) {
    pushActiveFeatureLabel(naniteBloodlineArcanaSource?.label ?? localizeMetamagic("featureNames.naniteBloodlineArcana"));
  }
  if (oneBodyTwoMindsUsed) {
    pushActiveFeatureLabel(oneBodyTwoMindsSource?.label ?? localizeMetamagic("featureNames.oneBodyTwoMinds"));
  }
  const maleficiumEffect = context?.featEffects?.maleficium;
  if (maleficiumEffect?.active) {
    pushActiveFeatureLabel(maleficiumEffect.label ?? localizeMetamagic("featureNames.maleficium"));
  }
  const maskFocusEffect = context?.featEffects?.[MASK_FOCUS_ID];
  if (maskFocusEffect?.active && context.duration?.maskFocusSelf) {
    pushActiveFeatureLabel(maskFocusEffect.label ?? localizeMetamagic("featureNames.maskFocus"));
  }
  if (mimicWasApplied) {
    pushActiveFeatureLabel(mimicState?.mimicMetamagicItem?.name ?? localizeMetamagic("featureNames.mimicMetamagic"));
  }
  if (peerlessWasApplied) {
    pushActiveFeatureLabel(peerlessSpeedSource?.label ?? localizeMetamagic("featureNames.peerlessSpeed"));
  }
  if (eldritchResearcherEffect?.active) {
    pushActiveFeatureLabel(eldritchResearcherEffect.label ?? localizeMetamagic("featureNames.eldritchResearcher"));
  }
  if (magicalLineageEffect?.active && hasAppliedMetamagic) {
    pushActiveFeatureLabel(magicalLineageEffect.label ?? localizeMetamagic("featureNames.magicalLineage"));
  }
  if (spellPerfectionStatus.enabled) {
    pushActiveFeatureLabel(spellPerfectionStatus.label ?? localizeMetamagic("featureNames.spellPerfection"));
  }
  if (spontaneousMetafocusStatus.enabled) {
    pushActiveFeatureLabel(spontaneousMetafocusStatus.label ?? localizeMetamagic("featureNames.spontaneousMetafocus"));
  }
  if (succorWaiverActive) {
    pushActiveFeatureLabel(oracleSuccorSource?.label ?? localizeMetamagic("featureNames.succorFinalRevelation"));
  }
  if (context?.metamagic?.timelessSoulActive === true) {
    pushActiveFeatureLabel(timelessSoulSource?.label ?? localizeMetamagic("featureNames.timelessSoul"));
  }
  if (transmuterUsed) {
    pushActiveFeatureLabel(transmuterSource?.label ?? localizeMetamagic("featureNames.transmuterOfKorada"));
  }
  if (hasAppliedMetamagic && isClassFeatureEnabledOrDefault(context, RETRIBUTION_FEATURE_ID, false)) {
    pushActiveFeatureLabel(localizeMetamagic("featureNames.retribution"));
  }

  let chosenBypass = null;
  const mimicOnlyRequested = mimicWasApplied && normalizedRealSelections.length === 0;
  const peerlessOnlyRequested = peerlessWasApplied && normalizedRealSelections.length === 0 && !mimicWasApplied;
  if (hasAppliedMetamagic) {
    const selectedArcaneApotheosis = isClassFeatureSelected(context, ARCANE_APOTHEOSIS_FEATURE_ID);
    const selectedMetamagicAdept = isClassFeatureSelected(context, METAMAGIC_ADEPT_FEATURE_ID);
    const selectedOracleSeeker = isClassFeatureSelected(context, SEEKER_ETERNAL_EMPEROR_FEATURE_ID);
    const selectedTraitIds = new Set(getSelectedTraitIds(context));
    const traitById = new Map(Array.isArray(traitSources) ? traitSources.map((source) => [source.id, source]) : []);
    const selectedCuratorTrait = selectedTraitIds.has(CURATOR_MYSTIC_SECRETS_ID)
      ? traitById.get(CURATOR_MYSTIC_SECRETS_ID) ?? null
      : null;
    const hasExtendOrEnlargeApplied =
      Array.isArray(context?.metamagic?.applied) &&
      (
        context.metamagic.applied.includes(ExtendSpell.name)
        || context.metamagic.applied.includes(EnlargeSpell.name)
      );
    const seekerEligibleForBypass =
      shouldApplyOracleSeeker &&
      selectedOracleSeeker &&
      isDivinationSpell(action?.item ?? null) &&
      hasExtendOrEnlargeApplied;
    const tailBypassCandidates = [];

    if (isClassFeatureSelected(context, METAMIXING_FEATURE_ID)) {
      const mxActor = context?.actor ?? action?.actor;
      const mxState = await getMetamixingState(mxActor, action?.item ?? null);
      if (mxState.eligible && mxState.metamixingItem && mxState.reservoirItem) {
        if (!mxState.hasUsesData) {
          ui.notifications.warn(game.i18n.localize("NAS.metamagic.metamixing.missingReservoirUsesData"));
          action.shared.reject = true;
          return;
        }
        if (mxState.usesRemaining < 1) {
          ui.notifications.warn(game.i18n.localize("NAS.metamagic.metamixing.insufficientReservoir"));
          action.shared.reject = true;
          return;
        }
        tailBypassCandidates.push({
          id: METAMIXING_FEATURE_ID,
          label: mxState.metamixingItem?.name ?? localizeMetamagic("featureNames.metamixing"),
          queue: {
            type: "arcaneReservoir",
            payload: {
              itemUuid: mxState.reservoirItem?.uuid ?? null,
              debitCount: 1
            }
          }
        });
      }
    }

    if (sorcererState?.canUsePassive && selectedArcaneApotheosis && sorcererState.arcaneApotheosisAvailable) {
      tailBypassCandidates.push({
        id: ARCANE_APOTHEOSIS_FEATURE_ID,
        label: sorcererState?.arcaneApotheosisItem?.name ?? localizeMetamagic("featureNames.arcaneApotheosis"),
        queue: null
      });
    }
    if (sorcererState?.canUsePassive && selectedMetamagicAdept && sorcererState.metamagicAdeptAvailable) {
      const useValidation = validateLimitedFeatureUse(sorcererState.metamagicAdeptItem);
      if (!useValidation.ok) {
        if (useValidation.reason === "missingUsesData") {
          ui.notifications.warn(localizeMetamagic("metamagicAdept.missingUsesData"));
        } else {
          ui.notifications.warn(localizeMetamagic("metamagicAdept.noRemainingUses"));
        }
      } else {
        tailBypassCandidates.push({
          id: METAMAGIC_ADEPT_FEATURE_ID,
          label: sorcererState?.metamagicAdeptItem?.name ?? localizeMetamagic("featureNames.metamagicAdept"),
          queue: {
            type: "metamagicFeature",
            payload: {
              featureId: METAMAGIC_ADEPT_FEATURE_ID,
              itemUuid: sorcererState?.metamagicAdeptItem?.uuid ?? null
            }
          }
        });
      }
    }
    if (selectedCuratorTrait) {
      const traitValidation = validateTraitSourceUse(selectedCuratorTrait);
      if (!traitValidation.ok) {
        if (traitValidation.reason === "missingUsesData") {
          ui.notifications.warn(`${selectedCuratorTrait.label} has no uses metadata; skipping cast-time bypass.`);
        } else {
          ui.notifications.warn(`${selectedCuratorTrait.label} has no remaining uses.`);
        }
      } else {
        tailBypassCandidates.push({
          id: CURATOR_MYSTIC_SECRETS_ID,
          label: selectedCuratorTrait.label ?? localizeMetamagic("featureNames.curatorMysticSecrets"),
          queue: {
            type: "trait",
            payload: {
              itemUuid: selectedCuratorTrait.itemUuid ?? null,
              mode: selectedCuratorTrait.hasUsesData === true ? "use" : "displayCard"
            }
          }
        });
      }
    }
    if (seekerEligibleForBypass) {
      tailBypassCandidates.push({
        id: SEEKER_ETERNAL_EMPEROR_FEATURE_ID,
        label: oracleSeekerSource?.label ?? localizeMetamagic("featureNames.seekerEternalEmperor"),
        queue: null
      });
    }

    let mmCandidate = null;
    if (isClassFeatureSelected(context, METAMAGIC_MASTERY_FEATURE_ID)) {
      const mmActor = context?.actor ?? action?.actor;
      const wmState = await getWizardMetamagicMasteryState(mmActor, action?.item ?? null);
      const preR = slotMath.consumedSlotIncrease;
      const debitCount = Math.max(1, preR);
      if (wmState.eligible && wmState.metamagicMasteryItem) {
        const spellbookKey = (action?.item?.system?.spellbook ?? "").toString();
        const maxCastable = getSpellbookMaxCastableSpellLevel(mmActor, spellbookKey);
        const baseLevel = Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0;
        if (maxCastable > 0 && baseLevel + preR > maxCastable) {
          ui.notifications.warn(game.i18n.localize("NAS.metamagic.metamagicMastery.maxLevelExceeded"));
          action.shared.reject = true;
          return;
        }
        if (!wmState.hasUsesData) {
          ui.notifications.warn(game.i18n.localize("NAS.metamagic.metamagicMastery.missingUsesData"));
          action.shared.reject = true;
          return;
        }
        if (wmState.usesRemaining < debitCount) {
          ui.notifications.warn(game.i18n.localize("NAS.metamagic.metamagicMastery.insufficientUses"));
          action.shared.reject = true;
          return;
        }
        mmCandidate = {
          id: METAMAGIC_MASTERY_FEATURE_ID,
          label: wmState.metamagicMasteryItem?.name ?? localizeMetamagic("featureNames.metamagicMastery"),
          queue: {
            type: "metamagicFeature",
            payload: {
              featureId: METAMAGIC_MASTERY_FEATURE_ID,
              itemUuid: wmState.metamagicMasteryItem?.uuid ?? null,
              debitCount
            }
          }
        };
      }
    }

    const castTimeBypassCandidates = mmCandidate ? [mmCandidate, ...tailBypassCandidates] : tailBypassCandidates;
    chosenBypass = castTimeBypassCandidates[0] ?? null;
    if (chosenBypass) {
      bypassCastTimeIncrease = true;
      context.metamagic.castTimeBypassSource = chosenBypass.id;
      pushActiveFeatureLabel(chosenBypass.label);
      if (chosenBypass.queue?.type === "metamagicFeature") {
        action.shared ??= {};
        action.shared.nasPendingMetamagicUse = chosenBypass.queue.payload;
      } else if (chosenBypass.queue?.type === "arcaneReservoir") {
        action.shared ??= {};
        action.shared.nasPendingArcaneReservoirSpend = chosenBypass.queue.payload;
      } else if (chosenBypass.queue?.type === "trait") {
        appendPendingTraitUse(action, chosenBypass.queue.payload);
      }
    }
  }

  if (mimicOnlyRequested && chosenBypass?.id !== METAMAGIC_MASTERY_FEATURE_ID) {
    bypassCastTimeIncrease = true;
    context.metamagic.castTimeBypassSource ??= MIMIC_METAMAGIC_FEATURE_ID;
  }
  if (peerlessOnlyRequested) {
    bypassCastTimeIncrease = true;
    context.metamagic.castTimeBypassSource ??= PEERLESS_SPEED_FEATURE_ID;
  }
  if (spellPerfectionStatus.enabled && hasAppliedMetamagic) {
    bypassCastTimeIncrease = true;
    context.metamagic.castTimeBypassSource ??= SPELL_PERFECTION_FEATURE_ID;
  }
  if (spontaneousMetafocusStatus.enabled && hasAppliedMetamagic) {
    bypassCastTimeIncrease = true;
    context.metamagic.castTimeBypassSource ??= SPONTANEOUS_METAFOCUS_FEATURE_ID;
  }
  if (succorWaiverActive) {
    bypassCastTimeIncrease = true;
    context.metamagic.castTimeBypassSource ??= SUCCOR_FINAL_REVELATION_FEATURE_ID;
  }
  if (mimicWasApplied && action.shared?.reject !== true && action.shared?.scriptData?.reject !== true) {
    action.shared ??= {};
    action.shared.nasPendingPhrenicPoolSpend = {
      itemUuid: mimicState?.phrenicPoolItem?.uuid ?? null,
      debitCount: Math.max(1, mimicMetaDebitCount)
    };
  }

  context.metamagic.bypassCastTimeIncrease = bypassCastTimeIncrease;

  applyMetamagicCastTime(action, context);

  const metamagicMasterySlotBypass = chosenBypass?.id === METAMAGIC_MASTERY_FEATURE_ID;
  if (metamagicMasterySlotBypass) {
    context.metamagic.metamagicMasterySlotBypass = true;
  }

  const {
    normalizedOtherIncrease,
    heightenDelta,
    slotAdjustment,
    timelessSoulReduction,
    timelessSoulActive,
    reducedOtherSlotIncrease,
    consumedSlotIncrease: rawConsumedSlotIncrease
  } = slotMath;
  const mimicSlotWaiver = mimicWasApplied ? Math.max(0, mimicMetaIncrease) : 0;
  const consumedAfterMimic = Math.max(0, rawConsumedSlotIncrease - mimicSlotWaiver);
  const consumedSlotIncrease = metamagicMasterySlotBypass ? 0 : consumedAfterMimic;

  context.metamagic.otherSlotIncrease = normalizedOtherIncrease;
  context.metamagic.heightenDelta = Number.isFinite(heightenDelta) ? heightenDelta : 0;
  context.metamagic.slotAdjustment = slotAdjustment;
  context.metamagic.timelessSoulReduction = Number.isFinite(timelessSoulReduction) ? timelessSoulReduction : 0;
  context.metamagic.timelessSoulActive = timelessSoulActive === true;
  context.metamagic.reducedOtherSlotIncrease = reducedOtherSlotIncrease;
  context.metamagic.mimicSlotWaiver = mimicSlotWaiver;
  context.metamagic.consumedSlotIncrease = consumedSlotIncrease;
  // Backward-compatible mirror while consumers migrate to consumedSlotIncrease.
  context.metamagic.effectiveSlotIncrease = consumedSlotIncrease;

  const heightenApplied = Array.isArray(context?.metamagic?.applied)
    && context.metamagic.applied.includes(HeightenSpell.name);
  const shouldApplyArcaneDcBonus =
    hasAppliedMetamagic &&
    sorcererState?.canUsePassive &&
    isClassFeatureEnabledOrDefault(context, ARCANE_BLOODLINE_FEATURE_ID, true) &&
    !heightenApplied &&
    hasAnySlotIncreasingMetamagic(context, Number.isFinite(baseSpellLevel) ? baseSpellLevel : 0) &&
    !metamagicMasterySlotBypass;
  if (shouldApplyArcaneDcBonus) {
    applyArcaneBloodlineDcBonus(action, context);
    pushActiveFeatureLabel(sorcererState?.arcaneBloodlineItem?.name ?? localizeMetamagic("featureNames.arcaneBloodline"));
  }
  const shouldApplyGrandMaestro =
    sorcererState?.grandMaestroAvailable &&
    isClassFeatureEnabledOrDefault(context, GRAND_MAESTRO_FEATURE_ID, true);
  if (shouldApplyGrandMaestro && applyGrandMaestroComponents(context)) {
    pushActiveFeatureLabel(sorcererState?.grandMaestroItem?.name ?? localizeMetamagic("featureNames.grandMaestro"));
  }
  context.metamagic.activeFeatureLabels = activeFeatureLabels;

  if (consumedSlotIncrease > 0) {
    await consumeHigherSpellSlot(action, consumedSlotIncrease);
  }

  await applyRacialTraitSelections(action, context, allTraitSources);

  if (
    context.featEffects?.[MASK_FOCUS_ID]?.active === true
    && context.duration?.maskFocusSelf
    && action.shared?.reject !== true
    && action.shared?.scriptData?.reject !== true
  ) {
    action.shared ??= {};
    action.shared.nasPendingMaskFocusUse = {
      featureId: MASK_FOCUS_FEATURE_ID,
      itemUuid: context.featEffects[MASK_FOCUS_ID]?.itemUuid ?? null
    };
  }
}
