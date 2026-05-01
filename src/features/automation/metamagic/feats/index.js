export {
  applyEldritchResearcherPostMetamagic,
  getEldritchResearcherOptionId,
  getEldritchResearcherSource,
  getEldritchResearcherSpellChoices,
  getEldritchResearcherState,
  getEldritchResearcherStateModes,
  getEldritchResearcherSlotAdjustment,
  getExtendedScryingOptionId,
  getExtendedScryingSource,
  getExtendedScryingState,
  getSpellPerfectionOptionId,
  getSpellPerfectionSource,
  getSpellPerfectionState,
  getSpellPerfectionStatus,
  getSpontaneousMetafocusOptionId,
  getSpontaneousMetafocusSource,
  getSpontaneousMetafocusStatus,
  prepareEldritchResearcherContext,
  prepareExtendedScryingContext,
  prepareSpontaneousMetafocusContext,
  prepareSpellPerfectionContext,
  promptEldritchResearcherSpellSelection,
  registerEldritchResearcherItemHook,
  registerSpontaneousMetafocusItemHook,
  registerSpellPerfectionItemHook,
  resolveSpellLabelFromChoices,
  setEldritchResearcherState,
  setExtendedScryingState,
  setSpellPerfectionState
} from "./standaloneFeats.js";

export {
  applyMaleficiumPostMetamagic,
  getMaleficiumMinimumConsumedSlotLevel,
  getMaleficiumOptionId,
  getMaleficiumSlotAdjustment,
  getMaleficiumSource,
  getMaleficiumState,
  prepareMaleficiumContext,
  setMaleficiumState,
  spellItemHasEvilDescriptor
} from "./damnationFeats.js";

export {
  getMaskFocusSource,
  MASK_FOCUS_FEATURE_ID,
  MASK_FOCUS_ID,
  prepareMaskFocusContext
} from "./maskFocusFeats.js";
