
import { MODULE } from '../../../common/module.js';
import { socket } from '../../../integration/moduleSockets.js';

export async function handleBuffAutomation(action) {
  
  if (action.item.type === "feat" && action.item.subType === "classFeat") {
    const isBuff = action.item.hasItemBooleanFlag('buff');
    if (!isBuff) return;
  }

  let searchName = action.item.name;
  if (action.item.type === "consumable" && typeof action.item.subType === "string") {
    const subType = action.item.subType.toLowerCase();
    let prefixKey = null;
    if (subType === "wand") prefixKey = "PF1.CreateItemWandOf";
    else if (subType === "scroll") prefixKey = "PF1.CreateItemScrollOf";
    else if (subType === "potion") prefixKey = "PF1.CreateItemPotionOf";
    if (prefixKey) {
      let localized = game.i18n.localize(prefixKey); 
      let prefix = localized.replace(/\{name\}/, "").trim();
      if (searchName.toLowerCase().startsWith(prefix.toLowerCase())) {
        searchName = searchName.slice(prefix.length).trim();
      }
    }
  }
  
  const modifierNames = game.settings.get(MODULE.ID, 'modifierNames') || {};
  const communalString = modifierNames.communal || 'Communal';
  let isCommunal = false;
  const communalEndRegex = new RegExp(`(?:,\\s*|\\s*\\(|\\s*\\[|\\s+)${communalString}\\s*(?:\\)|\\])?$`, 'i');
  const communalStartRegex = new RegExp(`^${communalString}[,\s]+`, 'i');

  if (communalStartRegex.test(searchName)) {
    isCommunal = true;
    searchName = searchName.replace(communalStartRegex, '').trim();
  } else if (communalEndRegex.test(searchName)) {
    isCommunal = true;
    searchName = searchName.replace(communalEndRegex, '').trim();
    searchName = searchName.replace(/[\s,]+$/, '').trim();
  }
  
  const hasTargets = action.shared.targets && action.shared.targets.length > 0;
  
  const rangeUnits =
    action.shared?.nasSpellContext?.range?.range?.units ?? action.action?.range?.units;
  const targetValue = action.action?.target?.value;
  
  const isSelfTargeting = rangeUnits === "personal" || targetValue === "you";
  
  const casterLevel = action.shared.rollData?.cl;

  const durationContext = action.shared?.nasSpellContext?.duration;
  const durationUnits = durationContext?.units ?? action.action?.duration?.units;

  const rawDurationValue = durationContext?.value ?? action.action?.duration?.value ?? '';

  let durationValue;
  if (durationContext?.evaluated?.total != null) {
    durationValue = durationContext.evaluated.total;
  } else {
    try {
      durationValue = (await new Roll(rawDurationValue, action.shared.rollData).evaluate()).total;
    } catch (err) {
      console.warn(`${MODULE.ID} | Failed to evaluate duration formula "${rawDurationValue}". Using numeric fallback if possible.`, err);
      const numericFallback = Number(rawDurationValue);
      durationValue = Number.isNaN(numericFallback) ? 0 : numericFallback;
    }
  }

  let communalPromptForManual = false;
  let communalIncrement = null;
  let communalTotalDuration = null;
  let communalDurationUnit = null;
  let communalDurationFormula = null;

  if (isCommunal) {
    const communalHandling = game.settings.get(MODULE.ID, 'communalHandling');
    const communalParse = await parseCommunalDuration({
      action,
      durationUnits,
      rawDurationValue,
      casterLevel
    });

    if (communalParse && communalParse.totalDuration !== null) {
      communalIncrement = communalParse.increment;
      communalTotalDuration = communalParse.totalDuration;
      communalDurationUnit = communalParse.unit || durationUnits;
      communalDurationFormula = communalParse.formula;
      communalPromptForManual = communalHandling === 'prompt' || communalHandling === 'even';
    }
  }
  
  const areaString = action.action?.area;
  const measureTemplateEnabled = action.formData && action.formData["measure-template"];
  const templateSize = Number(action.action?.measureTemplate?.size || 0);
  const isAreaOfEffect = !!areaString || (measureTemplateEnabled && templateSize > 5);
      
  const matchingBuffs = await findMatchingBuffs(searchName);
  
  if (matchingBuffs.length > 0) {
    if (!hasTargets && !isSelfTargeting) {
      const mode = game.settings.get(MODULE.ID, 'buffAutomationMode');
      
      if (mode === "strict") {
        console.warn(`${MODULE.ID} | Buff automation canceled: No targets selected for ${action.item.name}`);
        action.shared.reject = true;
        ui.notifications.warn(game.i18n.format('NAS.buffs.NoTargetsSelected', { name: action.item.name }));
        return;
      } else if (mode === "lenient") {
        console.warn(`${MODULE.ID} | Buff automation skipped: No targets selected for ${action.item.name}`);
        ui.notifications.info(game.i18n.format('NAS.buffs.UnableToApplyAutomaticBuffs', { name: action.item.name }));
      }
    }
    let selectedBuff = null;
    
    const categorizedMatches = categorizeBuffMatches(action.item.name, matchingBuffs);
    
    if (categorizedMatches.exact.length === 1) {
      selectedBuff = categorizedMatches.exact[0];
    } 
    else if (categorizedMatches.variants.length > 0) {
      const targetContext = await gatherTargetsForApplication({
        action,
        isSelfTargeting,
        isCommunal,
        durationUnits,
        durationValue,
        communalIncrement,
        communalTotalDuration,
        communalDurationUnit,
        communalPromptForManual,
        isAreaOfEffect
      });
      if (targetContext.rejected) return;

      const variantPlan = await promptBuffSelection(categorizedMatches.variants, action, {
        mode: 'variant',
        targets: targetContext.filteredTargets,
        perTargetDurations: targetContext.perTargetDurations
      });
      if (!variantPlan) {
        action.shared.reject = true;
        return;
      }

      await handleVariantPlanApplication({
        action,
        variants: categorizedMatches.variants,
        plan: variantPlan,
        targetContext,
        durationUnits,
        durationValue,
        casterLevel
      });
      return;
    }
    else if (categorizedMatches.versions.length > 0 && categorizedMatches.exact.length === 0) {
      const exactNameMatch = categorizedMatches.versions.find(
        b => b.name.toLowerCase() === action.item.name.toLowerCase()
      );
      
      if (exactNameMatch) {
        selectedBuff = exactNameMatch;
      } else {
        selectedBuff = await promptBuffSelection(categorizedMatches.versions, action);
        if (!selectedBuff) {
          action.shared.reject = true;
          return;
        }
      }
    }
    else if (matchingBuffs.length > 0) {
      selectedBuff = await promptBuffSelection(matchingBuffs, action);
      if (!selectedBuff) {
        action.shared.reject = true;
        return;
      }
    }
    
    if (selectedBuff) {
      const targetContext = await gatherTargetsForApplication({
        action,
        isSelfTargeting,
        isCommunal,
        durationUnits,
        durationValue,
        communalIncrement,
        communalTotalDuration,
        communalDurationUnit,
        communalPromptForManual,
        isAreaOfEffect
      });
      if (targetContext.rejected) return;

      const { filteredTargets, perTargetDurations } = targetContext;

      if (perTargetDurations && perTargetDurations.length > 0) {
        for (const entry of perTargetDurations) {
          await applyBuffToTargets(selectedBuff, [entry.target], {
            units: entry.duration.units,
            value: String(entry.duration.value)
          }, casterLevel);
        }
        return;
      } else {
        await applyBuffToTargets(selectedBuff, filteredTargets, {
          units: durationUnits,
          value: String(durationValue)
        }, casterLevel);
      }
    }
  }
}

function categorizeBuffMatches(spellName, buffs) {
  const normalizedSpellName = spellName.toLowerCase();
  const result = {
    exact: [],    
    versions: [], 
    variants: []  
  };
  
  buffs.forEach(buff => {
    const buffName = buff.name.toLowerCase();
    
    if (buffName === normalizedSpellName) {
      result.exact.push(buff);
    } 
    else if (buffName.includes('(') && buffName.includes(')')) {
      result.variants.push(buff);
    } 
    else if (buffName.includes(',')) {
      result.versions.push(buff);
    } 
    else {
      result.exact.push(buff);
    }
  });
  
  return result;
}

export async function findMatchingBuffs(name) {
  const normalizedName = name.toLowerCase();
  let exactMatches = [];
  let partialMatches = [];

  try {
    const customCompendia = game.settings.get(MODULE.ID, 'customBuffCompendia') || [];
    const useWorldBuffs = customCompendia.includes("__world__");
    const compendia = customCompendia.filter(packPath => packPath && packPath !== "__world__" && game.packs.get(packPath));

    for (const packKey of compendia) {
      const pack = game.packs.get(packKey);
      if (!pack) {
        console.warn(`${MODULE.ID} | Compendium ${packKey} not found`);
        continue;
      }

      const index = await pack.getIndex();

      const exactEntries = index.filter(i => i.name.toLowerCase() === normalizedName);
      const partialEntries = index.filter(i =>
        i.name.toLowerCase().includes(normalizedName) &&
        !exactEntries.some(em => em._id === i._id)
      );

      for (const entry of exactEntries) {
        const document = await pack.getDocument(entry._id);
        if (document.type !== "buff") continue;
        exactMatches.push({
          name: document.name,
          id: document.id,
          pack: packKey,
          document: document
        });
      }

      for (const entry of partialEntries) {
        const document = await pack.getDocument(entry._id);
        if (document.type !== "buff") continue;
        partialMatches.push({
          name: document.name,
          id: document.id,
          pack: packKey,
          document: document
        });
      }
    }

    let worldExactMatches = [];
    let worldPartialMatches = [];
    if (useWorldBuffs) {
      const worldBuffs = game.items.filter(item => item.type === "buff");
      worldExactMatches = worldBuffs.filter(item => item.name.toLowerCase() === normalizedName).map(item => ({
        name: item.name,
        id: item.id,
        pack: null,
        document: item
      }));
      worldPartialMatches = worldBuffs.filter(item =>
        item.name.toLowerCase().includes(normalizedName) &&
        !worldExactMatches.some(em => em.id === item.id)
      ).map(item => ({
        name: item.name,
        id: item.id,
        pack: null,
        document: item
      }));
    }

    if (exactMatches.length > 0 || worldExactMatches.length > 0) {
      return [...exactMatches, ...worldExactMatches];
    }

    if (partialMatches.length > 0 || worldPartialMatches.length > 0) {
      return [...partialMatches, ...worldPartialMatches];
    }

    function normalizeTokens(str) {
      return str
        .toLowerCase()
        .replace(/[,()]/g, '') 
        .split(/\s+/)
        .filter(Boolean)
        .sort()
        .join(' ');
    }

    const normalizedSpellTokens = normalizeTokens(name);

    for (const packKey of compendia) {
      const pack = game.packs.get(packKey);
      if (!pack) continue;
      const index = await pack.getIndex();
      for (const entry of index) {
        const buffTokens = normalizeTokens(entry.name);
        if (buffTokens === normalizedSpellTokens) {
          const document = await pack.getDocument(entry._id);
          if (document.type === "buff") {
            return [{
              name: document.name,
              id: document.id,
              pack: packKey,
              document: document
            }];
          }
        }
      }
    }
    if (useWorldBuffs) {
      const worldBuffs = game.items.filter(item => item.type === "buff");
      for (const item of worldBuffs) {
        const buffTokens = normalizeTokens(item.name);
        if (buffTokens === normalizedSpellTokens) {
          return [{
            name: item.name,
            id: item.id,
            pack: null,
            document: item
          }];
        }
      }
    }
  } catch (error) {
    console.error(`${MODULE.ID} | Error searching for buffs:`, error);
  }

  return [];
}

async function gatherTargetsForApplication({
  action,
  isSelfTargeting,
  isCommunal,
  durationUnits,
  durationValue,
  communalIncrement,
  communalTotalDuration,
  communalDurationUnit,
  communalPromptForManual,
  isAreaOfEffect
}) {
  let filteredTargets = action.shared.targets || [];
  const filteringMode = game.settings.get(MODULE.ID, 'buffTargetFiltering');
  const personalTargeting = game.settings.get(MODULE.ID, 'personalTargeting');
  let perTargetDurations = null;

  if (filteringMode === "byDisposition") {
    if (isSelfTargeting) {
      if (personalTargeting === 'deny') {
        filteredTargets = [action.token];
      } else {
        filteredTargets = filteredTargets.filter(target => {
          const targetDisposition = target.document ? target.document.disposition : target.disposition;
          const actionDisposition = action.token.disposition;
          return targetDisposition === actionDisposition;
        });
        if (!filteredTargets.some(t => t.id === action.token.id)) {
          filteredTargets.unshift(action.token);
        }
      }
    } else {
      filteredTargets = filteredTargets.filter(target => {
        const targetDisposition = target.document ? target.document.disposition : target.disposition;
        const actionDisposition = action.token.disposition;
        return targetDisposition === actionDisposition;
      });
      if (isCommunal) {
        perTargetDurations = await handleCommunalDuration({
          isCommunal,
          filteredTargets,
          durationUnits: communalDurationUnit || durationUnits,
          durationValue,
          communalIncrement,
          communalTotalDuration,
          communalDurationUnit,
          action
        });
        if (!perTargetDurations) return { rejected: true };
      }
    }
  } else if (filteringMode === "manualSelection") {
    if (communalPromptForManual && communalIncrement && communalTotalDuration) {
      const communalResult = await promptTargetSelection(filteredTargets, action, {
        communal: true,
        increment: communalIncrement,
        total: communalTotalDuration,
        unit: communalDurationUnit || durationUnits
      });
      if (communalResult.length > 0 && communalResult[0].target && communalResult[0].duration !== undefined) {
        perTargetDurations = communalResult;
      } else {
        filteredTargets = communalResult;
      }
    } else {
      if (isSelfTargeting) {
        if (personalTargeting === 'deny') {
          filteredTargets = [action.token];
        } else {
          if (!filteredTargets.some(t => t.id === action.token.id)) {
            filteredTargets.unshift(action.token);
          }
          if (isCommunal) {
            filteredTargets = await promptTargetSelection(filteredTargets, action, { communal: isCommunal });
          } else {
            filteredTargets = await promptTargetSelection(filteredTargets, action);
          }
        }
      } else {
        if (isCommunal) {
          perTargetDurations = await handleCommunalDuration({
            isCommunal,
            filteredTargets,
            durationUnits: communalDurationUnit || durationUnits,
            durationValue,
            communalIncrement,
            communalTotalDuration,
            communalDurationUnit,
            action
          });
          if (!perTargetDurations) return { rejected: true };
        } else {
          filteredTargets = await promptTargetSelection(filteredTargets, action);
        }
      }
    }
  } else {
    if (isSelfTargeting) {
      if (personalTargeting === 'deny') {
        filteredTargets = [action.token];
      } else {
        if (!filteredTargets.some(t => t.id === action.token.id)) {
          filteredTargets.unshift(action.token);
        }
      }
    } else {
      if (isCommunal) {
        perTargetDurations = await handleCommunalDuration({
          isCommunal,
          filteredTargets,
          durationUnits: communalDurationUnit || durationUnits,
          durationValue,
          communalIncrement,
          communalTotalDuration,
          communalDurationUnit,
          action
        });
        if (!perTargetDurations) return { rejected: true };
      }
    }
  }

  if ((!filteredTargets || filteredTargets.length === 0) && game.settings.get(MODULE.ID, 'buffAutomationMode') === "strict") {
    action.shared.reject = true;
    return { rejected: true };
  }

  const slotInfo = checkAndConsumeSpellSlots({
    action,
    filteredTargets,
    isCommunal,
    isAreaOfEffect
  });
  if (slotInfo && slotInfo.rejected) return { rejected: true };

  return { filteredTargets, perTargetDurations, slotInfo, rejected: false };
}

export async function promptBuffSelection(buffs, action, options = {}) {
  if (!buffs || buffs.length === 0) return null;
  const mode = options.mode || 'simple';

  if (mode === 'variant') {
    const targets = options.targets || [];
    const spellKey = getSpellKey(action);
    const mappings = game.settings.get(MODULE.ID, 'pairedBuffMappings') || {};
    const remembered = mappings[spellKey] || {};
    const variantCapMode = game.settings.get(MODULE.ID, 'variantTargetCap') || 'hint';
    const casterLevel = action.shared.rollData?.cl ?? action.item?.system?.level ?? 0;
    const parsedCap = estimateScalableTargets(
      action.action?.target?.value ||
      action.item?.system?.actions?.[0]?.target?.value ||
      action.item?.system?.target?.value,
      casterLevel
    );
    const targetCap = parsedCap && parsedCap > 0 ? parsedCap : null;

    const variantLabels = buffs.map(b => {
      const paren = b.name.match(/\(([^)]+)\)/);
      return paren ? paren[1].trim() : b.name;
    });

    const findVariantIndex = (ref) => {
      if (!ref) return null;
      const idx = buffs.findIndex(b => b.id === ref.id && (b.pack || null) === (ref.pack || null));
      return idx >= 0 ? idx : null;
    };

    const renderOptions = (selectedIdx = 0) => variantLabels.map((lbl, idx) =>
      `<option value="${idx}" ${idx === selectedIdx ? 'selected' : ''}>${lbl}</option>`
    ).join('');

    const rememberedAlliesToggle = (remembered.applyAllies !== undefined ? remembered.applyAllies : !!remembered.allies);
    const rememberedFoesToggle = (remembered.applyFoes !== undefined ? remembered.applyFoes : !!remembered.foes);
    const allyDefaultIdx = findVariantIndex(remembered.allies) ?? 0;
    const foeDefaultIdx = findVariantIndex(remembered.foes) ?? 0;

    const targetCards = targets.map((target, index) => {
      const tokenName = target.name || target.actor?.name || `Target ${index + 1}`;
      const tokenImg = target.document?.texture?.src || target.texture?.src || "";
      const disposition = target.document?.disposition ?? target?.disposition;
      const isSameDisposition = disposition === action.token?.disposition;
      const rememberedEntry = remembered?.perTarget?.find?.(pt => (pt.actorId && pt.actorId === target.actor?.id) || (pt.tokenId && pt.tokenId === target.id));
      const rememberedVariantIdx = typeof rememberedEntry?.variantIndex === 'number' ? rememberedEntry.variantIndex : (isSameDisposition ? allyDefaultIdx : foeDefaultIdx);
      const rememberedApplyTiming = rememberedEntry?.applyTiming || (rememberedEntry?.applyOnTurn ? 'turn' : 'cast');
      const preChecked = targetCap ? (index < targetCap) : true;
      return `
        <div class="target-option" data-target-index="${index}" style="display: flex; flex-direction: column; align-items: center; width: 170px; border: 1px solid #ccc; border-radius: 6px; padding: 6px;">
          <div style="align-self: flex-start; margin-bottom: 4px;">
            <input type="checkbox" class="ic-target-enabled" id="ic-target-enabled-${index}" ${preChecked ? 'checked' : ''}/>
          </div>
          <img src="${tokenImg}" style="width: 64px; height: 64px; border: 2px solid ${isSameDisposition ? 'green' : 'red'}; border-radius: 5px;" />
          <label style="margin: 4px 0; font-weight: 600;">${tokenName}</label>
          <select id="ic-target-variant-${index}" style="width: 100%;">${renderOptions(rememberedVariantIdx)}</select>
          <div style="margin-top: 6px; display: flex; gap: 12px; align-items: center; justify-content: center;">
            <label style="display:flex; gap:4px; align-items:center;" title="${game.i18n.localize('NAS.buffs.ApplyOnCastTooltip') || 'Apply immediately on cast'}">
              <input type="radio" name="ic-target-timing-${index}" value="cast" ${rememberedApplyTiming === 'turn' ? '' : 'checked'}/>
              ${game.i18n.localize('NAS.buffs.ApplyOnCastShort') || 'On cast'}
            </label>
            <label style="display:flex; gap:4px; align-items:center;" title="${game.i18n.localize('NAS.buffs.ApplyOnTurnTooltip') || 'Apply at the start of the target’s turn'}">
              <input type="radio" name="ic-target-timing-${index}" value="turn" ${rememberedApplyTiming === 'turn' ? 'checked' : ''}/>
              ${game.i18n.localize('NAS.buffs.ApplyOnTurnShort') || 'On turn'}
            </label>
          </div>
          <small style="color: ${isSameDisposition ? 'green' : 'red'};">${isSameDisposition ? 'Ally' : 'Foe'}</small>
        </div>
      `;
    }).join('');

    const initialSwitching = !!remembered?.allowSwitching;
    let content = `<p>${game.i18n.localize('NAS.buffs.SelectBuffVariant')}: ${action.item?.name || 'Spell'}</p>`;
    content += `
      <div class="form-group" style="display:grid; grid-template-columns: auto 1fr; align-items:center; gap:8px;">
        <label style="display:flex; gap:6px; align-items:center;">
          <input type="checkbox" id="ic-ally-toggle" ${rememberedAlliesToggle ? 'checked' : ''}/>
          ${game.i18n.localize('NAS.buffs.Allies') || 'Allies'}
        </label>
        <select id="ic-ally-variant" style="width: 100%; display:${initialSwitching ? 'none' : (rememberedAlliesToggle ? 'block' : 'none')};">
          <option value="none">None</option>
          ${renderOptions(allyDefaultIdx)}
        </select>
      </div>
      <div class="form-group" style="display:grid; grid-template-columns: auto 1fr; align-items:center; gap:8px;">
        <label style="display:flex; gap:6px; align-items:center;">
          <input type="checkbox" id="ic-foe-toggle" ${rememberedFoesToggle ? 'checked' : ''}/>
          ${game.i18n.localize('NAS.buffs.Foes') || 'Foes'}
        </label>
        <select id="ic-foe-variant" style="width: 100%; display:${initialSwitching ? 'none' : (rememberedFoesToggle ? 'block' : 'none')};">
          <option value="none">None</option>
          ${renderOptions(foeDefaultIdx)}
        </select>
      </div>
      <div class="form-group" style="display:flex; gap:8px; align-items:center; margin-top:6px;">
        <input type="checkbox" id="ic-allow-switching" ${initialSwitching ? 'checked' : ''}/>
        <label for="ic-allow-switching">${game.i18n.localize('NAS.buffs.AllowSwitchingEachRound') || 'Allow switching each round'}</label>
      </div>
      ${targetCap ? `<div id="ic-cap-hint" style="margin: 6px 0; color: var(--color-text);">${game.i18n.format('NAS.buffs.TargetCapHintRemaining', { cap: targetCap, remaining: targetCap })}</div>` : ''}
      <div id="ic-target-section" style="display:${initialSwitching ? 'grid' : 'none'}; grid-template-columns: repeat(auto-fit, minmax(160px, 170px)); gap: 12px; justify-content: flex-start; border:1px solid #ccc; padding:8px; border-radius:6px; max-height: 340px; overflow-y:auto;">
        ${targetCards}
      </div>
      <div class="form-group" style="display: flex; align-items: center; gap: 6px; margin-top: 8px;">
        <label style="display: inline-flex; align-items: center; gap: 6px; margin: 0;">
          <input type="checkbox" id="ic-remember-mapping" style="margin: 0;"/>
          ${game.i18n.localize('NAS.buffs.RememberForSpell') || 'Remember for this spell'}
        </label>
      </div>
    `;

    return new Promise(resolve => {
      const dlg = new Dialog({
        title: game.i18n.localize('NAS.buffs.SelectBuffVariant'),
        content,
        buttons: {
          apply: {
            icon: '<i class="fas fa-check"></i>',
            label: game.i18n.localize('NAS.buffs.ApplyBuff'),
            callback: html => {
              const getEl = (sel) => typeof html.find === 'function' ? html.find(sel) : html.querySelector(sel);
              const allyToggle = getEl('#ic-ally-toggle');
              const foeToggle = getEl('#ic-foe-toggle');
              const allowSwitchingEl = getEl('#ic-allow-switching');
              const rememberEl = getEl('#ic-remember-mapping');
              const allySelect = getEl('#ic-ally-variant');
              const foeSelect = getEl('#ic-foe-variant');

              const allyChecked = allyToggle && (allyToggle.checked ?? allyToggle.prop?.('checked'));
              const foeChecked = foeToggle && (foeToggle.checked ?? foeToggle.prop?.('checked'));

              const allyIdxRaw = allySelect ? (allySelect.value ?? allySelect.val?.()) : null;
              const foeIdxRaw = foeSelect ? (foeSelect.value ?? foeSelect.val?.()) : null;
              const allyIdx = allyChecked && allyIdxRaw !== 'none' ? Number(allyIdxRaw) : null;
              const foeIdx = foeChecked && foeIdxRaw !== 'none' ? Number(foeIdxRaw) : null;
              const allowSwitching = allowSwitchingEl ? (allowSwitchingEl.checked ?? allowSwitchingEl.prop?.('checked')) : false;
              const remember = rememberEl ? (rememberEl.checked ?? rememberEl.prop?.('checked')) : false;

              const perTarget = [];
              if (allowSwitching && targets.length > 0) {
                let enabledCount = 0;
                targets.forEach((t, idx) => {
                  const enabledEl = getEl(`#ic-target-enabled-${idx}`);
                  const enabled = enabledEl ? (enabledEl.checked ?? enabledEl.prop?.('checked')) : true;
                  if (enabled) enabledCount += 1;
                  const selectEl = getEl(`#ic-target-variant-${idx}`);
                  const variantIdx = Number((selectEl?.value ?? selectEl?.val?.())) ?? 0;
                  const timingCast = getEl(`input[name="ic-target-timing-${idx}"][value="cast"]`);
                  const timingTurn = getEl(`input[name="ic-target-timing-${idx}"][value="turn"]`);
                  let applyTiming = 'cast';
                  if (timingTurn && (timingTurn.checked ?? timingTurn.prop?.('checked'))) applyTiming = 'turn';
                  else if (timingCast && (timingCast.checked ?? timingCast.prop?.('checked'))) applyTiming = 'cast';
                  perTarget.push({
                    targetId: t.id,
                    actorId: t.actor?.id,
                    variantIndex: variantIdx,
                    applyTiming,
                    enabled
                  });
                });

                if (targetCap && enabledCount > targetCap && variantCapMode === 'enforce') {
                  ui.notifications.warn(game.i18n.format('NAS.buffs.TargetCapExceeded', { cap: targetCap }));
                  action.shared.reject = true;
                  return;
                }
              }

      const result = {
                allies: allowSwitching ? null : (allyIdx !== null && !Number.isNaN(allyIdx) ? buffs[allyIdx] : null),
                foes: allowSwitching ? null : (foeIdx !== null && !Number.isNaN(foeIdx) ? buffs[foeIdx] : null),
                allowSwitching,
                remember,
        perTarget,
                applyAllies: !!allyChecked,
                applyFoes: !!foeChecked,
                variants: buffs
      };

      if (allowSwitching && targetCap) {
        result.perTarget = result.perTarget.filter(pt => pt.enabled !== false);
      }

      resolve(result);
            }
          },
          cancel: {
            icon: '<i class="fas fa-times"></i>',
            label: game.i18n.localize('NAS.common.buttons.cancel'),
            callback: () => resolve(null)
          }
        },
        default: "apply",
        close: () => resolve(null),
        render: html => {
          const $html = typeof html.find === 'function' ? html : $(html);
          const relayout = () => dlg.setPosition({ height: 'auto' });
          $html.find('#ic-ally-toggle').on('change', ev => {
            const checked = ev.currentTarget.checked;
            const allowSwitch = $html.find('#ic-allow-switching').prop ? $html.find('#ic-allow-switching').prop('checked') : $html.find('#ic-allow-switching')[0]?.checked;
            if (!allowSwitch) $html.find('#ic-ally-variant').css('display', checked ? 'block' : 'none');
            relayout();
          });
          $html.find('#ic-foe-toggle').on('change', ev => {
            const checked = ev.currentTarget.checked;
            const allowSwitch = $html.find('#ic-allow-switching').prop ? $html.find('#ic-allow-switching').prop('checked') : $html.find('#ic-allow-switching')[0]?.checked;
            if (!allowSwitch) $html.find('#ic-foe-variant').css('display', checked ? 'block' : 'none');
            relayout();
          });
          $html.find('#ic-allow-switching').on('change', ev => {
            const checked = ev.currentTarget.checked;
            $html.find('#ic-target-section').css('display', checked ? 'grid' : 'none');
            $html.find('#ic-ally-variant').css('display', checked ? 'none' : ($html.find('#ic-ally-toggle').prop('checked') ? 'block' : 'none'));
            $html.find('#ic-foe-variant').css('display', checked ? 'none' : ($html.find('#ic-foe-toggle').prop('checked') ? 'block' : 'none'));
            relayout();
          });
          if (targetCap) {
            const updateCapHint = () => {
              const enabledCount = $html.find('.ic-target-enabled:checked').length;
              const hintEl = $html.find('#ic-cap-hint');
              if (hintEl.length) {
                const remaining = Math.max(targetCap - enabledCount, 0);
                hintEl.text(game.i18n.format('NAS.buffs.TargetCapHintRemaining', { cap: targetCap, remaining }));
                hintEl.css('color', enabledCount > targetCap ? 'red' : '');
              }
            };
            $html.find('.ic-target-enabled').on('change', updateCapHint);
            updateCapHint();
          }
          setTimeout(relayout, 0);
        }
      }, { width: (() => {
        const cols = Math.min(Math.max(targets.length || 1, 1), 4);
        const card = 170;
        const gap = 12;
        const chrome = 120; 
        return Math.max(400, cols * card + (cols - 1) * gap + chrome);
      })() });
      dlg.render(true);
    });
  }

  if (buffs.length === 1) return buffs[0];

  return new Promise(resolve => {
    const baseItemName = action.item.name;
    let content = `<p>${game.i18n.format('NAS.buffs.MultipleBuffOptionsFound', { name: baseItemName })}</p>`;
    content += `<div class="form-group"><select id="buff-select" name="buff-select" style="width: 100%;">`;
    buffs.forEach((buff, index) => {
      let displayName = buff.name;
      const match = buff.name.match(/\((.*?)\)/);
      if (match) displayName = match[1];
      else if (buff.name.includes(',')) displayName = buff.name;
      const pack = game.packs.get(buff.pack);
      let packName = buff.pack;
      if (pack) {
        const label = pack.metadata.label;
        packName = label && label.includes('.') ? game.i18n.localize(label) : label;
      }
      content += `<option value="${index}">${displayName} (${packName})</option>`;
    });
    content += `</select></div>`;
    const dialog = new Dialog({
      title: game.i18n.localize('NAS.buffs.SelectBuffVariant'),
      content: content,
      buttons: {
        select: {
          icon: '<i class="fas fa-check"></i>',
          label: game.i18n.localize('NAS.buffs.Select'),
          callback: html => {
            let selectedIndex;
            if (typeof html.find === 'function') {
              selectedIndex = Number(html.find('#buff-select').val());
            } else {
              const select = html.querySelector('#buff-select');
              selectedIndex = select ? Number(select.value) : 0;
            }
            resolve(buffs[selectedIndex]);
          }
        },
        cancel: {
          icon: '<i class="fas fa-times"></i>',
          label: game.i18n.localize('NAS.common.buttons.cancel'),
          callback: () => resolve(null)
        }
      },
      default: "select",
      close: () => resolve(null)
    });
    dialog.render(true);
  });
}

export async function promptTargetSelection(targets, action, communalOptions = null) {
  let filteredTargets = targets;
  if (!game.user.isGM) {
    const casterToken = action.token;
    const casterHasSeeInvisibility = casterToken?.actor?.system?.traits?.senses?.si === true;
    filteredTargets = targets.filter(token => {
      const actor = token.actor;
      const isInvisible = actor.statuses.has("invisible");
      const isHidden = token.document.hidden;
      const disposition = token.document?.disposition ?? token?.disposition;
      if (disposition === CONST.TOKEN_DISPOSITIONS.SECRET) return false;
      if (isHidden) return false;
      if (isInvisible && !casterHasSeeInvisibility) return false;
      if (casterToken && canvas?.visibility?.testVisibility) {
        const isVisible = canvas.visibility.testVisibility(token.center, {
          object: token,
          visionSource: casterToken.vision,
        });
        if (!isVisible) return false;
      }
      return true;
    });
  }

  const useEnhancedCommunalDialog = communalOptions &&
    communalOptions.communal &&
    communalOptions.increment &&
    communalOptions.total &&
    communalOptions.unit;

  if (useEnhancedCommunalDialog) {
    const communalHandling = game.settings.get(MODULE.ID, 'communalHandling');
    const increment = communalOptions.increment;
    const total = communalOptions.total;
    const unit = communalOptions.unit;
    const n = filteredTargets.length;
    if (n <= 0) return [];
    const perTargetEven = Math.floor(total / n / increment) * increment;
    const isDivisible = perTargetEven > 0 && (perTargetEven * n === total);
    if (communalHandling === 'even' && isDivisible) {
      return filteredTargets.map(target => ({ target, duration: { value: perTargetEven, units: unit } }));
    }
    let perTarget = Math.floor(total / n / increment) * increment;
    let assigned = Array(n).fill(perTarget);
    let assignedTotal = perTarget * n;
    let remaining = total - assignedTotal;
    for (let i = 0; i < n && remaining >= increment; i++) {
      assigned[i] += increment;
      remaining -= increment;
      assignedTotal += increment;
    }

    return new Promise(resolve => {
      let applied = false;
      let content = `<p>Total available duration: <b>${total} ${unit || ''}</b></p>`;
      content += `<div class="target-selection-container" style="max-height: 400px; overflow-y: auto; border: 1px solid #ccc; border-radius: 5px; padding: 10px; margin-top: 10px;">`;
      content += `<div style="display: flex; flex-wrap: wrap; gap: 10px;">`;
      filteredTargets.forEach((target, index) => {
        const tokenName = target.name || target.actor.name;
        const tokenImg = target.document?.texture?.src || target.texture?.src;
        content += `
          <div class="target-option" style="display: flex; flex-direction: column; align-items: center; width: 120px;">
            <div style="font-weight: bold; margin-bottom: 2px;">
              <span id="duration-${index}">${assigned[index]}</span> ${unit || ''}
            </div>
            <div style="display: flex; flex-direction: row; align-items: center; margin-bottom: 2px;">
              <button type="button" class="communal-down" data-index="${index}" style="width: 24px; height: 24px;">-</button>
              <button type="button" class="communal-up" data-index="${index}" style="width: 24px; height: 24px; margin-left: 4px;">+</button>
            </div>
            <img src="${tokenImg}" style="width: 64px; height: 64px; border: 2px solid #888; border-radius: 5px;" />
            <label style="margin-bottom: 3px;">${tokenName}</label>
          </div>
        `;
      });
      content += `</div></div>`;
      content += `<div style="margin-top: 10px;">Unassigned duration: <b><span id="unassigned">${total - assigned.reduce((a, b) => a + b, 0)}</span> ${unit || ''}</b></div>`;

      const dialog = new Dialog({
        title: game.i18n.localize('NAS.buffs.SelectBuffTargets'),
        content: content,
        buttons: {
          apply: {
            icon: '<i class="fas fa-check"></i>',
            label: game.i18n.localize('NAS.buffs.ApplyBuff'),
            callback: html => {
              applied = true;
              resolve(filteredTargets.map((t, i) => ({ target: t, duration: { value: assigned[i], units: unit } })));
            }
          },
          cancel: {
            icon: '<i class="fas fa-times"></i>',
            label: game.i18n.localize('NAS.common.buttons.cancel'),
            callback: () => {
              action.shared.reject = true;
              resolve([]);
            }
          }
        },
        default: "apply",
        close: () => {
          if (!applied) {
            action.shared.reject = true;
            resolve([]);
          }
        }
      }, { width: Math.max(400, n * 140) });

      dialog.render(true);
      Hooks.once('renderDialog', (app, html) => {
        html.find('.communal-up').on('click', function() {
          const idx = Number(this.dataset.index);
          if ((assigned.reduce((a, b) => a + b, 0) + increment) <= total) {
            assigned[idx] += increment;
            html.find(`#duration-${idx}`).text(assigned[idx]);
            html.find('#unassigned').text(total - assigned.reduce((a, b) => a + b, 0));
          }
        });
        html.find('.communal-down').on('click', function() {
          const idx = Number(this.dataset.index);
          if (assigned[idx] - increment >= 0) {
            assigned[idx] -= increment;
            html.find(`#duration-${idx}`).text(assigned[idx]);
            html.find('#unassigned').text(total - assigned.reduce((a, b) => a + b, 0));
          }
        });
      });
    });
  }

  return new Promise(resolve => {
    let applied = false;
    const spellName = action.item.name;
    let content = `<p>${game.i18n.format('NAS.buffs.SelectTargets', { name: spellName })}</p>`;
    
    content += `<div class="target-selection-container" style="max-height: 400px; overflow-y: auto; border: 1px solid #ccc; border-radius: 5px; padding: 10px; margin-top: 10px;">`;
    content += `<div style="display: flex; flex-wrap: wrap; gap: 10px;">`;
    
    filteredTargets.forEach((target, index) => {
      const tokenName = target.name || target.actor.name;
      const tokenImg = target.document?.texture?.src || target.texture?.src;
      const targetDisposition = target.document?.disposition || target?.disposition;
      const actionDisposition = action.token?.disposition;
      const isSameDisposition = targetDisposition === actionDisposition;
      
      let dispositionName = "Unknown";
      if (targetDisposition === CONST.TOKEN_DISPOSITIONS.NEUTRAL) dispositionName = "Neutral";
      else if (targetDisposition === CONST.TOKEN_DISPOSITIONS.FRIENDLY) dispositionName = "Friendly";
      else if (targetDisposition === CONST.TOKEN_DISPOSITIONS.HOSTILE) dispositionName = "Hostile";
      else if (targetDisposition === CONST.TOKEN_DISPOSITIONS.SECRET) dispositionName = "Secret";
      
      content += `
        <div class="target-option" style="display: flex; flex-direction: column; align-items: center; width: 100px;">
          <img src="${tokenImg}" style="width: 64px; height: 64px; border: 2px solid ${isSameDisposition ? 'green' : 'red'}; border-radius: 5px;" />
          <input type="checkbox" id="target-${index}" name="target-${index}" checked style="margin: 6px 0 3px 0;" />
          <label for="target-${index}" style="margin-bottom: 3px;">${tokenName}</label>
          <div style="font-size: 0.8em; color: ${isSameDisposition ? 'green' : 'red'};">${dispositionName}</div>
        </div>
      `;
    });
    
    content += `</div></div>`;
    
    const dialog = new Dialog({
      title: game.i18n.localize('NAS.buffs.SelectBuffTargets'),
      content: content,
      buttons: {
        apply: {
          icon: '<i class="fas fa-check"></i>',
          label: game.i18n.localize('NAS.buffs.ApplyBuff'),
          callback: html => {
            applied = true;
            const selectedTargets = [];
            filteredTargets.forEach((target, index) => {
              let isChecked;
              if (typeof html.find === 'function') {
                isChecked = html.find(`#target-${index}`).prop('checked');
              } else {
                const checkbox = html.querySelector(`#target-${index}`);
                isChecked = checkbox ? checkbox.checked : false;
              }
              if (isChecked) {
                selectedTargets.push(target);
              }
            });
            resolve(selectedTargets);
          }
        },
        cancel: {
          icon: '<i class="fas fa-times"></i>',
          label: game.i18n.localize('NAS.common.buttons.cancel'),
          callback: () => {
            action.shared.reject = true;
            resolve([]);
          }
        }
      },
      default: "apply",
      close: () => {
        if (!applied) {
          action.shared.reject = true;
          resolve([]);
        }
      }
    });
    
    dialog.render(true);
  });
}

const variantMappingSetting = 'pairedBuffMappings';

function toVariantRef(buff) {
  if (!buff) return null;
  return { id: buff.id, pack: buff.pack || null, name: buff.name };
}

function getSpellKey(action) {
  const uuid = action?.item?.uuid;
  if (uuid) return uuid;
  const id = action?.item?.id;
  if (id) return `item:${id}`;
  const name = action?.item?.name || '';
  return `name:${name.toLowerCase()}`;
}

function splitTargetsByDisposition(casterToken, targets) {
  const casterDisposition = casterToken?.document?.disposition ?? casterToken?.disposition;
  const allies = [];
  const foes = [];
  for (const t of targets) {
    const d = t?.document?.disposition ?? t?.disposition;
    if (d === undefined || d === null) continue;
    if (d === CONST.TOKEN_DISPOSITIONS.NEUTRAL || d === CONST.TOKEN_DISPOSITIONS.SECRET) continue;
    if (d === casterDisposition) allies.push(t);
    else foes.push(t);
  }
  return { allies, foes };
}

function durationForTarget(target, perTargetDurations, fallbackDuration) {
  const entry = perTargetDurations?.find?.(d => d.target?.id === target.id);
  if (entry) {
    return {
      units: entry.duration.units,
      value: String(entry.duration.value ?? entry.duration)
    };
  }
  return fallbackDuration;
}

async function storeVariantMapping(spellKey, plan, buffs) {
  const mappings = game.settings.get(MODULE.ID, variantMappingSetting) || {};
  const toStore = { ...mappings };
  const mapEntry = {
    allies: plan.allies ? toVariantRef(plan.allies) : null,
    foes: plan.foes ? toVariantRef(plan.foes) : null,
    allowSwitching: !!plan.allowSwitching,
    applyAllies: plan.applyAllies !== false,
    applyFoes: plan.applyFoes !== false,
    perTarget: (plan.perTarget || []).map(pt => ({
      actorId: pt.actorId,
      tokenId: pt.targetId,
      variantIndex: pt.variantIndex,
      applyTiming: pt.applyTiming || (pt.applyOnTurn ? 'turn' : 'cast')
    }))
  };
  toStore[spellKey] = mapEntry;
  await game.settings.set(MODULE.ID, variantMappingSetting, toStore);
}

async function handleVariantPlanApplication({ action, variants, plan, targetContext, durationUnits, durationValue, casterLevel }) {
  const { filteredTargets, perTargetDurations } = targetContext;
  if (!filteredTargets?.length) return;

  const defaultDuration = { units: durationUnits, value: String(durationValue) };
  const perTargetMap = new Map((plan.perTarget || []).map(pt => [pt.targetId, pt]));
  const { allies, foes } = splitTargetsByDisposition(action.token, filteredTargets);
  const applyAllies = plan.applyAllies !== false;
  const applyFoes = plan.applyFoes !== false;

  if (plan.remember) {
    await storeVariantMapping(getSpellKey(action), plan, variants);
  }

  const immediateBuckets = new Map();
  const combat = game.combat;
  const scheduled = [];

  for (const target of filteredTargets) {
    const isAlly = allies.includes(target);
    if (plan.allowSwitching) {
      if (isAlly && !applyAllies) continue;
      if (!isAlly && !applyFoes) continue;
    }
    const assignment = perTargetMap.get(target.id);
    const variant = assignment ? variants[assignment.variantIndex] : (allies.includes(target) ? plan.allies : plan.foes);
    if (!variant) continue;

    const duration = durationForTarget(target, perTargetDurations, defaultDuration);

    if (plan.allowSwitching) {
      await ensureVariantsOnTarget(target, variants, duration, casterLevel, { silent: true });
      const applyTiming = assignment?.applyTiming || 'cast';
      if (applyTiming !== 'turn') {
        await activateVariantForTarget(target, variant, variants, duration, casterLevel, { silent: true });
      }
      if (combat) {
        const timing = computeApplyTiming(combat, target.id);
        if (timing) {
          scheduled.push({
            tokenId: target.id,
            actorId: target.actor?.id,
            variantIndex: variants.findIndex(v => v === variant),
            duration,
            switching: true,
            applyTiming: applyTiming || 'cast',
            turnIndex: timing.turnIndex,
            applyRound: timing.round,
            applyTurn: timing.turn
          });
        }
      }
    } else {
      const key = `${variant.id}|${variant.pack || 'world'}`;
      if (!immediateBuckets.has(key)) immediateBuckets.set(key, { variant, targets: [] });
      immediateBuckets.get(key).targets.push(target);
    }
  }

  for (const bucket of immediateBuckets.values()) {
    await applyBuffToTargets(bucket.variant, bucket.targets, defaultDuration, casterLevel);
  }

  if (plan.allowSwitching && scheduled.length > 0 && combat) {
    await queueVariantTracker({
      combat,
      action,
      variants,
      scheduled,
      casterLevel,
      defaultDuration
    });
  }
}

function computeApplyTiming(combat, tokenId) {
  if (!combat || !combat.turns) return null;
  const turnIndex = combat.turns.findIndex(t => t?.token?.id === tokenId);
  if (turnIndex === -1) return null;
  const currentRound = combat.round;
  const currentTurn = combat.turn;
  const applyRound = turnIndex > currentTurn ? currentRound : currentRound + 1;
  return { round: applyRound, turn: turnIndex, turnIndex };
}

async function queueVariantTracker({ combat, action, variants, scheduled, casterLevel, defaultDuration }) {
  const tracker = combat.getFlag(MODULE.ID, "variantBuffTracker") || [];
  tracker.push({
    spellKey: getSpellKey(action),
    spellName: action.item?.name,
    variants: variants.map(toVariantRef),
    caster: {
      tokenId: action.token?.id,
      actorId: action.token?.actor?.id,
      round: combat.round,
      turn: combat.turn,
      level: casterLevel
    },
    defaultDuration,
    targets: scheduled
  });
  await combat.setFlag(MODULE.ID, "variantBuffTracker", tracker);
}

export async function resolveBuffReference(ref) {
  if (!ref) return null;
  try {
    if (!ref.pack) {
      const item = game.items.get(ref.id);
      if (item && item.type === 'buff') {
        return { name: item.name, id: item.id, pack: null, document: item };
      }
      return null;
    }
    const pack = game.packs.get(ref.pack);
    if (!pack) return null;
    const doc = await pack.getDocument(ref.id);
    if (!doc || doc.type !== 'buff') return null;
    return { name: doc.name, id: doc.id, pack: ref.pack, document: doc };
  } catch (err) {
    console.error(`${MODULE.ID} | resolveBuffReference failed`, err);
    return null;
  }
}

async function ensureVariantsOnTarget(target, variants, duration, casterLevel, options = {}) {
  const ensurePromises = [];
  for (const variant of variants) {
    if (!variant) continue;
    ensurePromises.push(applyBuffToTargets(variant, [target], duration, casterLevel, { activate: false, silent: options.silent }));
  }
  await Promise.all(ensurePromises);
}

export async function activateVariantForTarget(target, variant, variants, duration, casterLevel, options = {}) {
  if (!variant) return;
  await applyBuffToTargets(variant, [target], duration, casterLevel, { activate: true, silent: options.silent });
  const actor = target.actor;
  if (!actor) return;
  for (const other of variants) {
    if (!other) continue;
    if (other.name === variant.name && (other.pack || null) === (variant.pack || null)) continue;
    const existing = actor.items.find(item => item.type === "buff" && item.name === other.name);
    if (existing?.system?.active) {
      await existing.update({ "system.active": false });
    }
  }
}

/**
 * Apply a buff to appropriate targets
 * @param {Object} buff - The buff item to apply
 * @param {Array} targets - Array of target tokens
 * @param {Object} duration - The duration information for the buff
 * @param {number} casterLevel - The caster level of the spell
 * @param {Object} options - { activate?: boolean, silent?: boolean }
 * @returns {Promise<void>}
 */
export async function applyBuffToTargets(buff, targets, duration, casterLevel, options = {}) {
  const activate = options.activate !== false;
  const silent = !!options.silent;
  if (!game.user.isGM) {
    await socket.executeAsGM(
      "applyBuffToTargetsSocket",
      { name: buff.name, id: buff.id, pack: buff.pack },
      targets.map(t => t.id),
      duration,
      casterLevel,
      { activate, silent }
    );
  }
  
  if (!buff || !targets || targets.length === 0) {
    console.warn(`${MODULE.ID} | Cannot apply buff: Invalid buff or no targets`);
    return;
  }
  
  for (const target of targets) {
    try {
      const actor = target.actor;
      if (!actor) {
        console.warn(`${MODULE.ID} | Target has no actor, skipping buff application`);
        continue;
      }
      const nameMatches = actor.items.filter(item => item.type === "buff" && item.name === buff.name);

      let existingBuff = null;
      if (buff.pack) {
        existingBuff = nameMatches.find(item => {
          const source = item.flags?.[MODULE.ID]?.sourceId || item.flags?.core?.sourceId || item._stats?.compendiumSource;
          if (!source || !source.startsWith("Compendium.")) return false;

          const parts = source.split('.');

          const itemIndex = parts.findIndex(part => part === "Item");
          if (itemIndex > 1) {
            const sourcePackId = parts.slice(1, itemIndex).join('.');
            return sourcePackId === buff.pack;
          }
          return false;
        });
      }

      if (!existingBuff && nameMatches.length > 0) {
        existingBuff = nameMatches[0];
      }
      
      if (existingBuff) {
        if (!existingBuff.flags?.[MODULE.ID]?.sourceId && buff.document?.uuid) {
          await existingBuff.update({ [`flags.${MODULE.ID}.sourceId`]: buff.document.uuid });
        }
        
        const isActive = existingBuff.isActive;
        if (isActive && !activate) {
          await existingBuff.update({
            "system.duration.units": duration.units,
            "system.duration.value": String(duration.value),
            "system.active": false,
            ...(casterLevel !== undefined ? { "system.level": casterLevel } : {})
          });
          continue;
        }

        await existingBuff.update({
          "system.duration.units": duration.units,
          "system.duration.value": String(duration.value),
          "system.active": activate,
          ...(casterLevel !== undefined ? { "system.level": casterLevel } : {})
        });
        
        if (!silent) ui.notifications.info(game.i18n.format('NAS.buffs.UpdatedExisting', { name: buff.name, actor: actor.name }));
      } else {
        let buffData;
        if (typeof Item?.implementation?.fromCompendium === "function") {
          buffData = await Item.implementation.fromCompendium(buff.document);
        } else if (typeof Item?.fromCompendium === "function") {
          buffData = await Item.fromCompendium(buff.document);
        } else {
          buffData = buff.document.toObject();
        }
        buffData.flags = buffData.flags || {};
        buffData.flags[MODULE.ID] = buffData.flags[MODULE.ID] || {};
        if (!buffData.flags[MODULE.ID].sourceId) {
          buffData.flags[MODULE.ID].sourceId = buff.document.uuid;
        }
        
        if (duration && duration.units) {
          buffData.system = buffData.system || {};
          buffData.system.duration = buffData.system.duration || {};
          buffData.system.duration.units = duration.units;
          buffData.system.duration.value = String(duration.value); 
        }
        
        if (casterLevel !== undefined) {
          buffData.system = buffData.system || {};
          buffData.system.level = casterLevel;
        }
        
        buffData.system = buffData.system || {};
        buffData.system.active = activate;
        
        const newItems = await actor.createEmbeddedDocuments("Item", [buffData]);
        
        if (newItems && newItems.length > 0 && activate) {
          const newBuff = newItems[0];
          await newBuff.update({"system.active": true});
        }
        
        if (!silent) ui.notifications.info(game.i18n.format('NAS.buffs.Applied', { name: buff.name, actor: actor.name }));
      }
    } catch (error) {
      console.error(`${MODULE.ID} | Error applying buff to target:`, error);
      if (!silent) ui.notifications.error(game.i18n.format('NAS.buffs.FailedToApply', { name: buff.name, error: error.message }));
    }
  }
}

function checkAndConsumeSpellSlots({ action, filteredTargets, isCommunal, isAreaOfEffect }) {
  const scalableInfo = (!isCommunal && !isAreaOfEffect && filteredTargets.length > 1)
    ? evaluateScalableTargetAllowance(action, filteredTargets)
    : null;
  if (scalableInfo?.allowSingleCast) {
    action.shared.rollData.chargeCost = scalableInfo.originalCost;
    return {
      spellbook: null,
      spellLevel: null,
      spellLevelKey: null,
      originalCost: scalableInfo.originalCost,
      totalCost: scalableInfo.originalCost,
      parsedTargetAllowance: scalableInfo.parsedTargetAllowance
    };
  }

  if (
    action.item.type === "spell" &&
    !isCommunal &&
    filteredTargets.length > 1 &&
    !isAreaOfEffect 
  ) {
    const numTargets = filteredTargets.length;
    const spellbook = action.item.system.spellbook;
    const baseSpellLevel = action.item.system.level;
    const actor = action.token?.actor;

    const spellbookData = actor?.system?.attributes?.spells?.spellbooks?.[spellbook];
    const slotIncrease = Number(action.shared?.nasSpellContext?.metamagic?.slotIncrease ?? 0);
    const isSpontaneous = Boolean(spellbookData?.spontaneous);
    const targetSpellLevel = isSpontaneous ? baseSpellLevel + slotIncrease : baseSpellLevel;
    const spellLevelKey = `spell${targetSpellLevel}`;
    const spellLevelData = spellbookData?.spells?.[spellLevelKey];

    let maxSlots, remainingSlots, usedSlots;
    if (spellbookData?.prepared && !spellbookData?.spontaneous) {
      maxSlots = action.item.system?.preparation?.max ?? 0;
      remainingSlots = action.item.system?.preparation?.value ?? 0;
      usedSlots = maxSlots - remainingSlots;
    } else if (spellbookData?.spontaneous) {
      maxSlots = spellLevelData.max ?? 0;
      remainingSlots = spellLevelData.value ?? 0;
      usedSlots = maxSlots - remainingSlots;
    } else {
      maxSlots = spellLevelData.max ?? 0;
      remainingSlots = spellLevelData.value ?? 0;
      usedSlots = maxSlots - remainingSlots;
    }
    let originalCost = 1;
    const costStr = action.item.system?.uses?.autoDeductChargesCost;
    if (typeof costStr === 'string' && costStr.trim() !== '') {
      const parsed = parseInt(costStr, 10);
      if (!isNaN(parsed) && parsed > 0) originalCost = parsed;
    }
    const totalCost = originalCost * numTargets;
    if (remainingSlots < totalCost) {
      action.shared.reject = true;
      ui.notifications.warn(
        game.i18n.format("NAS.buffs.NotEnoughSpellSlots", {
          remaining: remainingSlots,
          needed: totalCost
        })
      );
      return { rejected: true };
    }
    if (typeof action.shared.rollData.chargeCost === 'number') {
      action.shared.rollData.chargeCost = totalCost;
    } else {
      action.shared.rollData.chargeCost = totalCost;
    }
    return { spellbook, spellLevel: targetSpellLevel, spellLevelKey, originalCost, totalCost };
  }

  if (
    action.item.type === "spell" &&
    !isCommunal &&
    !isAreaOfEffect &&
    filteredTargets.length > 1
  ) {
    const targetText = action.action?.target?.value
      || action.item?.system?.actions?.[0]?.target?.value
      || action.item?.system?.target?.value;
    const casterLevel = action.shared.rollData?.cl ?? action.item?.system?.level ?? 0;
    const parsedCount = estimateScalableTargets(targetText, casterLevel);
    if (parsedCount && parsedCount >= filteredTargets.length) {
      let originalCost = 1;
      const costStr = action.item.system?.uses?.autoDeductChargesCost;
      if (typeof costStr === 'string' && costStr.trim() !== '') {
        const parsed = parseInt(costStr, 10);
        if (!isNaN(parsed) && parsed > 0) originalCost = parsed;
      }
      action.shared.rollData.chargeCost = originalCost;
      return { spellbook: null, spellLevel: null, spellLevelKey: null, originalCost, totalCost: originalCost, parsedTargetAllowance: parsedCount };
    }
  }
  return {};
}

function evaluateScalableTargetAllowance(action, filteredTargets) {
  const forceScalable = action.item?.getFlag?.(MODULE.ID, 'scalableTargets') === true;
  const targetText = action.action?.target?.value
    || action.item?.system?.actions?.[0]?.target?.value
    || action.item?.system?.target?.value;
  const casterLevel = action.shared.rollData?.cl ?? action.item?.system?.level ?? 0;
  const parsedCount = estimateScalableTargets(targetText, casterLevel);
  if (!forceScalable && (!parsedCount || parsedCount < filteredTargets.length)) {
    return null;
  }

  let originalCost = 1;
  const costStr = action.item.system?.uses?.autoDeductChargesCost;
  if (typeof costStr === 'string' && costStr.trim() !== '') {
    const parsed = parseInt(costStr, 10);
    if (!isNaN(parsed) && parsed > 0) originalCost = parsed;
  }
  return {
    allowSingleCast: true,
    originalCost,
    parsedTargetAllowance: parsedCount || filteredTargets.length
  };
}

function estimateScalableTargets(rawTarget, casterLevel) {
  if (!rawTarget || typeof rawTarget !== 'string') return null;
  const WORD_NUM = { one:1, two:2, three:3, four:4, five:5, six:6, seven:7, eight:8, nine:9, ten:10 };
  const normalize = (s) => s
    .toLowerCase()
    .replace(/caster levels?/g, (m) => (m.endsWith('s') ? 'levels' : 'level'))
    .replace(/-/g, ' ')
    .replace(/[“”"’']/g, '')
    .replace(/[.,;()]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/\b(one|two|three|four|five|six|seven|eight|nine|ten)\b/g, (m) => String(WORD_NUM[m]))
    .trim();

  const s = normalize(rawTarget);
  if (!s) return null;

  const capMatch =
    s.match(/\bmaximum of (\d+)\b/) ||
    s.match(/\bto a maximum of (\d+)\b/) ||
    s.match(/\bmaximum (\d+)\b/) ||
    s.match(/\bmax(?:imum)?\s*(\d+)\b/);
  const cap = capMatch ? Number(capMatch[1]) : null;
  const applyCap = (n) => (Number.isFinite(cap) ? Math.min(n, cap) : n);

  const TYPE_IN_PARSE = "(?:creature|object|animal|construct|humanoid|ally|enemy|target|targets)";

  const countFormula = ({ base = 0, mult = 1, divisor = 1, mode = "floor" }) => {
    const scaled = mode === "linear" ? (casterLevel / divisor) * mult : Math.floor(casterLevel / divisor) * mult;
    return Math.max(1, applyCap(base + scaled));
  };

  let m = s.match(new RegExp(
    `^you\\s+(?:plus|and)\\s+(?:up to\\s+)?(\\d+)\\s+.*?${TYPE_IN_PARSE}.*?(?:\\/|per)\\s*level\\b`
  ));
  if (m) return countFormula({ base: 1, mult: Number(m[1]), divisor: 1, mode: "linear" });

  m = s.match(new RegExp(
    `^you\\s+(?:plus|and)\\s+(?:up to\\s+)?(\\d+)\\s+.*?${TYPE_IN_PARSE}.*?(?:\\/|per)\\s*(\\d+)\\s*levels?\\b`
  ));
  if (m) return countFormula({ base: 1, mult: Number(m[1]), divisor: Number(m[2]), mode: "floor" });

  m = s.match(/\bup to (\d+)\s*(?:\/|per)\s*level\b/);
  if (m) return countFormula({ base: 0, mult: Number(m[1]), divisor: 1, mode: "linear" });

  m = s.match(/^(?:up to\s+)?(\d+)\s+.*?(?:\/|per)\s*level\b/);
  if (m) return countFormula({ base: 0, mult: Number(m[1]), divisor: 1, mode: "linear" });

  m = s.match(/^(?:up to\s+)?(\d+)\s+.*?(?:\/|per)\s*(\d+)\s*levels?\b/);
  if (m) return countFormula({ base: 0, mult: Number(m[1]), divisor: Number(m[2]), mode: "floor" });

  m = s.match(new RegExp(`\\b(\\d+)\\s+.*?${TYPE_IN_PARSE}.*?(?:\\/|per)\\s*level\\b`));
  if (m) return countFormula({ base: 0, mult: Number(m[1]), divisor: 1, mode: "linear" });

  return null;
}

async function handleCommunalDuration({
  isCommunal,
  filteredTargets,
  durationUnits,
  durationValue,
  communalIncrement,
  communalTotalDuration,
  communalDurationUnit,
  action
}) {
  if (isCommunal && filteredTargets && filteredTargets.length > 0) {
    const communalHandling = game.settings.get(MODULE.ID, 'communalHandling');
    const n = filteredTargets.length;
    if ((durationUnits === 'hour' || durationUnits === 'hours') && Number(durationValue) === 24) {
      const increment = 1;
      const total = 24;
      if (communalHandling === 'prompt') {
        const communalResult = await promptTargetSelection(filteredTargets, action, {
          communal: true,
          increment,
          total,
          unit: durationUnits
        });
        if (communalResult.length > 0 && communalResult[0].target && communalResult[0].duration !== undefined) {
          return communalResult;
        } else {
          return null;
        }
      } else if (communalHandling === 'even') {
        const perTarget = Math.floor(total / n);
        const assignedTotal = perTarget * n;
        if (assignedTotal === total && perTarget > 0) {
          return filteredTargets.map(target => ({ target, duration: { value: perTarget, units: durationUnits } }));
        } else {
          const communalResult = await promptTargetSelection(filteredTargets, action, {
            communal: true,
            increment,
            total,
            unit: durationUnits
          });
          if (communalResult.length > 0 && communalResult[0].target && communalResult[0].duration !== undefined) {
            return communalResult;
          } else {
            return null;
          }
        }
      }
    } else if (communalIncrement && communalTotalDuration) {
      if (communalHandling === 'prompt') {
        const communalResult = await promptTargetSelection(filteredTargets, action, {
          communal: true,
          increment: communalIncrement,
          total: communalTotalDuration,
          unit: communalDurationUnit || durationUnits
        });
        if (communalResult.length > 0 && communalResult[0].target && communalResult[0].duration !== undefined) {
          return communalResult;
        } else {
          return null;
        }
      } else if (communalHandling === 'even') {
        const perTarget = Math.floor(communalTotalDuration / n / communalIncrement) * communalIncrement;
        const assignedTotal = perTarget * n;
        if (assignedTotal === communalTotalDuration && perTarget > 0) {
          return filteredTargets.map(target => ({ target, duration: { value: perTarget, units: communalDurationUnit || durationUnits } }));
        } else {
          const communalResult = await promptTargetSelection(filteredTargets, action, {
            communal: true,
            increment: communalIncrement,
            total: communalTotalDuration,
            unit: communalDurationUnit || durationUnits
          });
          if (communalResult.length > 0 && communalResult[0].target && communalResult[0].duration !== undefined) {
            return communalResult;
          } else {
            return null;
          }
        }
      }
    }
  }
  return null;
}

async function parseCommunalDuration({ action, durationUnits, rawDurationValue, casterLevel }) {
  const normalizeUnit = (unit) => {
    if (!unit) return null;
    const lower = unit.toString().toLowerCase();
    if (lower.startsWith('hour')) return 'hour';
    if (lower.startsWith('min')) return 'minute';
    if (lower.startsWith('day')) return 'day';
    if (lower.startsWith('round')) return 'round';
    return unit;
  };

  const setUnitFromText = (text, currentUnit) => {
    if (currentUnit) return currentUnit;
    if (/min/i.test(text)) return 'minute';
    if (/hour|hr/i.test(text)) return 'hour';
    if (/day/i.test(text)) return 'day';
    if (/round/i.test(text)) return 'round';
    return null;
  };

  const unit = normalizeUnit(durationUnits);
  const formulaStr = typeof rawDurationValue === 'string' ? rawDurationValue.trim() : '';
  const rollData = action.shared?.rollData ?? {};

  let increment = null;
  let totalDuration = null;
  let derivedUnit = unit;

  const numericValue = Number(rawDurationValue);
  if (!Number.isNaN(numericValue) && numericValue > 0) {
    increment = 1;
    totalDuration = numericValue;
    return { increment, totalDuration, unit: derivedUnit, formula: rawDurationValue };
  }

  if (formulaStr) {
    const perLevelMatch = formulaStr.match(/(\d+)\s*(min\.?|minute|hr|hour|day|round)s?\.?\s*\/\s*level/i);
    if (perLevelMatch) {
      increment = parseInt(perLevelMatch[1], 10);
      derivedUnit = setUnitFromText(perLevelMatch[0], derivedUnit) || derivedUnit;
      totalDuration = increment * (casterLevel || 0);
      return { increment, totalDuration, unit: derivedUnit, formula: rawDurationValue };
    }

    const multiplierMatch = formulaStr.match(/(\d+)\s*\*\s*@cl/i);
    if (multiplierMatch) {
      increment = parseInt(multiplierMatch[1], 10);
      totalDuration = increment * (casterLevel || 0);
    }

    if (increment === null && /@cl/i.test(formulaStr)) {
      increment = 1;
      totalDuration = (casterLevel || 0);
    }

    const formulaLooksRollable = !/\/level/i.test(formulaStr) && !/until discharged/i.test(formulaStr);
    if (formulaLooksRollable) {
      try {
        const evaluated = await new Roll(formulaStr, rollData).evaluate();
        if (evaluated?.total !== undefined) {
          totalDuration = evaluated.total;
        }
      } catch (err) {
        console.warn(`${MODULE.ID} | parseCommunalDuration: Failed to evaluate formula "${formulaStr}"`, err);
      }
    }

    if (increment === null) {
      const fallbackMultiplier = formulaStr.match(/(\d+)\s*\*\s*@cl/i);
      if (fallbackMultiplier) {
        increment = parseInt(fallbackMultiplier[1], 10);
      } else if (/@cl/i.test(formulaStr)) {
        increment = 1;
      }
    }

    if (increment === null && totalDuration !== null) {
      increment = 1;
    }

    if (increment !== null || totalDuration !== null) {
      return {
        increment,
        totalDuration,
        unit: derivedUnit,
        formula: rawDurationValue
      };
    }
  }

  return null;
}




