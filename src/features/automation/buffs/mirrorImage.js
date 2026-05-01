import { MODULE } from "../../../common/module.js";
import { elementFromHtmlLike } from "../../../common/foundryCompat.js";
import { socket } from "../../../integration/moduleSockets.js";
import { tokenCanSeeToken, tokenDistance } from "../utils/tokenVisibility.js";

export const MIRROR_IMAGE_SPELL_UUID = "Compendium.pf1.spells.Item.4jsss37x0pplib8f";
export const MIRROR_IMAGE_BUFF_UUID = "Compendium.nevelas-automation-suite.Buffs.Item.01rgplbC3MNbskNF";

const MIRROR_IMAGE_FLAG = "mirrorImage";
const MIRROR_IMAGE_BADGE_NAME = "nasMirrorImageCountBadge";
const MIRROR_IMAGE_INLINE_ROLL_FLAG = "nasMirrorImageInlineRoll";
const TOUCH_ATTACK_TYPES = new Set(["msak", "rsak", "twak"]);
const IMAGE_RESULTS_THAT_BLOCK_DAMAGE = new Set(["imageHit", "nearMissImageDestroyed", "missNoImage"]);

function localize(path) {
  return game.i18n.localize(`NAS.buffs.mirrorImage.${path}`);
}

function format(path, data = {}) {
  return game.i18n.format(`NAS.buffs.mirrorImage.${path}`, data);
}

function remainGrammarForImageCount(images) {
  const n = Number(images) || 0;
  const one = n === 1;
  return {
    imageNoun: localize(one ? "counts.imageOne" : "counts.imageMany"),
    remainVerb: localize(one ? "counts.remainOne" : "counts.remainMany")
  };
}

function escHtml(value) {
  return foundry.utils.escapeHTML(String(value ?? ""));
}

function sourceIdForItem(item) {
  return String(
    item?.uuid
    ?? item?.flags?.core?.sourceId
    ?? item?._stats?.compendiumSource
    ?? item?.flags?.[MODULE.ID]?.sourceId
    ?? ""
  ).trim();
}

function itemMatchesMirrorImageSpell(item) {
  if (!item) return false;
  const ids = [
    item.uuid,
    item.flags?.core?.sourceId,
    item._stats?.compendiumSource,
    item.flags?.[MODULE.ID]?.sourceId
  ].map((v) => String(v ?? "").trim()).filter(Boolean);
  if (ids.includes(MIRROR_IMAGE_SPELL_UUID)) return true;
  return String(item.name ?? "").trim().toLowerCase() === "mirror image";
}

function actorCanModify(actor) {
  return Boolean(actor) && (game.user?.isGM || actor.isOwner);
}

function itemCanModify(item) {
  return Boolean(item) && (game.user?.isGM || item.isOwner || item.actor?.isOwner);
}

function actorFromUuidSync(actorUuid) {
  if (!actorUuid || typeof foundry.utils?.fromUuidSync !== "function") return null;
  try {
    const doc = foundry.utils.fromUuidSync(actorUuid);
    return doc?.documentName === "Actor" ? doc : null;
  } catch (_err) {
    return null;
  }
}

function actorFromUuid(actorUuid) {
  return actorFromUuidSync(actorUuid) ?? game.actors?.get?.(String(actorUuid ?? "").replace(/^Actor\./, "")) ?? null;
}

function tokenDocumentUuid(target) {
  return String(target?.document?.uuid ?? target?.uuid ?? target?.token?.uuid ?? "").trim();
}

function tokenObject(target) {
  return target?.object ?? target?.token?.object ?? target?.token ?? target ?? null;
}

function targetActor(target) {
  return target?.actor ?? target?.document?.actor ?? target?.token?.actor ?? null;
}

function tokenForActor(actor, { preferredTokens = [], message = null } = {}) {
  if (!actor) return null;
  for (const token of preferredTokens) {
    const object = tokenObject(token);
    if (object?.actor?.uuid === actor.uuid) return object;
  }

  const speakerTokenId = message?.speaker?.token;
  if (speakerTokenId) {
    const speakerToken = canvas?.tokens?.get?.(speakerTokenId);
    if (speakerToken?.actor?.uuid === actor.uuid) return speakerToken;
  }

  return actor?.token?.object
    ?? actor?.token
    ?? actor?.getActiveTokens?.(true, true)?.[0]
    ?? actor?.getActiveTokens?.()?.[0]
    ?? null;
}

function hasStatus(actor, ids) {
  const statuses = actor?.statuses;
  if (!statuses) return false;
  for (const id of ids) {
    if (statuses.has?.(id)) return true;
  }
  return false;
}

function trueSeeingRange(actor) {
  const range = Number(actor?.system?.traits?.senses?.tr?.total);
  return Number.isFinite(range) ? Math.max(0, range) : 0;
}

function attackerCanBeFooled({ sourceActor, targetActor: defenderActor, sourceToken, targetToken }) {
  if (hasStatus(sourceActor, ["blind", "blinded"])) return false;
  if (hasStatus(defenderActor, ["invisible"])) return false;
  const range = trueSeeingRange(sourceActor);
  if (
    range > 0
    && sourceToken
    && targetToken
    && tokenDistance(sourceToken, targetToken) <= range
    && tokenCanSeeToken(sourceToken, targetToken)
  ) {
    return false;
  }
  return true;
}

function getAttackTotal(attackRoll) {
  const direct = Number(attackRoll?.total);
  if (Number.isFinite(direct)) return Math.floor(direct);
  const d20Total = Number(attackRoll?.d20?.total);
  if (Number.isFinite(d20Total)) return Math.floor(d20Total);
  return null;
}

function isTouchAttack(action) {
  const type = String(action?.actionType ?? "").toLowerCase();
  return TOUCH_ATTACK_TYPES.has(type) || String(action?.range?.units ?? "").toLowerCase() === "touch";
}

function getTargetAc(targetActorDoc, action) {
  if (action?.isCombatManeuver) {
    const cmd = Number(targetActorDoc?.system?.attributes?.cmd?.total);
    return Number.isFinite(cmd) ? Math.floor(cmd) : null;
  }
  const ac = targetActorDoc?.system?.attributes?.ac;
  const value = isTouchAttack(action)
    ? Number(ac?.touch?.total)
    : Number(ac?.normal?.total);
  return Number.isFinite(value) ? Math.floor(value) : null;
}

function normalizeState(raw = {}) {
  const images = Math.max(0, Math.floor(Number(raw.images) || 0));
  const maxImages = Math.max(0, Math.floor(Number(raw.maxImages) || 8));
  const operations = raw.operations && typeof raw.operations === "object" ? raw.operations : {};
  return {
    active: raw.active !== false,
    sourceSpellUuid: String(raw.sourceSpellUuid ?? MIRROR_IMAGE_SPELL_UUID),
    sourceSpellName: String(raw.sourceSpellName ?? "Mirror Image"),
    casterLevel: Math.max(0, Math.floor(Number(raw.casterLevel) || 0)),
    images,
    maxImages,
    createdAt: Number(raw.createdAt) || Date.now(),
    operations
  };
}

function getMirrorImageStateFromBuff(buff) {
  return normalizeState(buff?.flags?.[MODULE.ID]?.[MIRROR_IMAGE_FLAG] ?? {});
}

export function refreshMirrorImageTokenEffects(actor) {
  const tokens = actor?.getActiveTokens?.(true, true) ?? actor?.getActiveTokens?.() ?? [];
  for (const token of tokens) {
    token?.drawEffects?.();
  }
}

export function refreshMirrorImageSceneTokenEffects() {
  for (const token of canvas?.tokens?.placeables ?? []) {
    if (getMirrorImageBuff(token?.actor, { includeInactive: true })) token?.drawEffects?.();
  }
}

export function isMirrorImageBuff(item) {
  if (item?.type !== "buff") return false;
  const flagState = item.flags?.[MODULE.ID]?.[MIRROR_IMAGE_FLAG];
  const sourceId = sourceIdForItem(item);
  const sourceMatch = sourceId === MIRROR_IMAGE_BUFF_UUID;
  const nameMatch = String(item.name ?? "").trim().toLowerCase() === "mirror image";
  return Boolean(flagState || sourceMatch || nameMatch);
}

export function getMirrorImageBuff(actor, { includeInactive = false } = {}) {
  if (!actor?.items) return null;
  const candidates = actor.items.filter?.((item) => {
    if (!isMirrorImageBuff(item)) return false;
    return includeInactive || item.system?.active === true;
  }) ?? [];
  return candidates[0] ?? null;
}

async function getBuffDataFromCompendium() {
  if (typeof fromUuid !== "function") return null;
  try {
    const doc = await fromUuid(MIRROR_IMAGE_BUFF_UUID);
    if (!doc) return null;
    if (doc.pack && typeof Item?.implementation?.fromCompendium === "function") {
      return Item.implementation.fromCompendium(doc);
    }
    if (doc.pack && typeof Item?.fromCompendium === "function") {
      return Item.fromCompendium(doc);
    }
    return doc.toObject();
  } catch (_err) {
    return null;
  }
}

function packKeyFromCompendiumUuid(uuid) {
  const parts = String(uuid ?? "").split(".");
  const itemIndex = parts.findIndex((part) => part === "Item");
  if (parts[0] !== "Compendium" || itemIndex <= 1) return "";
  return parts.slice(1, itemIndex).join(".");
}

async function getBuffReferenceFromCompendium() {
  if (typeof fromUuid !== "function") return null;
  try {
    const document = await fromUuid(MIRROR_IMAGE_BUFF_UUID);
    if (!document) return null;
    return {
      name: document.name,
      id: document.id,
      pack: typeof document.pack === "string" ? document.pack : packKeyFromCompendiumUuid(MIRROR_IMAGE_BUFF_UUID),
      document
    };
  } catch (_err) {
    return null;
  }
}

async function createMirrorImageBuffLocal(actor, state) {
  if (!actorCanModify(actor)) return null;
  const data = await getBuffDataFromCompendium();
  if (!data) return null;
  data.flags ??= {};
  data.flags[MODULE.ID] ??= {};
  data.flags[MODULE.ID].sourceId = MIRROR_IMAGE_BUFF_UUID;
  data.flags[MODULE.ID][MIRROR_IMAGE_FLAG] = normalizeState(state);
  data.system ??= {};
  data.system.active = true;
  const created = await actor.createEmbeddedDocuments("Item", [data]);
  refreshMirrorImageTokenEffects(actor);
  return created?.[0] ?? null;
}

async function updateMirrorImageBuffLocal(buff, state) {
  if (!itemCanModify(buff)) return false;
  const next = normalizeState(state);
  await buff.update({
    "system.active": next.active === true && next.images > 0,
    [`flags.${MODULE.ID}.${MIRROR_IMAGE_FLAG}`]: next,
    [`flags.${MODULE.ID}.sourceId`]: MIRROR_IMAGE_BUFF_UUID
  });
  refreshMirrorImageTokenEffects(buff.actor);
  return true;
}

async function ensureMirrorImageBuff(actor, state) {
  if (!actor) return null;
  let buff = getMirrorImageBuff(actor, { includeInactive: true });
  if (actorCanModify(actor)) {
    if (!buff) buff = await createMirrorImageBuffLocal(actor, state);
    else await updateMirrorImageBuffLocal(buff, state);
    return getMirrorImageBuff(actor, { includeInactive: true }) ?? buff;
  }
  if (socket) {
    await socket.executeAsGM("setMirrorImageBuffStateSocket", actor.uuid, normalizeState(state), true);
  }
  return buff;
}

async function writeMirrorImageState(actor, buff, state) {
  const next = normalizeState(state);
  if (buff && itemCanModify(buff)) return updateMirrorImageBuffLocal(buff, next);
  if (actor && socket) {
    await socket.executeAsGM("setMirrorImageBuffStateSocket", actor.uuid, next, false);
    return true;
  }
  return false;
}

function inferCasterLevel(actionUse) {
  const fromShared = Number(actionUse?.shared?.rollData?.cl);
  if (Number.isFinite(fromShared) && fromShared > 0) return Math.floor(fromShared);
  const fromAction = Number(actionUse?.action?.getRollData?.()?.cl);
  if (Number.isFinite(fromAction) && fromAction > 0) return Math.floor(fromAction);
  const spellbook = actionUse?.item?.system?.spellbook;
  const fromBook = Number(actionUse?.actor?.system?.attributes?.spells?.spellbooks?.[spellbook]?.cl?.total);
  if (Number.isFinite(fromBook) && fromBook > 0) return Math.floor(fromBook);
  return 0;
}

async function rollInitialImages(casterLevel) {
  const roll = await new Roll("1d4").evaluate();
  const clImages = Math.floor(Math.max(0, Number(casterLevel) || 0) / 3);
  return {
    roll,
    images: Math.min(8, Math.max(1, Math.floor(Number(roll.total) || 0) + clImages))
  };
}

function mirrorImageInlineRollFormula(casterLevel) {
  return `min(8, (1d4 + floor(${Math.max(0, Math.floor(Number(casterLevel) || 0))} / 3)))`;
}

function mirrorImageInlineRollText(casterLevel) {
  return `[[${mirrorImageInlineRollFormula(casterLevel)}]]`;
}

function looksLikeMirrorImageEffectNote(text) {
  const value = String(text ?? "");
  return /\[\[.*1d4.*floor\(@cl\s*\/\s*3\).*]]/i.test(value) && /\bimages?\b/i.test(value);
}

function extractInlineRollTotal(html) {
  if (!html) return null;
  const doc = new DOMParser().parseFromString(String(html), "text/html");
  const rolls = doc.querySelectorAll(".inline-roll[data-roll]");
  for (const el of rolls) {
    const raw = el.dataset.roll;
    if (!raw) continue;
    try {
      const json = JSON.parse(unescape(raw));
      const total = Number(json?.total);
      if (Number.isFinite(total)) return Math.floor(total);
    } catch (_err) {
    }
  }
  return null;
}

function mirrorImageCreatedFootnote(actionUse, casterLevel) {
  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? actionUse?.item?.actor ?? null;
  return format("created", {
    actor: actor?.name ?? "",
    images: mirrorImageInlineRollText(casterLevel),
    cl: casterLevel
  });
}

export function prepareMirrorImageEffectNote(chatAttack) {
  const actionUse = chatAttack?.actionUse;
  if (!itemMatchesMirrorImageSpell(actionUse?.item) || !Array.isArray(chatAttack?.effectNotes)) return false;
  const casterLevel = inferCasterLevel(actionUse);
  for (const note of chatAttack.effectNotes) {
    if (looksLikeMirrorImageEffectNote(note?.text)) {
      note.text = mirrorImageCreatedFootnote(actionUse, casterLevel);
      return true;
    }
  }
  return false;
}

export function captureMirrorImageEffectNoteRoll(chatAttack) {
  const actionUse = chatAttack?.actionUse;
  if (!itemMatchesMirrorImageSpell(actionUse?.item)) return;
  const total = extractInlineRollTotal(chatAttack?.effectNotesHTML);
  if (Number.isFinite(total) && total > 0) {
    actionUse.shared ??= {};
    actionUse.shared[MIRROR_IMAGE_INLINE_ROLL_FLAG] = Math.min(8, Math.max(1, Math.floor(total)));
  }
}

async function getMirrorImageDuration(actionUse) {
  const durationContext = actionUse?.shared?.nasSpellContext?.duration;
  const duration = actionUse?.action?.duration ?? actionUse?.item?.system?.duration ?? {};
  const units = durationContext?.units ?? duration?.units ?? "";
  if (durationContext?.evaluated?.total != null) {
    return { units, value: String(durationContext.evaluated.total) };
  }

  const rawValue = durationContext?.value ?? duration?.value ?? "";
  try {
    const roll = await new Roll(rawValue, actionUse?.shared?.rollData ?? {}).evaluate({ async: true });
    return { units, value: String(roll?.total ?? "") };
  } catch (_err) {
    const numericFallback = Number(rawValue);
    return {
      units,
      value: String(Number.isNaN(numericFallback) ? "" : numericFallback)
    };
  }
}

async function applyMirrorImageBuffThroughGenericAutomation(actionUse, casterLevel) {
  const sourceToken = tokenObject(actionUse?.token);
  if (!sourceToken?.actor) return false;

  const buff = await getBuffReferenceFromCompendium();
  if (!buff) return false;

  const duration = await getMirrorImageDuration(actionUse);
  const { applyBuffToTargets } = await import("./buffs.js");
  await applyBuffToTargets(buff, [sourceToken], duration, casterLevel, { silent: true });
  return true;
}

export async function handleMirrorImageCast(actionUse) {
  if (!itemMatchesMirrorImageSpell(actionUse?.item)) return null;
  const actor = actionUse?.actor ?? actionUse?.token?.actor ?? actionUse?.item?.actor ?? null;
  if (!actor) return null;
  const casterLevel = inferCasterLevel(actionUse);
  const footnoteImages = Number(actionUse?.shared?.[MIRROR_IMAGE_INLINE_ROLL_FLAG]);
  const fallback = Number.isFinite(footnoteImages) && footnoteImages > 0 ? null : await rollInitialImages(casterLevel);
  const images = Math.min(8, Math.max(1, Math.floor(footnoteImages || fallback?.images || 1)));
  const state = normalizeState({
    active: true,
    sourceSpellUuid: MIRROR_IMAGE_SPELL_UUID,
    sourceSpellName: actionUse?.item?.name ?? "Mirror Image",
    casterLevel,
    images,
    maxImages: 8,
    createdAt: Date.now(),
    operations: {}
  });
  await applyMirrorImageBuffThroughGenericAutomation(actionUse, casterLevel);
  await ensureMirrorImageBuff(actor, state);
  return state;
}

function getChatFlagOperations(message) {
  const flag = message?.flags?.[MODULE.ID]?.[MIRROR_IMAGE_FLAG] ?? message?.getFlag?.(MODULE.ID, MIRROR_IMAGE_FLAG) ?? {};
  const operations = flag?.operations;
  return operations && typeof operations === "object" ? operations : {};
}

function findChatOperation({ message, attackIndex, targetActorUuid, targetTokenUuid, includeUndone = false }) {
  const operations = getChatFlagOperations(message);
  for (const op of Object.values(operations)) {
    if (!includeUndone && op?.undone === true) continue;
    if (Number(op?.attackIndex) !== Number(attackIndex)) continue;
    if (targetTokenUuid && op?.targetTokenUuid === targetTokenUuid) return op;
    if (targetActorUuid && op?.targetActorUuid === targetActorUuid) return op;
  }
  return null;
}

function operationBlocksDamage(operation) {
  return IMAGE_RESULTS_THAT_BLOCK_DAMAGE.has(String(operation?.result ?? ""));
}

async function rollImageTarget(images) {
  const sides = Math.max(2, Math.floor(Number(images) || 0) + 1);
  const roll = await new Roll(`1d${sides}`).evaluate();
  const total = Math.max(1, Math.floor(Number(roll.total) || 1));
  return {
    formula: roll.formula,
    total,
    realTarget: total === 1
  };
}

async function resolveAgainstMirrorImage({
  sourceActor,
  targetActor: defenderActor,
  sourceToken,
  targetToken,
  targetTokenUuid,
  attackTotal,
  attackIndex,
  action,
  recordFullMiss = false,
  operationId = foundry.utils.randomID(16)
} = {}) {
  if (!sourceActor || !defenderActor || !Number.isFinite(Number(attackTotal))) return null;
  if (!attackerCanBeFooled({ sourceActor, targetActor: defenderActor, sourceToken, targetToken })) return null;
  const buff = getMirrorImageBuff(defenderActor);
  if (!buff) return null;
  const state = getMirrorImageStateFromBuff(buff);
  if (!state.active || state.images <= 0) return null;
  const ac = getTargetAc(defenderActor, action);
  if (!Number.isFinite(ac)) return null;

  const previousImages = state.images;
  let nextImages = previousImages;
  let result = "missNoImage";
  let rollFormula = "";
  let rollTotal = null;

  const total = Math.floor(Number(attackTotal));
  if (total >= ac) {
    const imageRoll = await rollImageTarget(previousImages);
    rollFormula = imageRoll.formula;
    rollTotal = imageRoll.total;
    if (imageRoll.realTarget) {
      result = "realTarget";
    } else {
      result = "imageHit";
      nextImages = Math.max(0, previousImages - 1);
    }
  } else if ((ac - total) >= 1 && (ac - total) <= 5) {
    result = "nearMissImageDestroyed";
    nextImages = Math.max(0, previousImages - 1);
  } else if (!recordFullMiss) {
    return null;
  }

  const operation = {
    operationId,
    attackIndex: Number.isInteger(Number(attackIndex)) ? Number(attackIndex) : 0,
    targetTokenUuid: String(targetTokenUuid ?? ""),
    targetActorUuid: defenderActor.uuid,
    sourceActorUuid: sourceActor.uuid,
    attackTotal: total,
    ac,
    previousImages,
    nextImages,
    result,
    rollFormula,
    rollTotal,
    undone: false,
    timestamp: Date.now()
  };

  state.images = nextImages;
  state.active = nextImages > 0;
  state.operations = {
    ...state.operations,
    [operation.operationId]: operation
  };
  await writeMirrorImageState(defenderActor, buff, state);
  return operation;
}

export async function resolveMirrorImagesForActionUse(actionUse) {
  if (!actionUse?.action?.hasAttack) return [];
  const sourceActor = actionUse.actor ?? actionUse.token?.actor ?? actionUse.item?.actor ?? null;
  if (!sourceActor) return [];
  const sourceToken = tokenObject(actionUse.token);
  const targets = Array.isArray(actionUse.shared?.targets) ? actionUse.shared.targets : [];
  if (!targets.length) return [];
  const chatAttacks = Array.isArray(actionUse.shared?.chatAttacks) ? actionUse.shared.chatAttacks : [];
  const operations = [];

  for (let attackIndex = 0; attackIndex < chatAttacks.length; attackIndex += 1) {
    const chatAttack = chatAttacks[attackIndex];
    if (!chatAttack?.attack) continue;
    const attackTotal = getAttackTotal(chatAttack.attack);
    if (!Number.isFinite(attackTotal)) continue;
    for (const target of targets) {
      const defenderActor = targetActor(target);
      if (!defenderActor) continue;
      const defenderToken = tokenObject(target);
      const op = await resolveAgainstMirrorImage({
        sourceActor,
        targetActor: defenderActor,
        sourceToken,
        targetToken: defenderToken,
        targetTokenUuid: tokenDocumentUuid(target),
        attackTotal,
        attackIndex,
        action: actionUse.action
      });
      if (op) operations.push(op);
    }
  }

  if (operations.length) {
    actionUse.shared.chatData ??= {};
    actionUse.shared.chatData.flags ??= {};
    actionUse.shared.chatData.flags[MODULE.ID] ??= {};
    const existing = actionUse.shared.chatData.flags[MODULE.ID][MIRROR_IMAGE_FLAG] ?? {};
    const existingOps = existing.operations && typeof existing.operations === "object" ? existing.operations : {};
    actionUse.shared.chatData.flags[MODULE.ID][MIRROR_IMAGE_FLAG] = {
      ...existing,
      operations: {
        ...existingOps,
        ...Object.fromEntries(operations.map((op) => [op.operationId, op]))
      }
    };
  }
  return operations;
}

function getMessageAttackRoll(message, attackIndex) {
  const attacks = message?.system?.rolls?.attacks ?? message?.rolls?.attacks ?? message?.systemRolls?.attacks ?? [];
  const attack = attacks?.[Number(attackIndex)];
  return attack?.attack ?? null;
}

function getActionFromOptions(options = {}) {
  if (options?.action) return options.action;
  const item = options?.item ?? options?.message?.item ?? null;
  const actionId = options?.message?.system?.action?.id ?? options?.actionId;
  return item?.actions?.get?.(actionId) ?? null;
}

export async function resolveMirrorImageForApplyDamage({ sourceActor, targetActor: defenderActor, options = {} } = {}) {
  if (!sourceActor || !defenderActor) return null;
  const attackIndex = Number(options?.attackIndex);
  if (!Number.isInteger(attackIndex)) return null;
  const targetActorUuid = defenderActor.uuid;
  const existing = findChatOperation({
    message: options.message,
    attackIndex,
    targetActorUuid
  });
  if (existing) {
    return {
      operation: existing,
      blockDamage: false,
      fromChatFlag: true
    };
  }
  const undoneOperation = findChatOperation({
    message: options.message,
    attackIndex,
    targetActorUuid,
    includeUndone: true
  });

  const attackRoll = getMessageAttackRoll(options.message, attackIndex);
  const attackTotal = getAttackTotal(attackRoll);
  if (!Number.isFinite(attackTotal)) return null;
  const action = getActionFromOptions(options);
  const selectedTokens = Array.from(canvas?.tokens?.controlled ?? []);
  const sourceToken = tokenForActor(sourceActor, { preferredTokens: selectedTokens, message: options.message });
  const defenderToken = tokenForActor(defenderActor, { preferredTokens: selectedTokens });
  const op = await resolveAgainstMirrorImage({
    sourceActor,
    targetActor: defenderActor,
    sourceToken,
    targetToken: defenderToken,
    attackTotal,
    attackIndex,
    action,
    recordFullMiss: true
  });
  if (!op) return null;
  await appendOperationToMessage(options.message, op);
  return {
    operation: op,
    blockDamage: undoneOperation ? false : operationBlocksDamage(op),
    fromChatFlag: false
  };
}

function operationSummaryText(operation) {
  const images = Number(operation?.nextImages) || 0;
  const remainWords = remainGrammarForImageCount(images);
  switch (String(operation?.result ?? "")) {
    case "realTarget":
      return format("results.realTarget", { attack: operation.attackTotal, ac: operation.ac, images, ...remainWords });
    case "imageHit":
      return format("results.imageHit", {
        attack: operation.attackTotal,
        ac: operation.ac,
        images,
        roll: operation.rollTotal,
        formula: operation.rollFormula,
        ...remainWords
      });
    case "nearMissImageDestroyed":
      return format("results.nearMiss", { attack: operation.attackTotal, ac: operation.ac, images, ...remainWords });
    case "missNoImage":
      return format("results.missNoImage", { attack: operation.attackTotal, ac: operation.ac, images });
    default:
      return "";
  }
}

function operationPublicText(operation) {
  switch (String(operation?.result ?? "")) {
    case "realTarget":
      return localize("public.realTarget");
    case "imageHit":
      return localize("public.imageHit");
    case "nearMissImageDestroyed":
      return localize("public.nearMiss");
    default:
      return "";
  }
}

function operationCanUndo(operation) {
  if (operation?.undone === true) return false;
  return Number(operation?.previousImages) !== Number(operation?.nextImages);
}

function operationSensitiveUuid(operation) {
  return String(operation?.targetTokenUuid || operation?.targetActorUuid || "").trim();
}

function buildMirrorImageOperationLine(operation) {
  const publicText = operationPublicText(operation);
  const detailText = operationSummaryText(operation);
  if (!publicText && !detailText) return "";

  const targetName = actorFromUuid(operation.targetActorUuid)?.name ?? game.i18n.localize("NAS.common.labels.target");
  const sensitiveUuid = operationSensitiveUuid(operation);
  const sensitiveAttrs = sensitiveUuid ? ` data-gm-sensitive-uuid="${escHtml(sensitiveUuid)}"` : "";
  const undo = operationCanUndo(operation)
    ? `<button type="button" data-nas-mirror-image-undo="${escHtml(operation.operationId)}" data-actor-uuid="${escHtml(operation.targetActorUuid)}">${escHtml(localize("undo"))}</button>`
    : operation.undone === true
      ? `<em>${escHtml(localize("undone"))}</em>`
      : "";
  const sensitive = detailText
    ? `<span${sensitiveAttrs}> ${escHtml(detailText)} ${undo}</span>`
    : "";
  const targetLabel = sensitiveUuid
    ? `<span${sensitiveAttrs}><strong>${escHtml(targetName)}:</strong> </span>`
    : "";

  return `<li>${targetLabel}${escHtml(publicText)}${sensitive}</li>`;
}

function buildMirrorImageBlockHtml(operations) {
  const lines = operations
    .map((op) => buildMirrorImageOperationLine(op))
    .filter(Boolean)
    .join("");
  if (!lines) return "";
  return `<div class="nas-mirror-image-chat"><div><strong>${escHtml(localize("title"))}</strong></div><ul>${lines}</ul></div>`;
}

function latestRenderableOperations(operations) {
  const latest = new Map();
  for (const op of operations) {
    if (op?.undone === true) continue;
    const attackIndex = Number(op?.attackIndex) || 0;
    const targetKey = operationSensitiveUuid(op) || String(op?.targetActorUuid ?? "");
    const key = `${attackIndex}:${targetKey}`;
    const current = latest.get(key);
    if (!current || Number(op?.timestamp ?? 0) >= Number(current?.timestamp ?? 0)) {
      latest.set(key, op);
    }
  }
  return Array.from(latest.values());
}

function contentWithMirrorImageBlocks(content, operations) {
  if (!operations.length || typeof content !== "string" || !content.trim()) return content;
  const parser = new DOMParser();
  const doc = parser.parseFromString(content, "text/html");
  const byAttack = new Map();
  const affectedAttacks = new Set(operations.map((op) => Number(op?.attackIndex) || 0));
  for (const op of latestRenderableOperations(operations)) {
    const index = Number(op?.attackIndex) || 0;
    if (!byAttack.has(index)) byAttack.set(index, []);
    byAttack.get(index).push(op);
  }

  let changed = false;
  for (const attackIndex of affectedAttacks) {
    const attackEl = doc.body.querySelector(`.chat-attack[data-index="${attackIndex}"]`);
    if (!attackEl) continue;
    const existingBlock = attackEl.querySelector(".nas-mirror-image-chat");
    const ops = byAttack.get(attackIndex) ?? [];
    const blockHtml = buildMirrorImageBlockHtml(ops);
    if (!blockHtml) {
      if (existingBlock) {
        existingBlock.remove();
        changed = true;
      }
      continue;
    }
    if (existingBlock) {
      existingBlock.outerHTML = blockHtml;
      changed = true;
      continue;
    }
    const buttonGroup = attackEl.querySelector(":scope > .card-button-group");
    if (buttonGroup) buttonGroup.insertAdjacentHTML("beforebegin", blockHtml);
    else attackEl.insertAdjacentHTML("beforeend", blockHtml);
    changed = true;
  }
  return changed ? doc.body.innerHTML : content;
}

async function updateChatContentWithMirrorImageOperations(message, operations) {
  if (!message?.update || !Array.isArray(operations)) return;
  const content = String(message?.content ?? message?._source?.content ?? "");
  const nextContent = contentWithMirrorImageBlocks(content, operations);
  if (nextContent === content) return;
  try {
    await message.update({ content: nextContent });
  } catch (_err) {
  }
}

function updateOperationInFlagObject(flag, operationId, patch) {
  const operations = flag?.operations && typeof flag.operations === "object" ? foundry.utils.deepClone(flag.operations) : {};
  if (!operations[operationId]) return flag;
  operations[operationId] = {
    ...operations[operationId],
    ...patch
  };
  return {
    ...(flag ?? {}),
    operations
  };
}

async function markChatOperationUndone(message, operationId) {
  if (!message?.getFlag || !message?.setFlag) return;
  const flag = message.getFlag(MODULE.ID, MIRROR_IMAGE_FLAG) ?? {};
  const next = updateOperationInFlagObject(flag, operationId, { undone: true });
  try {
    await message.setFlag(MODULE.ID, MIRROR_IMAGE_FLAG, next);
    await updateChatContentWithMirrorImageOperations(message, Object.values(next.operations ?? {}));
  } catch (_err) {
  }
}

async function appendOperationToMessage(message, operation) {
  if (!message?.getFlag || !message?.setFlag || !operation?.operationId) return;
  const flag = message.getFlag(MODULE.ID, MIRROR_IMAGE_FLAG) ?? {};
  const operations = flag.operations && typeof flag.operations === "object" ? foundry.utils.deepClone(flag.operations) : {};
  operations[operation.operationId] = operation;
  try {
    await message.setFlag(MODULE.ID, MIRROR_IMAGE_FLAG, {
      ...flag,
      operations
    });
    await updateChatContentWithMirrorImageOperations(message, Object.values(operations));
  } catch (_err) {
  }
}

export async function undoMirrorImageOperation({ actorUuid, operationId, messageId } = {}) {
  const actor = actorFromUuid(actorUuid);
  if (!actor || !operationId) return false;
  const buff = getMirrorImageBuff(actor, { includeInactive: true });
  const state = getMirrorImageStateFromBuff(buff);
  const operation = state.operations?.[operationId];
  if (!buff || !operation || operation.undone === true) return false;

  if (!actorCanModify(actor) && socket) {
    return socket.executeAsGM("undoMirrorImageOperationSocket", actorUuid, operationId, messageId ?? null);
  }

  const previousImages = Math.max(0, Math.floor(Number(operation.previousImages) || 0));
  state.images = previousImages;
  state.active = previousImages > 0;
  state.operations = {
    ...state.operations,
    [operationId]: {
      ...operation,
      undone: true
    }
  };
  await writeMirrorImageState(actor, buff, state);
  const message = messageId ? game.messages?.get?.(messageId) : null;
  await markChatOperationUndone(message, operationId);
  return true;
}

export async function applyMirrorImageStateSocket(actorUuid, state, createIfMissing = false) {
  const actor = actorFromUuid(actorUuid);
  if (!actor) return false;
  const next = normalizeState(state);
  const buff = getMirrorImageBuff(actor, { includeInactive: true });
  if (!buff && createIfMissing) {
    return Boolean(await createMirrorImageBuffLocal(actor, next));
  }
  if (!buff) return false;
  return updateMirrorImageBuffLocal(buff, next);
}

export async function undoMirrorImageOperationSocket(actorUuid, operationId, messageId = null) {
  const actor = actorFromUuid(actorUuid);
  if (!actor || !operationId) return false;
  const buff = getMirrorImageBuff(actor, { includeInactive: true });
  if (!buff) return false;
  const state = getMirrorImageStateFromBuff(buff);
  const operation = state.operations?.[operationId];
  if (!operation || operation.undone === true) return false;
  const previousImages = Math.max(0, Math.floor(Number(operation.previousImages) || 0));
  state.images = previousImages;
  state.active = previousImages > 0;
  state.operations = {
    ...state.operations,
    [operationId]: {
      ...operation,
      undone: true
    }
  };
  await updateMirrorImageBuffLocal(buff, state);
  const message = messageId ? game.messages?.get?.(messageId) : null;
  await markChatOperationUndone(message, operationId);
  return true;
}

export function applyMirrorImageChatContent(message) {
  const operations = Object.values(getChatFlagOperations(message));
  if (!operations.length) return;
  const content = String(message?.content ?? message?._source?.content ?? "");
  const nextContent = contentWithMirrorImageBlocks(content, operations);
  if (nextContent === content) return;
  message.updateSource?.({ content: nextContent });
}

export function renderMirrorImageChatControls(message, htmlLike) {
  const root = elementFromHtmlLike(htmlLike);
  if (!root) return;
  root.querySelectorAll("[data-nas-mirror-image-undo]").forEach((button) => {
    button.addEventListener("click", async (event) => {
      event.preventDefault();
      const btn = event.currentTarget;
      btn.disabled = true;
      const ok = await undoMirrorImageOperation({
        actorUuid: btn.dataset.actorUuid,
        operationId: btn.dataset.nasMirrorImageUndo,
        messageId: message.id
      });
      if (!ok) btn.disabled = false;
    });
  });
}

function clearMirrorImageBadges(token) {
  const effects = token?.effects;
  if (!effects?.children) return;
  const stack = [...effects.children];
  while (stack.length) {
    const child = stack.pop();
    if (!child) continue;
    if (child.name === MIRROR_IMAGE_BADGE_NAME) {
      child.parent?.removeChild?.(child);
      child.destroy?.({ children: true });
      continue;
    }
    if (child.children?.length) stack.push(...child.children);
  }
}

function texturePath(displayObject) {
  const texture = displayObject?.texture;
  return String(
    texture?.baseTexture?.resource?.src
    ?? texture?.baseTexture?.cacheId
    ?? texture?.source?.resource?.src
    ?? texture?.source?.label
    ?? texture?.textureCacheIds?.[0]
    ?? ""
  );
}

function normalizePath(value) {
  return String(value ?? "").replace(/\\/g, "/").toLowerCase();
}

function findMirrorImageEffectIcon(token, buff) {
  const effects = token?.effects;
  if (!effects?.children || !buff?.img) return null;
  const needle = normalizePath(buff.img);
  return effects.children.find((child) => {
    const path = normalizePath(texturePath(child));
    return path && (path === needle || path.endsWith(needle) || needle.endsWith(path));
  }) ?? null;
}

function makeMirrorImageBadge(icon, images) {
  const iconSize = Math.max(16, Math.min(Number(icon?.width) || 32, Number(icon?.height) || 32));
  const text = new PIXI.Text(String(images), {
    fontFamily: "Arial",
    fontSize: Math.max(18, Math.round(iconSize * 0.56)),
    fontWeight: "bold",
    fill: 0xff2020,
    stroke: 0x000000,
    strokeThickness: Math.max(4, Math.round(iconSize * 0.06)),
    align: "center"
  });
  text.name = MIRROR_IMAGE_BADGE_NAME;
  text.anchor.set(0.5);
  text.position.set(iconSize * 0.86, iconSize * 0.22);
  return text;
}

function drawMirrorImageCountBadge(token) {
  clearMirrorImageBadges(token);
  const buff = getMirrorImageBuff(token?.actor);
  const state = getMirrorImageStateFromBuff(buff);
  if (!buff || !state.active || state.images <= 0) return;
  const icon = findMirrorImageEffectIcon(token, buff);
  if (!icon) return;
  const badge = makeMirrorImageBadge(icon, state.images);
  icon.addChild(badge);
}

export function registerMirrorImageTokenEffectBadges() {
  if (!globalThis.libWrapper || !globalThis.Token?.prototype?.drawEffects) return;
  libWrapper.register(
    MODULE.ID,
    "Token.prototype.drawEffects",
    async function (wrapped, ...args) {
      const result = await wrapped.apply(this, args);
      try {
        drawMirrorImageCountBadge(this);
      } catch (_err) {}
      return result;
    },
    "WRAPPER"
  );
}
