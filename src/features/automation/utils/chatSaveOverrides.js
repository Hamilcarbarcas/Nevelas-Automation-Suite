import { MODULE } from "../../../common/module.js";
import { getDazingExtraRoundsForTarget } from "../metamagic/dazingSpell.js";

const SAVE_BUTTON_SELECTOR = 'button[data-action="save"]';
const LISTENER_FLAG = "nasPersistentSaveListener";
const MESSAGE_HOOK_FLAG = "nasPersistentSaveHook";
const SAVE_TOKEN_FLAG = "nasSaveTokenInteraction";
const PERSISTENT_SAVE_QUEUE = [];
let hooksRegistered = false;

function getMessageFromEvent(event) {
  const messageEl = event.target?.closest?.(".chat-message");
  if (!messageEl) return null;
  const messageId = messageEl.dataset?.messageId;
  if (!messageId) return null;
  return game.messages?.get(messageId) ?? null;
}

function getTokenFromTarget(target) {
  if (!target) return null;
  if (target.object) return target.object;
  if (target.document) return target.document;
  return target;
}

function queuePersistentSave(entry) {
  if (!entry?.actorId) return;
  PERSISTENT_SAVE_QUEUE.push({
    actorId: entry.actorId,
    tokenUuid: entry.tokenUuid ?? null,
    saveType: entry.saveType ?? null,
    second: Boolean(entry.second),
  });
}

function takePersistentSave({ actorId, saveType }) {
  if (!actorId) return null;
  const index = PERSISTENT_SAVE_QUEUE.findIndex(
    (entry) => entry.actorId === actorId && (!saveType || entry.saveType === saveType)
  );
  if (index === -1) return null;
  return PERSISTENT_SAVE_QUEUE.splice(index, 1)[0];
}

function buildTokenHeaderHtml(tokenDoc, labelText) {
  const img = tokenDoc?.texture?.src ?? tokenDoc?.img ?? "";
  const name = tokenDoc?.name ?? "";
  const uuid = tokenDoc?.uuid ?? "";
  const showLabel = Boolean(labelText);
  return `
    <div class="NAS-token" data-uuid="${uuid}" style="margin-bottom: 6px;">
      <div style="display: flex; justify-content: center;">
        <img src="${img}" title="${name}" width="48" height="48" style="margin-bottom: 6px; cursor: pointer;"/>
      </div>
      ${showLabel ? `<span style="text-align: center; display: block;">${name} ${labelText}</span>` : ""}
    </div>
  `;
}

function getPersistentLabel(isSecond) {
  if (!isSecond) return "";
  return game.i18n.localize("NAS.metamagic.PersistentSaveSecond");
}

function attachTokenImageInteractions(html) {
  const tokenImgs =
    typeof html.find === "function"
      ? html.find(".NAS-token img")
      : html.querySelectorAll(".NAS-token img");
  if (!tokenImgs?.length) return;
  const attach = (img) => {
    img.addEventListener("click", async (ev) => {
      const icToken = ev.currentTarget.closest(".NAS-token");
      const tokenUuid = icToken?.dataset?.uuid;
      if (!tokenUuid) return;
      const tokenDocument = await fromUuid(tokenUuid);
      const token = canvas.tokens?.get?.(tokenDocument?.id);
      if (token) {
        token.control({ releaseOthers: true });
        canvas.animatePan({ x: token.center?.x, y: token.center?.y, duration: 1000 });
      }
    });
    img.addEventListener("mouseenter", async (ev) => {
      const icToken = ev.currentTarget.closest(".NAS-token");
      const tokenUuid = icToken?.dataset?.uuid;
      if (!tokenUuid) return;
      const tokenDocument = await fromUuid(tokenUuid);
      const token = canvas.tokens?.get?.(tokenDocument?.id);
      if (token) token._onHoverIn(ev);
    });
    img.addEventListener("mouseleave", async (ev) => {
      const icToken = ev.currentTarget.closest(".NAS-token");
      const tokenUuid = icToken?.dataset?.uuid;
      if (!tokenUuid) return;
      const tokenDocument = await fromUuid(tokenUuid);
      const token = canvas.tokens?.get?.(tokenDocument?.id);
      if (token) token._onHoverOut(ev);
    });
  };

  if (typeof tokenImgs.each === "function") {
    tokenImgs.each((_, img) => attach(img));
  } else {
    tokenImgs.forEach((img) => attach(img));
  }
}

async function resolveTokenDoc({ entry, message, actor, allowGlobal }) {
  if (entry?.tokenUuid) {
    const tokenDoc = await fromUuid(entry.tokenUuid);
    if (tokenDoc) return tokenDoc;
  }
  if (message?.speaker?.token && canvas?.scene?.id === message.speaker?.scene) {
    const tokenDoc = canvas.tokens?.get?.(message.speaker.token)?.document ?? null;
    if (tokenDoc) return tokenDoc;
  }
  if (allowGlobal && actor?.id && canvas?.tokens?.placeables?.length) {
    const token = canvas.tokens.placeables.find((placeable) => placeable?.actor?.id === actor.id);
    return token?.document ?? null;
  }
  return null;
}

function registerPersistentSaveChatMessageHook() {
  if (hooksRegistered) return;
  hooksRegistered = true;

  Hooks.on("pf1ActorRollSave", async (actor, message, savingThrowId) => {
    const entry = takePersistentSave({ actorId: actor?.id, saveType: savingThrowId });
    if (!message) return;
    const allowGlobal = game.settings.get(MODULE.ID, "saveRollTokenInteraction");
    const labelText = getPersistentLabel(entry?.second);
    if (!allowGlobal && !labelText) return;

    const tokenDoc = await resolveTokenDoc({ entry, message, actor, allowGlobal });

    const canAddToken = Boolean(allowGlobal && tokenDoc);
    const header = canAddToken
      ? buildTokenHeaderHtml(tokenDoc, labelText)
      : labelText
        ? `<p>${labelText}</p>`
        : "";
    if (!header) return;
    const content = `${header}${message.content ?? ""}`;

    const updateData = { content };
    if (canAddToken) {
      foundry.utils.setProperty(
        updateData,
        `flags.${MODULE.ID}.${SAVE_TOKEN_FLAG}`,
        true
      );
    }
    if (entry) {
      foundry.utils.setProperty(
        updateData,
        `flags.${MODULE.ID}.metamagic.persistentSave`,
        true
      );
    }
    foundry.utils.setProperty(
      updateData,
      `flags.${MODULE.ID}.metamagic.persistentSecond`,
      Boolean(entry?.second)
    );
    if (tokenDoc?.uuid) {
      foundry.utils.setProperty(
        updateData,
        `flags.${MODULE.ID}.metamagic.tokenUuid`,
        tokenDoc.uuid
      );
    }
    await message.update(updateData);
  });

  Hooks.on("renderChatMessage", (message, html) => {
    const allowGlobal = message?.flags?.[MODULE.ID]?.[SAVE_TOKEN_FLAG];
    const isPersistent = message?.flags?.[MODULE.ID]?.metamagic?.persistentSave;
    if (!allowGlobal && !isPersistent) return;
    attachTokenImageInteractions(html);
  });
}

function getSelectedTokens() {
  return Array.from(canvas?.tokens?.controlled ?? []);
}

async function resolveTargetsFromMessage(message) {
  const uuids =
    message?.flags?.[MODULE.ID]?.targets ??
    message?.system?.targets ??
    message?.data?.system?.targets ??
    [];
  if (!Array.isArray(uuids) || uuids.length === 0) return [];
  const resolved = await Promise.all(
    uuids.map(async (uuid) => {
      try {
        return await fromUuid(uuid);
      } catch {
        return null;
      }
    })
  );
  return resolved.filter(Boolean);
}

function getTargetsFromSetting(message) {
  const mode = game.settings.get(MODULE.ID, "persistentSpellTargetMode");
  if (mode === "message") {
    return resolveTargetsFromMessage(message);
  }
  if (mode === "selected") {
    return getSelectedTokens();
  }
  return Array.from(game.user?.targets ?? []);
}

function getNoTargetsMessage(mode) {
  const key = `NAS.metamagic.PersistentNoTargets.${mode ?? "current"}`;
  return game.i18n.localize(key);
}

async function applyDazedCondition(targetActor, rounds) {
  if (!targetActor || !Number.isFinite(rounds) || rounds <= 0) return;
  const duration = {
    rounds,
    seconds: rounds * 6,
    startRound: game.combat?.round ?? null,
    startTurn: game.combat?.turn ?? null,
    startTime: game.time?.worldTime ?? null,
  };
  await targetActor.setCondition("dazed", { duration });
}

async function rollMetamagicSaveForTarget(target, saveType, dc, options) {
  const token = getTokenFromTarget(target);
  const tokenDoc = token?.document ?? token;
  const actor = tokenDoc?.actor ?? token?.actor ?? null;
  if (!actor) return;

  if (options?.persistent) {
    queuePersistentSave({
      actorId: actor.id,
      tokenUuid: tokenDoc?.uuid,
      saveType,
      second: false,
    });
  }
  const first = await actor.rollSavingThrow(saveType, {
    skipDialog: true,
    token: token?.object ?? token,
    dc,
  });

  const firstRoll = first?.rolls?.[0];
  const firstRollData =
    typeof firstRoll === "string"
      ? (() => {
        try {
          return JSON.parse(firstRoll);
        } catch {
          return null;
        }
      })()
      : firstRoll ?? null;
  const firstTotal = Number(firstRollData?.total ?? 0);
  const firstSuccess = Number.isFinite(dc) ? firstTotal >= dc : false;
  if (!firstSuccess) {
    if (options?.dazing) {
      const extraRounds = getDazingExtraRoundsForTarget(options?.dazingSpellName, actor);
      const totalRounds = (options?.dazingRounds ?? 1) + extraRounds;
      await applyDazedCondition(actor, totalRounds);
    }
    return;
  }

  if (options?.persistent) {
    queuePersistentSave({
      actorId: actor.id,
      tokenUuid: tokenDoc?.uuid,
      saveType,
      second: true,
    });
    const second = await actor.rollSavingThrow(saveType, {
      skipDialog: true,
      token: token?.object ?? token,
      dc,
    });

    const secondRoll = second?.rolls?.[0];
    const secondRollData =
      typeof secondRoll === "string"
        ? (() => {
          try {
            return JSON.parse(secondRoll);
          } catch {
            return null;
          }
        })()
        : secondRoll ?? null;
    const secondTotal = Number(secondRollData?.total ?? 0);
    const secondSuccess = Number.isFinite(dc) ? secondTotal >= dc : false;
    if (options?.dazing && !secondSuccess) {
      const extraRounds = getDazingExtraRoundsForTarget(options?.dazingSpellName, actor);
      const totalRounds = (options?.dazingRounds ?? 1) + extraRounds;
      await applyDazedCondition(actor, totalRounds);
    }
  }
}

export function registerPersistentSpellSaveOverrides(html) {
  const root = Array.isArray(html) ? html[0] : html;
  if (!root || root.dataset?.[LISTENER_FLAG]) return;
  root.dataset[LISTENER_FLAG] = "true";
  registerPersistentSaveChatMessageHook();

  root.addEventListener(
    "click",
    async (event) => {
      const button = event.target?.closest?.(SAVE_BUTTON_SELECTOR);
      if (!button) return;
      const message = getMessageFromEvent(event);
      if (!message) return;
      const metamagicFlags = message.flags?.[MODULE.ID]?.metamagic ?? {};
      if (!metamagicFlags.persistent && !metamagicFlags.dazing) return;

      const saveType = button.dataset?.save;
      const dc = Number(button.dataset?.dc ?? 0);
      if (!saveType || !Number.isFinite(dc)) return;

      event.preventDefault();
      event.stopImmediatePropagation();

      const targets = await getTargetsFromSetting(message);
      if (!targets.length) {
        const mode = game.settings.get(MODULE.ID, "persistentSpellTargetMode");
        ui.notifications.warn(getNoTargetsMessage(mode));
        return;
      }

      for (const target of targets) {
        await rollMetamagicSaveForTarget(target, saveType, dc, {
          persistent: Boolean(metamagicFlags.persistent),
          dazing: Boolean(metamagicFlags.dazing),
          dazingRounds: metamagicFlags.dazingRounds,
          dazingSpellName: metamagicFlags.dazingSpellName,
        });
      }
    },
    true
  );
}
