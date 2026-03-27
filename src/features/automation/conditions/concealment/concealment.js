import { MODULE } from '../../../../common/module.js';
import { findMatchingBuffs, applyBuffToTargets } from '../../buffs/buffs.js';

export const CONCEALMENT_CONDITION_ID = "concealment";
const CONCEALMENT_FLAG_KEY = "concealmentState";
const CONCEALMENT_TEXTURE = "modules/nevelas-automation-suite/src/icons/concealment.png";
const concealmentProcessing = new Set();

function getPrimaryTokenForActor(actor) {
  if (!actor) return null;
  return actor.getActiveTokens?.()[0] ?? null;
}

export function actorHasBlindFight(actor) {
  return actor?.items?.some(i => i.type === "feat" && i.name?.toLowerCase() === "blind-fight") ?? false;
}

export function getConcealmentState(actor) {
  return actor?.getFlag?.(MODULE.ID, CONCEALMENT_FLAG_KEY) || null;
}

async function setConcealmentState(actor, state) {
  if (!actor) return;
  if (!state) {
    await actor.unsetFlag(MODULE.ID, CONCEALMENT_FLAG_KEY);
  } else {
    await actor.setFlag(MODULE.ID, CONCEALMENT_FLAG_KEY, state);
  }
}

async function promptConcealmentVariant() {
  return new Promise(resolve => {
    const buttons = {
      normal: {
        label: game.i18n.localize('NAS.conditions.main.ConcealmentNormal'),
        callback: () => resolve("normal")
      },
      total: {
        label: game.i18n.localize('NAS.conditions.main.ConcealmentTotal'),
        callback: () => resolve("total")
      },
      cancel: {
        label: game.i18n.localize('NAS.common.buttons.cancel') ?? "Cancel",
        callback: () => resolve(null)
      }
    };

    new Dialog({
      title: game.i18n.localize('NAS.conditions.main.ConcealmentDialogTitle'),
      content: `<p>${game.i18n.localize('NAS.conditions.list.concealment.description')}</p>`,
      buttons,
      default: "normal"
    }).render(true);
  });
}

async function createBuffOnActor(actor, sourceDoc, name) {
  const buffData = sourceDoc ? sourceDoc.toObject() : {
    name,
    type: "buff",
    img: CONCEALMENT_TEXTURE,
    system: { active: true }
  };
  buffData.name ??= name;
  buffData.type = "buff";
  buffData.img ??= CONCEALMENT_TEXTURE;
  buffData.system = buffData.system || {};
  buffData.system.active = true;
  const created = await actor.createEmbeddedDocuments("Item", [buffData]);
  return created?.[0] ?? null;
}

async function ensureConcealmentBuff(actor, variant) {
  const buffName = variant === "total" ? "Concealment (Total)" : "Concealment";
  const existing = actor.items.find(i => i.type === "buff" && i.name === buffName && i.system?.active);
  if (existing) {
    return { buffId: existing.id, managed: false };
  }

  let appliedBuff = null;
  let managed = true;
  const matches = await findMatchingBuffs(buffName);
  if (matches?.length > 0) {
    const match = matches[0];
    const token = getPrimaryTokenForActor(actor);
    if (token) {
      await applyBuffToTargets(match, [token], null, undefined, { silent: true });
      appliedBuff = actor.items.find(i => i.type === "buff" && i.name === buffName);
    } else {
      appliedBuff = await createBuffOnActor(actor, match.document, buffName);
    }
  } else {
    appliedBuff = await createBuffOnActor(actor, null, buffName);
  }

  if (appliedBuff) {
    await appliedBuff.setFlag(MODULE.ID, "concealmentManaged", true);
    return { buffId: appliedBuff.id, managed };
  }

  return { buffId: null, managed: false };
}

async function cleanupConcealment(actor) {
  const state = getConcealmentState(actor);
  const buffId = state?.buffId;
  const managed = state?.managed;
  if (buffId && managed) {
    const buff = actor.items.get(buffId);
    if (buff) {
      await buff.update({ "system.active": false });
    } else {
      const flagged = actor.items.find(i => i.type === "buff" && i.getFlag?.(MODULE.ID, "concealmentManaged"));
      if (flagged) {
        await flagged.update({ "system.active": false });
      }
    }
  }
  await setConcealmentState(actor, null);
}

export async function handleConcealmentToggle(actor, conditionId, value) {
  if (conditionId !== CONCEALMENT_CONDITION_ID) return;
  if (concealmentProcessing.has(actor.id)) return;

  concealmentProcessing.add(actor.id);
  try {
    const applying = value !== false && !(typeof value === "object" && value?.active === false);
    const removing = value === false || (typeof value === "object" && value?.active === false);
    const state = getConcealmentState(actor);

    if (applying) {
      if (!state) {
        const variant = await promptConcealmentVariant();
        if (!variant) {
          await actor.setCondition(CONCEALMENT_CONDITION_ID, false);
          return;
        }
        try {
          const buffInfo = await ensureConcealmentBuff(actor, variant);
          await setConcealmentState(actor, {
            variant,
            buffId: buffInfo.buffId,
            managed: buffInfo.managed
          });
        } catch (err) {
          console.error(`${MODULE.ID} | Failed to apply concealment buff`, err);
        }
      }
    } else if (removing) {
      await cleanupConcealment(actor);
    }
  } finally {
    concealmentProcessing.delete(actor.id);
  }
}


