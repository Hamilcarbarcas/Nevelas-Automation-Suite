const CONTEXT_TTL_MS = 2000;
const contextsByActorUuid = new Map();

function actorKey(actor) {
  return String(actor?.uuid ?? actor?.id ?? "");
}

function now() {
  return Date.now();
}

function pruneExpiredContexts(nowMs = now()) {
  for (const [key, entries] of contextsByActorUuid.entries()) {
    const fresh = entries.filter((entry) => nowMs - entry.timestamp <= CONTEXT_TTL_MS);
    if (fresh.length) contextsByActorUuid.set(key, fresh);
    else contextsByActorUuid.delete(key);
  }
}

export function recordCombatTextContext(actor, context = {}) {
  const key = actorKey(actor);
  if (!key) return;
  pruneExpiredContexts();

  const entries = contextsByActorUuid.get(key) ?? [];
  entries.push({
    isCritical: context.isCritical === true,
    critMult: Math.max(0, Number(context.critMult) || 0),
    messageUuid: context.messageUuid ?? null,
    attackIndex: Number.isInteger(Number(context.attackIndex)) ? Number(context.attackIndex) : null,
    timestamp: now()
  });
  contextsByActorUuid.set(key, entries.slice(-5));
}

export function consumeCombatTextContext(actor) {
  const key = actorKey(actor);
  if (!key) return null;

  const nowMs = now();
  pruneExpiredContexts(nowMs);

  const entries = contextsByActorUuid.get(key);
  if (!entries?.length) return null;

  const entry = entries.shift();
  if (entries.length) contextsByActorUuid.set(key, entries);
  else contextsByActorUuid.delete(key);

  return nowMs - entry.timestamp <= CONTEXT_TTL_MS ? entry : null;
}
