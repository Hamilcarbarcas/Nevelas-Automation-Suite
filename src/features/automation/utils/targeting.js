import { isActorDocument } from "../../../common/foundryCompat.js";

export function normalizeTargets(options) {
    const toArray = (maybe) => {
        if (maybe == null) return [];
        if (Array.isArray(maybe)) return maybe;
        if (typeof maybe[Symbol.iterator] === "function") return Array.from(maybe);
        return [maybe];
    };

    const toActor = (t) => {
        if (!t) return null;
        if (isActorDocument(t)) return t;
        if (isActorDocument(t.actor)) return t.actor;
        if (isActorDocument(t.document?.actor)) return t.document.actor;
        if (isActorDocument(t.object?.actor)) return t.object.actor;
        return null;
    };

    const hasClickContext = options?.element != null || options?.event != null || options?.message != null;

    let rawTargets = toArray(options?.targets);

    if (!rawTargets.length && hasClickContext) rawTargets = toArray(canvas.tokens?.controlled);
    if (!rawTargets.length && hasClickContext) rawTargets = toArray(game.user?.targets);
    if (!rawTargets.length && hasClickContext) rawTargets = toArray(options?.message?.targets);

    if (!rawTargets.length && !hasClickContext) rawTargets = toArray(options?.message?.targets);
    if (!rawTargets.length && !hasClickContext) rawTargets = toArray(game.user?.targets);
    if (!rawTargets.length && !hasClickContext) rawTargets = toArray(canvas.tokens?.controlled);

    const actors = rawTargets.map(toActor).filter(isActorDocument);
    return actors;
}
