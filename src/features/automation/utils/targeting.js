export function normalizeTargets(options) {
    if (!options.targets?.length) options.targets = canvas.tokens.controlled;
    if (!options.targets?.length && game.user.character) options.targets = [game.user.character];

    options.targets = options.targets
        .map((t) => t.actor || t)
        .filter((t) => t instanceof Actor);

    return options.targets;
}
