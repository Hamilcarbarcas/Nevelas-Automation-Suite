
import { MODULE } from '../../../common/module.js';
import { CONCEALED_CONDITION_ID } from './concealed/concealed.js';

export function registerConditions(registry) {
  registry.tracks.push("immobilize");

  const conditions = [
    {
      namespace: MODULE.ID,
      key: "anchored",
      value: {
        journal: "Compendium.nevelas-automation-suite.Conditions.JournalEntry.7OPIeT3M6ahk31kn.JournalEntryPage.2Hn63eW4jZ0XOgGI",
        flags: {},
        mechanics: {
          changes: [],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.anchored.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/anchored.svg",
        track: "immobilize"
      }
    },
    {
      namespace: MODULE.ID,
      key: "burning",
      value: {
        journal: "Compendium.nevelas-automation-suite.Conditions.JournalEntry.7OPIeT3M6ahk31kn.JournalEntryPage.Csyiw16C8wgu5C0D",
        flags: {},
        mechanics: {
          changes: [],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.burning.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/burning.svg",
        track: ""
      }
    },
    {
      namespace: MODULE.ID,
      key: "concealed",
      value: {
        journal: "Compendium.nevelas-automation-suite.Conditions.JournalEntry.7OPIeT3M6ahk31kn.JournalEntryPage.kfctyAUASIwxLFYG",
        flags: {},
        mechanics: {
          changes: [],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.concealed.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/concealed.svg",
        track: ""
      }
    },
    {
      namespace: MODULE.ID,
      key: "energyDrained",
      value: {
        journal: "Compendium.pf1.pf1e-rules.JournalEntry.NSqfXaj4MevUR2uJ.JournalEntryPage.onMPh2re6fIeNgNr",
        flags: {},
        mechanics: {
          changes: [],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.energyDrained.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/drained.svg",
        track: ""
      }
    },
    {
      namespace: MODULE.ID,
      key: "fascinated",
      value: {
        journal: "Compendium.pf1.pf1e-rules.JournalEntry.NSqfXaj4MevUR2uJ.JournalEntryPage.Hy0MHwpRRr5QxVj5",
        flags: {},
        mechanics: {
          changes: [
            { formula: '-4', target: 'skill.per', type: 'untyped' }
          ],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.fascinated.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/fascinated.svg",
        track: ""
      }
    },
    {
      namespace: MODULE.ID,
      key: "grappling",
      value: {
        journal: "Compendium.nevelas-automation-suite.Conditions.JournalEntry.7OPIeT3M6ahk31kn.JournalEntryPage.QeJmifftmLMo2MOl",
        flags: {},
        mechanics: {
          changes: [
            {"formula": "-4", "operator": "add", "target": "dexPen", "type": "untyped", "priority": 0}, 
            {"formula": "-2", "operator": "add", "target": "attack", "type": "untyped", "priority": 0}
          ],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.grappling.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/grappling.svg",
        track: "",
        statuses: new Set(["slowed"])
      }
    },
    {
      namespace: MODULE.ID,
      key: "squeezed",
      value: {
        journal: "Compendium.pf1.pf1e-rules.JournalEntry.NSqfXaj4MevUR2uJ.JournalEntryPage.jzeZ0Uf31CAMZra9",
        flags: {},
        mechanics: {
          changes: [
            { formula: "-4", operator: "add", target: "attack", type: "untyped", priority: 0 },
            { formula: "-4", operator: "add", target: "ac", type: "untyped", priority: 0 }
          ],
          flags: ["loseDexToAC"]
        },
        name: game.i18n.localize("NAS.conditions.list.squeezed.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/squeezed.svg",
        track: ""
      }
    },
    {
      namespace: MODULE.ID,
      key: "immobilized",
      value: {
        journal: "Compendium.nevelas-automation-suite.Conditions.JournalEntry.7OPIeT3M6ahk31kn.JournalEntryPage.ARVUTGA1smvapSJi",
        flags: {},
        mechanics: {
          changes: [],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.immobilized.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/immobilized.svg",
        track: ""
      }
    },
    {
      namespace: MODULE.ID,
      key: "slowed",
      value: {
        journal: "Compendium.nevelas-automation-suite.Conditions.JournalEntry.7OPIeT3M6ahk31kn.JournalEntryPage.9M72iVJj5h8nOGCf",
        flags: {},
        mechanics: {
          changes: [
            {
              "type": "untyped",
              "operator": "set",
              "formula": "@attributes.speed.land.total / 2",
              "target": "landSpeed"
            },
            {
              "type": "untyped",
              "operator": "set",
              "formula": "@attributes.speed.climb.total / 2",
              "target": "climbSpeed"
            },
            {
              "type": "untyped",
              "operator": "set",
              "formula": "@attributes.speed.swim.total / 2",
              "target": "swimSpeed",
            },
            {
              "type": "untyped",
              "operator": "set",
              "formula": "@attributes.speed.burrow.total / 2",
              "target": "burrowSpeed",
            },
            {
              "type": "untyped",
              "operator": "set",
              "formula": "@attributes.speed.fly.total / 2",
              "target": "flySpeed",
            }
          ],
          flags: []
        },
        name: game.i18n.localize("NAS.conditions.list.slowed.label"),
        showInAction: true,
        showInDefense: true,
        texture: "modules/nevelas-automation-suite/src/icons/slowed.svg",
        track: ""
      }
    }
  ];

  conditions.forEach(condition => {
    registry.register(condition.namespace, condition.key, condition.value);
  });

  const movementConditions = ["disabled", "entangled", "exhausted", "grappling"];
  for (const cond of movementConditions) {
    const condition = registry.get(cond);
    if (condition) {
      condition.updateSource({ statuses: new Set(["slowed"]) });
    }
  }

  const grappled = registry.get("grappled");
  if (grappled) {
    grappled.updateSource({
      track: "immobilize",
      statuses: new Set(["immobilized"])
    });
  }

  const entangled = registry.get("entangled");
  if (entangled) {
    entangled.updateSource({ track: "immobilize" });
  }

  const invisible = registry.get("invisible");
  if (invisible) {
    invisible.updateSource({
      statuses: new Set([...(invisible.statuses ?? []), CONCEALED_CONDITION_ID])
    });
  }
}


export function setupConditionsI18n(t) {
  t.conditions.anchored = "NAS.conditions.list.anchored.description";
  t.conditions.energyDrained = "NAS.conditions.list.energyDrained.description";
  t.conditions.fascinated = "NAS.conditions.list.fascinated.description";
  t.conditions.immobilized = "NAS.conditions.list.immobilized.description";
  t.conditions.slowed = "NAS.conditions.list.slowed.description";
  t.conditions.burning = "NAS.conditions.list.burning.description";
  t.conditions.grappling = "NAS.conditions.list.grappling.description";
  t.conditions.grappled = "NAS.conditions.list.grappled.description";
  t.conditions.squeezed = "NAS.conditions.list.squeezed.description";
  t.conditions.concealed = "NAS.conditions.list.concealed.description";
} 
