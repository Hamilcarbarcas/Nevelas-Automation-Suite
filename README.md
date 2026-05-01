# Nevela’s Automation Suite (PF1E)

Nevela’s Automation Suite is a Foundry VTT module for the **Pathfinder First Edition (PF1E)** system that bundles automation features to streamline gameplay and reduce repetitive bookkeeping.

---

## Requirements

- Foundry VTT v12/v13
- Pathfinder 1st Edition v11+
- SocketLib module
- LibWrapper

---

## Buff Automation

Automatically applies buffs when spells, consumables, or class features are used.

### Highlights
- Requires spells/consumables/class features to have valid targets
- Supports multiple buff compendia including PF1 core and PF-Content
- Smart buff selection system that handles:
  - Exact matches
  - Variants (such as Blessing of Fervor and Evil Eye)
  - Versions (such as Heroism, Greater and Invisibility, Greater)
- Target filtering modes:
  - **By disposition**: Automatically applies buffs to targets with the same disposition as the caster
  - **Manual selection**: Prompts the caster to choose which targets should receive the buff
  - **All targets**: Applies buffs to all targets of the spell/consumable
- Duration handling:
  - Supports caster level-based durations
  - Handles complex duration formulas
  - Proper turn-based duration support
- Buff management:
  - Updates existing buffs instead of stacking
  - Maintains buff duration information
  - Supports custom buff compendia

> **Note:** You need to add a boolean flag called **`buff`** in the **Advanced** tab of class features to set a class feature to apply buffs.

---

## Condition Automation

Adds additional PF1E conditions and automates behavior for many conditions to keep combat moving quickly and consistently.

### Added Conditions
- **Burning**
- **Concealed**
- **Energy Drained**
- **Fascinated**
- **Grappling**
- **Squeezed**

### Sub-Conditions
- **Anchored**: Applies a movement restriction on tokens (immovable)
- **Immobilized**: Denies Dex bonus to AC and prevents movement
- **Slowed**: Reduces movement speeds by half

These sub-conditions work with other conditions like **entangled**, **grappled**, and more.

### Automated Behaviors (Toggleable Where Noted)
- **Anchored**: Applies a movement restriction on tokens, making them immovable; automatically applied to conditions like *cowering, dazed, dying, helpless, paralyzed, petrified, pinned*. (Toggleable in settings.)
- **Blind**: Moving a token more than half its speed prompts an Acrobatics check—failure results in no movement and the **prone** condition. (Toggleable.)
- **Concealed**: When Concealment is applied, the user will be prompted to select between the normal or total concealment. Currently it supports the Blind-Fight feat.
- **Confused**: Creates a chat message every round showing how affected tokens act. During combat, posts on the affected token’s turn. Includes proper damage application compatibility with damage automation. (Toggleable.)
- **Dead**: Applied automatically when HP is ≤ total Constitution score; can be toggled off and disables Monk’s Combat Detail’s similar setting for compatibility.
- **Disabled**: Applied automatically when HP is exactly 0; prompts user decisions for actions and condition changes based on abilities like Diehard. (Toggleable.)
- **Dying**: Applied automatically at negative hp (but not dead). On the creature's turn, it rolls a constitution check automatically with a DC of 10 + a penalty equal to the negative HP.
- **Energy Drained**: Automatically applied to tokens with above 0 negative levels. (Toggleable.)
- **Entangled**: Applies **slowed** automatically and prompts concentration checks for spellcasting. (Toggleable.)
- **Fascinated**: Applies a **-4 penalty** on Perception checks.
- **Flat-Footed**: Applied automatically based on initiative rolls, with considerations for abilities like Uncanny Dodge. (Toggleable.)
- **Grappled**: The token is immobile. Prompts concentration checks for spellcasting with somatic components. (Toggleable.)
- **Grappling**: A grapple checkbox will be attack to your attack dialogs. When used with a target, the module will process your roll and the target's CMD and apply grappling on the attacker and grapple and the defender upon success.
- **Immobilized**: Sub-condition that renders the token immovable and denies Dex bonus to AC. (Toggleable.)
- **Invisible**: Applies Total Concealment automatically as a sub-condition.
- **Prone**: Applied automatically when certain conditions are met like falling unconscious and blind movement check failure.
- **Slowed**: Sub-condition applied automatically to tokens with conditions like *disabled, entangled, exhausted, and grappling*.
- **Staggered**: Applies automatically when the token has taken nonlethal damage equal to the token's current HP.
- **Squeezed / Squeezing**: Applied automatically when the token is entering a narrow area (wall based). When attempting to enter an area smaller by more than half of the token's width, the module automatically rolls for Escape Artist with a DC of 30 and upon success the token is moved into the area, upon failure the token moves to the square before the narrow area's entry. By default, each token has a base body width and a base head width. These percentage based values for width and the DC can be modified in the module settings and the token settings.
- **Unconscious**: Applied automatically when tokens fall below 0 HP, unless affected by abilities like Diehard. Also applied automatically when the token has taken nonlethal damage equal to more than the token's current HP.

### Surprise Rounds
Surprise Rounds are available only if the **Flat-Footed** setting is enabled.

- Start a surprise round by clicking the **“Surprise Round”** button above **“Begin Combat”**.
- Applies **Flat-Footed** to all tokens until their turns in the second round.
- To exclude tokens from being surprised, select them or toggle the new <span><img src="https://github.com/user-attachments/assets/0f2cb882-8a1e-46c5-8542-5a6fde928bfd" width="23"></span> icon on the before clicking the "Surprise Round" button.

---

## Damage Automation

Enhances and streamlines PF1E damage calculation and application by automatically accounting for defenses and damage rules.

### Damage Calculation
- Calculates damage based on the target’s **Damage Reduction (DR)**, **resistances**, **immunities**, and **vulnerabilities**, including logic differences such as spell vs weapon handling.
- You can still apply damage differently using normal PF1E behavior (e.g., **shift-click**).

### Personalized Damage Types

#### Custom Damage Types
When you add a new damage type, you must include:
- **Name**
- **Image/icon** (or icon class)
- **Category**

By default, the system has 3 categories: **Physical**, **Energy**, and **Miscellaneous**, but you can define additional categories.

> *Icon classes* are what the system uses currently (for example, Font Awesome icon class names).

#### Ability Score Damage Types
Each custom damage type can also:
- Target one specific ability score
- Be defined as **ability damage**, **ability drain**, or **ability penalty**

The module supports ability damage immunity. Use the following formats in the damage immunities custom section:

| Immunity Type | Description |
|---|---|
| **Ability Damage** | Immunity to any form of ability damage. |
| **Ability Drain** | Immunity to any form of ability drain. |
| **Ability Penalty** | Immunity to any form of ability penalty. |
| **All Ability Damage** | Complete immunity to any effect that causes ability damage, drain, or penalty. |
| **`Specific Ability` Damage** | Immunity to damage to a specific ability score. |
| **`Specific Ability` Drain** | Immunity to drain of a specific ability score. |
| **`Specific Ability` Penalty** | Immunity to penalties to a specific ability score. |
| **All `Specific Ability` Damage** | Complete immunity to any effect that causes damage, drain, or penalty to a specific ability score. |
| **Mental Ability Damage** | Immunity to damage to any mental ability score (INT, WIS, CHA). |
| **Mental Ability Drain** | Immunity to drain of any mental ability score (INT, WIS, CHA). |
| **Mental Ability Penalty** | Immunity to penalties to any mental ability score (INT, WIS, CHA). |
| **All Mental Abilities** | Complete immunity to any effect that causes damage, drain, or penalty to any mental ability score. |
| **Physical Ability Damage** | Immunity to damage to any physical ability score (STR, DEX, CON). |
| **Physical Ability Drain** | Immunity to drain of any physical ability score (STR, DEX, CON). |
| **Physical Ability Penalty** | Immunity to penalties to any physical ability score (STR, DEX, CON). |
| **All Physical Abilities** | Complete immunity to any effect that causes damage, drain, or penalty to any physical ability score. |

*`Specific Ability` refers to the ability score's full name or abbreviation (e.g., STR, DEX, CON, INT, WIS, CHA) and can be uppercase, lowercase, or capitalized.*

> If you check the damage type's checkbox, both damage and ability damage from this damage type will be nullified.

#### Custom Priority (DR Table)
- Comes with Pathfinder’s default DR priority table
- You can add, remove, and edit rows
- Includes a **Reset to Defaults** button
- Each enhancement row bypasses the damage types in the enhancement row before it
- Supports system-provided materials/alignments and custom values

---

### Damage Reduction (DR) and Elemental Resistance (ER)
When adding a custom DR or ER type to a token, separate each type by a semicolon (`;`). The module can also handle operators (`and/or`) in the custom section. The amount can be before or after the type.

**Examples**
- `5/Cold Iron and Silver; 10/Glass`
- `Magma or Frost/10; Poison/5`

**DR Magic**
- A weapon with an enhancement bonus greater than 0 counts as magic for bypassing DR/magic.
- If you have a magical weapon lacking enhancement bonuses (such as some natural attacks), check the **magic** checkbox in the attack’s detail tab.

---

### Hardness
Hardness reduces damage from all sources (except ability score damage) by the value entered in the Damage Reduction section of the actor’s sheet.

- **Adamantine**: If hardness is ≤ 20, adamantine bypasses hardness completely.

#### Bypassing Defenses

Items now include:
- **Global Damage Settings**
- **Action Damage Settings**

These let you configure how your damage interacts with:
- **Hardness** (bypass completely or ignore a set amount)
- **Immunities**
- **Resistances**
- **Damage Reduction**

---

### Notes
- **Silver and Alchemical Silver are treated as the same material.** If you add Silver as a custom material, it will not bypass the Silver DR from the system. Add a custom DR with the same name instead.

---

### Damage Roll Cards
When creating a macro or chat damage roll (using `/d` or `/damage`), specify damage types using bracketed flavor after the damage. Include multiple types separated by commas, and you can split types across different dice groups.

**Examples**
- `/d 3d6[Fire, Slashing]`
- `/d 1d8[Acid] + 2d4[Slashing, Piercing] + 5[Negative]`

---

### Alternative Roll Command and Macros

In addition to `/d` and `/damage`, you can use `/as` to roll damage using the Automate Damage dialog UI.

#### Macro Creation
Use `/as macro` to create a new macro. The module opens a dialog to configure formula, damage types, and macro name.

#### Macro API: `ASDamage.roll()`
You can call `ASDamage.roll()` in a macro using either a **string** or **object** syntax:

```js
// Simple usage with a string
ASDamage.roll("type: fire, for: 1d8+5");

// Object format with multiple damage types
ASDamage.roll({
  formula: "1d8+5",
  damageTypes: ["fire", "acid"]
});

// Mixed damage types with string format
ASDamage.roll("1d8 fire + 2d6 acid + 5 bludgeoning");

// Recommended: Component-based approach
ASDamage.roll({
  components: [
    { formula: "1d8", damageTypes: ["fire"] },
    { formula: "2d6+5", damageTypes: ["slashing"] },
    { formula: "10", damageTypes: ["acid"] }
  ]
});
```

This returns a `ChatMessage` and integrates with defenses (DR, hardness, resistances, vulnerabilities, etc.).

---

### Ammunition
Because ammunition does not have material/enhancement fully supported by the PF1E system yet:

- **Material**: Add a dictionary flag named **`Material`** with the material you want the ammo to be. You can include `magic` in this dictionary flag to treat the ammo as magic to overcome DR.
- **Enhancement Bonus**: If you do not use Roll Bonuses, add a boolean flag named **`Magic`** in the ammo to treat the ammo as magic to overcome magic DR.

---

### Reactive Effects (On Hit / On Struck)

- On the item sheet (details tab) \ action sheet (action tab) you will find a "On Hit" section with a checkbox for items that have damage actions. Checking it opens configuration that allows you to configure special effects when the attack hits the target (when damage is applied), such as lifesteal or applying a buff or condition to the attacker or the target.
- Similarly, on items such as armor or buffs, you will find a "On Struck" section with a checkbox. Checking it opens configuration that allows you to configure similar effects that happen when that token \ actor is struck.

> This allows you to configure things like the fire shield spell.

---

### Wounds & Vigor

- The module has a setting to toggle on wounds and vigor support. This is dependent on the system setting as well. When both are enabled, the module handles damage as the Wounds & Vigor rules. The module also includes a house rule for construct and undead creatures.

---

## Metamagic Automation

Adds a metamagic picker to spell dialogs and automatically applies supported metamagic adjustments at cast time—reducing manual edits to spell data, DCs, ranges, and damage formulas.

### Highlights
- Metamagic selection UI on spell dialogs (Attack Dialog → **Metamagic**)
- Pulls available metamagics from the caster’s:
  - **Metamagic feats** (feat tags containing “metamagic”)
  - **Metamagic rods** in inventory (supports common rod naming patterns and checks remaining daily uses)
- Smart filtering: only shows metamagics that are **valid for the current spell** (components, range, duration, save, damage, etc.)
- Spontaneous caster support:
  - Automatically consumes the **correct higher-level spell slot**
  - Adjusts casting time according to the **Metamagic Cast Time** setting
- Chat card integration for save-based metamagics:
  - **Persistent Spell** save rerolls are handled from the save button
  - **Dazing Spell** applies **Dazed** on a failed save (with proper turn-based duration)

---

### Supported Metamagics

- **Dazing Spell**
  - Requires a damaging spell.
  - Ensures the spell has a save button (defaults to **Will** only if the spell has no save).
  - On a failed save from the chat card, applies **Dazed** for a number of rounds based on spell level.
- **Empower Spell**
  - Applies 1.5× damage (rounded down) by rewriting damage formulas.
  - Preserves readable roll tooltips for empowered rolls.
- **Enlarge Spell**
  - Increases the range of spells with the close, medium, or long categories by 100% (excludes other range values).
- **Extend Spell**
  - Doubles duration (excludes concentration, instantaneous, and permanent durations).
- **Heighten Spell**
  - Prompts for a target spell level and updates effective spell level (and DC) accordingly.
- **Intensified Spell**
  - Increases common caster-level-based damage caps by **+5** (up to caster level), where applicable.
- **Maximize Spell**
  - Converts damage dice into maximum results by rewriting dice expressions.
- **Persistent Spell**
  - Adds a reminder footnote and upgrades the chat save workflow:
    - On a successful save, automatically rolls the **second** save.
- **Reach Spell**
  - Steps range upward (**Touch → Close → Medium → Long**).
  - Prompts you to choose how many steps to apply.
  - Converts melee spell attacks (**msak**) to ranged spell attacks (**rsak**) when applicable.
- **Selective Spell**
  - For **instantaneous area** spells, prompts you to exclude targets.
  - Max exclusions are based on your spellcasting ability modifier for the spellbook.
- **Silent Spell**
  - Removes verbal components.
- **Still Spell**
  - Removes somatic components.
  - **Blocked for bard spells** (warns and skips).
- **Quicken Spell**
  - Sets casting time to **swift**.

---

### Support for over 20 Class Features \ Feats \ Traits

- **Full List of supported items**:
  Arcane Apotheosis
  Arcane Bloodline
  Arcane Prodigy (Drow)
  Curator of Mystic Secrets
  Eldritch Researcher
  Extended Scrying
  Grand Maestro
  Healer's Blessing
  Intense Celebration
  Magical Lineage
  Maleficium
  Mask Focus
  Metamagic Adept
  Metamagic Mastery
  Metamixing
  Mimic Metamagic
  Nanite Bloodline
  One Body
  Two Minds
  Peerless Speed
  Prolong Magic (Tiefling)
  Retribution
  Seeker of the Eternal Emperor
  Spell Perfection
  Spontaneous Metafocus
  Succor Final Revelation
  Timeless Soul
  Transmuter of Korada
  Wayang Spellhunter (Minata)

---

### Usage
- Open a spell’s use dialog.
- Enable **Metamagic**.
- Select one or more metamagics from the dropdown.
- Some metamagics prompt for extra info:
  - **Reach Spell**: choose range steps
  - **Heighten Spell**: choose target spell level
  - **Selective Spell**: choose which targets to exclude

---

### Settings
- **Metamagic Cast Time**
  - Controls how metamagic affects **spontaneous** casting time.
- **Persistent Spell Targets**
  - Controls how targets are determined when using the **chat save button** for Persistent/Dazing workflows:
    - current user targets / selected tokens / targets stored on the chat card

---

## Quality of Life

### On Card Save

- When an action that has a save is used, the chat card typically is created with a save button. The module adds an optional setting that allows you to add an interactable and clickable icon of the token that is making the save. This helps when using a the save button while selection multiple tokens that have the same name and icon such as NPCs.

### Enhanced Dice Tooltips

- This setting changes how the system shows unlabeled modifiers in the roll tooltips. When certain variables are available (such as caster level), it will show that bonus and label them appropriately. If a bonus is added but the module could not label it, it will be labeled as "bonus".

### Combat Text

- This setting causes the floating combat text to change how it moves and looks.

### Enforce Spell Ability Minimum

- This setting enforces the PF1E rule where you have to have a casting ability score of 10 + spell level to be able to cast spells of certain level. While the system already handles it partially, this setting also completely hides those spell levels from the actor sheet when the casting ability score is too low.

---

## Installation
**Manifest URL:** https://github.com/Nevela-0/Nevelas-Automation-Suite/releases/latest/download/module.json

---

## Usage
Once installed and enabled, configure which automations you want active in the module’s settings.

---

## License
This suite includes components licensed under different terms:

- Damage Automation components are licensed under the **MIT License** (see included LICENSE file).
- Condition/Buff components include content covered under the **Open Game License (OGL)** (see `ogl.md`).

---

## Credits
- **Nevela**: Lead developer and creator.
- **Contributors**: PF1E system and module developers for their support.
- **Claudekennilol**: Huge general support with compatibility and module settings.
- **[McGreger](https://github.com/McGreger)**: Help with German localization.
- **Condition Icons**: Icons selected from [game-icons.net](https://game-icons.net/about.html).

---

## OGL
This module includes content covered under the Open Game License (OGL). For more details, please refer to the [OGL license](./ogl.md).
