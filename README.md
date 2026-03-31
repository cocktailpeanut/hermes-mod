# Hermes Mod

Hermes Mod lets you manage Hermes CLI skins in a web ui.

![nous.png](nous.png)

It does the manual work for you:
- lists built-in skins and custom skins
- opens a skin into a visual editor
- edits real Hermes skin fields, including generated banner logo text and higher-detail hero text art
- saves directly to `~/.hermes/skins/`
- activates a skin by updating `~/.hermes/config.yaml`
- shows generated YAML and a live preview
- supports multiple hero image render styles, including braille, ASCII ramp, blocks, and dots

## What it edits

The app works with the Hermes skin schema, including:
- colors
- spinner faces, verbs, and wings
- branding strings
- tool prefix
- tool emoji overrides
- `banner_logo`
- `banner_hero`

## Install

### 1. 1-Click Install

Find it on https://pinokio.co and 1-click install.

### 2. Run with npx

```bash
npx -y hermes-mod
```

The `-y` flag skips the install prompt and starts the published package immediately.

### 3. Manual Install

Go into `app` and run:

```
npm install
```

Then run:

```
npm start
```

## How to use


https://github.com/user-attachments/assets/52d911c3-6017-458c-92f6-c59f057c0528



1. Install the app in Pinokio.
2. Start the app.
3. Open Skin Studio.
4. Choose a built-in or custom skin.
5. Generate a logo from text and upload a PNG, JPG, GIF, or WEBP image to create hero art. Optionally change the hero look style or width.
6. Edit fields and click Save.
7. Click Activate to set it as the current Hermes skin.

## File locations used

By default the app uses:
- skins folder: `~/.hermes/skins/`
- config file: `~/.hermes/config.yaml`

It can also respect `HERMES_HOME` if that environment variable is set.

## Notes

- Built-in skins are shown as templates and can be duplicated into custom skins.
- Custom skins are saved as YAML files in your Hermes skin directory.
- Activation updates the `display.skin` setting in the Hermes config file.
