# 🚧 READ ME is still under construction

# BMAT (BA2 Merging Automation Tool) for CC (Creation Club) Mods

![BMAT Banner](https://img.shields.io/badge/Fallout_4-BA2_Limit_Fix-blue?style=for-the-badge&logo=fallout) ![Version](https://img.shields.io/badge/Version-1.0.0-green?style=for-the-badge) ![License](https://img.shields.io/badge/License-MIT-orange?style=for-the-badge)

**The ultimate "One-Click" solution for the Fallout 4 BA2 limit.**

[Download Latest Release](https://github.com/rjshadowface/RJs_BMAT_CC/releases) | [NexusMods Page (BMAT)](https://www.nexusmods.com/fallout4/mods/89306) | [Video Guides](https://youtube.com/playlist?list=PLnRJa1RqXU3rI2LYPPeHMsHmbQKEk9L1D&si=fQGpc-2eMrj1GKQx) | [Support Development](https://ko-fi.com/rjdevden)

---

## ⚠️ Important: Antivirus & Code Signing
**Please Read Before Downloading**

BMAT_CC is currently undergoing the verification process with **SignPath** to receive a trusted digital certificate. Until this process is complete, the executable is **unsigned**.

* **You may see:** A "Windows protected your PC" (SmartScreen) or Antivirus warning.
* **Why?** AV software flags new, unsigned `.exe` and `.dll` files by default as a precaution.
* **What to do:** You can safely create an exception for the BMAT_CC folder and the BMAT_CC files. If you are a developer, you can review the source code in this repository to verify its safety.
* **Status:** ⏳ *Pending SignPath approval.*

---

## 📖 Overview
Fallout 4 game has a hard engine limit of approximately **255 BA2 archives** (more details [here](https://github.com/rjshadowface/RJs_BMAT_CC/wiki/Fallout-4-Game-Engine-Limits)). Once you surpass this, your game will crash on start or experience missing textures.

**BMAT** automates the tedious process of unpacking and repacking these BA2 archives. Unlike other tools that require complex manual steps, BMAT is designed for mod collections of any size and combinations. It detects your mods, merges their BA2s into a one or more optimized archives, and frees up hundreds of BA2 slots in your mod collection so you can add even more mods to your game.

**BMAT_CC** was specifically designed to handle the CC mods by streamlining the process and to be the first mod manager agnostic version of BMAT.

### ✨ Key Features
* **Automated Merging:** Scans your game folder and load order and merges the CC BA2s files freeing up slots for more mods.
* **Automated Restoration:** Restores the CC mods files back to their original location with couple of clicks.
* **Smart Updates:** Detects when a mod has been updated, removed, or changed and automatically performs merge files update/re-creation.
* **Merge Traceability:** Automatically keeps track of which mods files have been merged and where.
* **Pre-packaged:** Comes pre-packaged and preconfigured with PowerShell 7.x, so you don't have to install it anymore ahead of time. The mod package on NexusMods comes with embedded tools (BSArch, etc.) so you don't need to download extra dependencies.

---

## ⚙️ Prerequisites
* **Game:** Fallout 4 (Steam or GOG).
* **Mod Manager:** Works with both Vortex and MO2.
* **OS:** Windows 10/11 (64-bit).

---

## 🚀 Installation - [Wiki](https://github.com/rjshadowface/RJs_BMAT_CC/wiki/Installation)

---

## 🛠 Configuration
BMAT uses a `config.json` file generated next to the EXE. You can edit this to customize the tool behavior.

| Setting | Description | Default |
| :--- | :--- | :--- |
| `BA2MergingStagingFolder` | [cite_start]Where original BA2s are moved to prevent double-loading[cite: 396]. | *User Promoted* |
| `BA2MergingTempFolder` | [cite_start]Fast disk location for extraction/repacking[cite: 397]. | `""` |
| `ModSizeBaseProcessingLimitMB` | [cite_start]Skips mods larger than this size (0 = process all)[cite: 435]. | `100` |
| `DetailedLogging` | [cite_start]Set to `false` to reduce log file size[cite: 440]. | `true` |


---

## 📋 How It Works 

---

## 🤝 Contributing & Support
Found a bug? Have a feature request?
* 🐛 **Report Issues:** [GitHub Issues](../../issues)
* 📺 **Guides:** [YouTube Channel](https://www.youtube.com/@RJDevDen)
* ☕ **Support:** [Buy me a Coffee](https://ko-fi.com/rjdevden)

### Credits
Special thanks to the community members who provided insights and testing:
* **VilanceD** for initial inspiration.
* **Exoclyps** for "A StoryWealth" collection and the **aSW community** for all the testing and guidance.
* **Zilav﻿** for BSArch tool as major component of BMAT.
* See the full list on the [NeuxMods page](https://www.nexusmods.com/fallout4/mods/89306).

---
*Disclaimer: ⚠️ Use at your own risk. Always backup your load order before performing a massive merge operation.*
