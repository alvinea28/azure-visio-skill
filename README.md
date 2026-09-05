# Azure Visio Skill

Create native, editable Azure architecture diagrams with **Microsoft Scout and desktop Visio**. Version 1.5 uses icon-first layouts, short captions, meaningful boundaries, reference-image conversion, and safe create/read/update/delete operations.

**Copilot Cowork supports planning and a Scout handoff, not direct control of desktop Visio.** This is an independent skill, not an official Microsoft or GitHub product.

## Install in Scout

1. Install and activate a licensed **desktop Visio** application on Windows. Visio for the web alone is insufficient.
2. Download [the skill ZIP](dist/azure-visio-1.5.0.zip) and extract it to a local folder outside OneDrive. Keep `SKILL.md` and all companion files together.
3. In Scout, ask:

   > Install azure-visio from [extracted folder path]. Register SKILL.md with your skill-management tools, copy every companion file into the skill's resource directory, and complete the official icon-library and reference-atlas setup under their published terms. Preserve existing settings and respect company policy.

4. Approve any required local file/shell access, then start a **new Scout chat**.
5. Try:

   > /azure-visio Create an icon-first Azure architecture with App Service, Microsoft Foundry and Azure AI Search. Use short captions, keep implementation detail in Shape Data, and save a new Visio drawing and PDF.

Visio opens automatically. Deletion, destructive replacement, and external sharing still require confirmation.

## Install in Copilot Cowork

1. Open Cowork with an account that has access.
2. Where custom-skill upload is available, open **Customize** from the left navigation or **+** menu, then **Skills > Add dropdown > Upload skill**.
3. Upload [the skill ZIP](dist/azure-visio-1.5.0.zip), not GitHub's **Code > Download ZIP** archive.
4. Start a new task: **"Use azure-visio to plan this architecture and prepare a Scout handoff."**
5. Give the handoff and original reference to Scout for native Visio creation.

If **Customize** or **Upload skill** is missing, availability depends on your Cowork interface and organization. Ask your administrator; do not bypass policy. Cowork installation does not provide local PowerShell/Visio access.

## Sample architecture images

- [Icon-first deployment architecture](examples/images/deployment-architecture.png)
- [Generic hub-and-spoke topology, legacy reference demo](examples/images/hub-spoke-reference.png)

![Icon-first deployment architecture](examples/images/deployment-architecture.png)

These are illustrative outputs, not proof of deployment, security, or production readiness.

## More information

See [README.txt](README.txt) for manual setup and [CRUD-Guide.txt](CRUD-Guide.txt) for commands and limitations.

Product names, trademarks, and imagery remain their respective owners' property. Microsoft icons and reference downloads retain their [published usage terms](https://learn.microsoft.com/en-us/azure/architecture/icons/). Downloaded libraries, private configuration, customer source files, and native working drawings are not bundled in this repository.
