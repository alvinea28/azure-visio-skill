AZURE VISIO SKILL 1.6

One reusable architecture skill, with two direct output modes:
- Microsoft Scout on Windows: native editable desktop Visio creation and editing.
- Microsoft Copilot Cowork with permitted Python execution: directly generates
  editable .vsdx and .pdf files inside the task using portable_visio.py.
  No Scout bridge, handoff, desktop Visio, Windows, or COM is needed for that
  creation path. A skill cannot grant a host missing execution/download rights.

WHAT CHANGED
Each newly authored drawing now has exactly three pages:
1. The requested main architecture.
2. A separately proposed hardened architecture when applicable and not already
   present or required; otherwise a substantive hardening applicability review.
3. A text-rich connected flowchart and descriptive write-up of page 1.
Both the editable Visio file and matching three-page PDF are real deliverables.
The final page traces every main service/function and relationship, rather than
describing the proposed alternative. Hardening reviews are not certifications.
The output contract is validated and retained in native Shape Data.
Connections use horizontal/vertical segments and right-angle bends, never
diagonal shortcuts. Endpoints attach to the intended services and return/control
paths use clear gutters instead of crossing unrelated icons or captions.

Cowork now has a cross-platform direct-file renderer instead of a Scout handoff.
Intermediate model JSON, scripts, notes and previews are not final deliverables.
Unavailable Python, dependencies, icons, source rights, or downloadable output
must be reported as blockers, never replaced with a success-shaped handoff.

Enterprise architecture now defaults to freestanding service icons, short
captions and meaningful boundaries, not paragraph-filled process cards.
Full requirements and implementation detail remain in native Shape Data.
The style is grounded in eleven supplied Architecture Center examples:
Maximo, Virtual WAN, M365 DevOps, ASE HA, write-through cache, SAP, AKS,
mainframe modernization, Foundry, Foundry landing zone and conversation insights.
This is a curated reference atlas and renderer/skill update, not model training.
See Architecture-Guide.txt for archetypes and local presentation guardrails.

Quiet, hands-free operation is now the default. See Quiet-Workflow.txt.
Pre-build prose is limited to a brief sentence or essential short decision.
Research is grouped into one bounded research task where supported, and full
reports/logs stay out of the main chat. Scout's own activity rows may still
appear; a skill cannot forcibly hide or collapse host UI.
The controller reuses or automatically opens installed desktop Visio.
Use -NoLaunchVisio only for an explicitly requested attach-only operation.
Destructive actions, external sharing, permissions, and real blockers retain
their normal safeguards.

Reference-first conversion now persists the actual attachment and an independent
source contract. Missing/changed sources, dropped components/detail, added
adapters, altered grouping, or reordered layers fail the fidelity gate.
Reference-plus-proposal requests preserve the reference as the first view;
proposed changes are separate. See Reference-Workflow.txt.
Detail cards on the final flowchart add native bold headings and left/top body
alignment; configurable connector attachment positions separate flows.

Start from the user's requirements or draft, not a fixed hub-and-spoke template.
Read Architecture-Guide.txt for source extraction, reference selection, coverage,
Azure adaptation, and safe CRUD. Custom workloads use custom three-page packs.
The generic hub-and-spoke reference is an opt-in legacy demo, not the default.

Attach a readable image, Excalidraw file, or flowchart, or paste a text diagram.
The assistant visually interprets images/text; Import-Draft.ps1 extracts
structured Excalidraw or normalized graph data. There is no automatic image OCR
or universal Mermaid/draw.io parser in this package. Cropped/ambiguous input
must be flagged, not filled in from unrelated earlier drafts.

REQUIREMENTS
For Scout's desktop path, use an interactive Windows session, PowerShell 5.1
(powershell.exe), licensed desktop Visio with its native container stencil,
and permitted local shell/file access in Scout. Visio for the web is not enough.
For Cowork's direct creation path, use its permitted Python execution and file
download tools, the bundled requirements-portable.txt libraries, and an official
icon catalog. The Python renderer creates new documents; it is not a desktop
automation bridge or an arbitrary existing-Visio editing/conversion engine.
Each product's licensing, rollout, admin controls, and enterprise policies apply.
Do not deploy this controller as an unattended service.

INSTALL IN SCOUT
This archive contains SKILL.md at its root and companion scripts beside it.
On builds supporting the documented custom-skill folder, extract everything to
%USERPROFILE%\.copilot\skills\azure-visio and start a new conversation.
If that folder already exists, do not overwrite it without reviewing your edits.

Some Scout builds manage skills in a different directory. In those builds,
extract to a new local folder and ask Scout:
"Register azure-visio using this folder's SKILL.md instructions and put its
companion files next to the registered skill. Preserve existing skills."
Use Scout's skill management tools; do not guess its internal settings format.
The skill's reported resource directory is authoritative. Keep all companion
files together there. The controller uses relative asset lookup, not author paths.

INSTALL IN MICROSOFT COPILOT COWORK
Open Customize from the sidebar or + menu, where available, then
Skills > Add dropdown > Upload skill and select the v1.6 skill ZIP.
Its root must contain SKILL.md, not another enclosing azure-visio directory.
The upload is stored in OneDrive and subject to your organization's policies.
Only upload the clean distribution ZIP, not your customer drawings or your
environment.json. Import availability is not proof of local execution access.
Start a new task after replacing an older skill:
"Use azure-visio to create the requested architecture. Deliver the actual
editable Visio file and PDF directly in this task, each with the three v1.6
pages. Run the portable renderer here; do not create a Scout handoff."
If skill upload or Python execution is absent, ask the administrator about the
supported rollout; do not bypass policy or claim files were created.

DIRECT PYTHON CREATION (SCOUT OR COWORK)
Resolve the extracted skill directory first. These commands use relative script
names from that directory; replace the example input/output paths with actual
permitted locations in the current execution environment:
python -B portable_visio.py capabilities
python -B portable_visio.py render --model model.json --icon-directory IconLibrary --output-directory new-output --name architecture --bundle

Python 3.10+, ReportLab, resvg-py and Pillow are required. If missing and permitted:
python -m pip install -r requirements-portable.txt

Use an existing approved official catalog, or the install-icons command after
reviewing its --help and the Microsoft icon terms. Network access is not needed
when the catalog and its artwork are already local. The render destination must
be new. --bundle adds a ZIP containing the actual standalone VSDX and PDF for
hosts that expose ZIP downloads but cannot preview VSDX.

The PDF is generated from the same model, not exported through desktop Visio.
Service artwork is preserved as PNG within editable native groups; captions,
shapes, metadata and glued connectors remain editable. Boundaries preserve
semantic ParentId hierarchy but are not native Visio containers. Native model
export reports this distinction rather than fabricating container membership.
Visible text is limited to Windows-1252; other scripts/fonts fail explicitly.
Impossible connector routes fail rather than falling back to diagonal shortcuts.
Arbitrary later geometry edits may require rerouting to retain obstacle clearance.
Reference conversion verifies the actual source before output creation; static
package inspection alone does not revalidate source fidelity. Actual execution
inside each Cowork tenant depends on its permitted capabilities.

FIRST RUN IN SCOUT
Invoke directly; installed desktop Visio opens automatically if needed:
/azure-visio Convert the attached draft into native editable Visio. Preserve
its components and intent, use official icons, flag unclear connections, and
do not add unrelated infrastructure. Keep proposed corrections separate.

For the enterprise presentation:
/azure-visio Create an icon-first enterprise architecture. Use short captions
and real deployment/ownership boundaries. Keep implementation detail in Shape
Data or companion notes, not big boxes. Follow the closest reference atlas style.

REFERENCE ATLAS
architecture-references.json includes exact article/SVG/VSDX URLs and visual
lessons. Install-ReferenceAtlas.ps1 -OutputDirectory <new local folder> downloads
the original official assets and records source URLs and SHA256 hashes.
Study native .vsdx files using desktop Visio read-only, never unzip Office XML.
The originals are kept outside the small skill ZIP and must retain attribution.
Do not copy a reference's services/regions/topology into unrelated work.
For custom glyph libraries, prefer a non-synced local directory; the renderer
does not traverse reparse points to reach icons. Keep originals unchanged.

For a manual prerequisite check, run from the extracted skill folder:
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ".\AzureVisio.ps1" -Action Check

Check verifies the actual installed universal master names. It creates no
drawing, saves no existing document, and closes only stencils it opened.
Use Check -NoLaunchVisio for an explicitly requested attach-only check.
If blocked by managed policy, stop. Do not weaken execution policy or
automatically remove downloaded-file blocking. Follow your organization's
approved code-review, signing, and execution process.

GENERIC DEMO ONLY
The following explicitly creates the old demo. Do NOT use it for custom workloads.
In Windows PowerShell, from the skill folder:
$output = Join-Path $env:LOCALAPPDATA ('AzureVisio\Example-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $output | Out-Null
.\AzureVisio.ps1 -Action New -LegacyModel -ModelPath (Join-Path $PWD 'hub-spoke-reference.json') -DocumentPath (Join-Path $output 'Hub-Spoke.vsdx') -OutputDirectory (Join-Path $output 'Exports')

Use a user-approved output destination for real projects. New never overwrites
an existing drawing. For a reusable template:
.\AzureVisio.ps1 -Action Template -DocumentPath (Join-Path $output 'Hub-Spoke.vsdx') -TemplatePath (Join-Path $output 'Landing-Zone.vstx')

ADDING A THIRD SPOKE
Use New-ReferenceModel.ps1 -SpokeCount 3 -OutputPath <new absolute JSON path>,
then AzureVisio.ps1 -Action Merge with that model and the saved reference-derived
drawing's absolute DocumentPath. Confirm reserved layout space first.
Merge is additive, not a general update/reconciliation engine.

ONEDRIVE AND SHAREPOINT
Prefer a local, non-synced folder for the first generic example. Visio may expose
synced files as cloud URLs. If needed, copy environment.example.json to your
own environment.json and add only verified local/cloud folder mappings:
{"syncRoots":[{"local":"C:\\Users\\YOUR-USER\\OneDrive - YOUR-ORG","cloud":"https://YOUR-TENANT-my.sharepoint.com/personal/YOUR-SITE/Documents"}]}

The above values are placeholders, not a guessed tenant mapping. Compare the
actual open document FullName with its known local path before configuring it.
Supply -EnvironmentPath with an absolute config path, or keep environment.json
beside the controller. It is local-machine configuration, never a shareable asset.
The most specific matching root wins. Unmatched same-name cloud documents cause
an explicit error rather than filename-based targeting.

OFFICIAL ICON LIBRARY
Install-IconLibrary.ps1 downloads the complete six published collections linked
from Microsoft Learn: Azure, Microsoft 365, Fabric, Entra, Power Platform, and
Dynamics 365. It preserves original archives and included terms, builds a
catalog with source URLs/hashes, and flags unsafe SVG content rather than
importing it. The product artwork is not cropped, stretched, rotated, or recolored.

Read the terms on the pages in icon-sources.json. For permitted architecture,
training, or documentation use, run from the skill folder:
powershell.exe -NoProfile -ExecutionPolicy RemoteSigned -File ".\Install-IconLibrary.ps1" -OutputDirectory "$PWD\IconLibrary" -AcceptIconTerms

Use a NEW directory when refreshing. Failed downloads remain visible for
diagnosis; do not treat a partial folder without catalog.json as installed.
The catalog is an array of id/path/name/collection/sourcePage/downloadUrl/sha256
records. Model iconRef selects an exact catalog ID. The controller uses
IconLibrary beside itself, or an explicit -IconDirectory.

These are all files in the selected published collections, NOT a claim that
every Microsoft product has an official downloadable icon. Counts include
size/style variants. The M365 collection is from archived official guidance.
Do not claim older icons are current branding. Do not substitute Copilot Studio
for an unspecified Copilot or use a generic Fabric logo as a distinct Fabric IQ
logo. Use a generic labeled card when no appropriate official icon is found.

The full library is downloaded locally, not embedded in the small Cowork skill
ZIP. The ZIP includes the installer/source list so other users can obtain the
same official collections subject to their policies and terms.
Some desktop Visio versions do not faithfully import all SVG gradient/opacity
effects. Inspect the native result and select a faithful official variant or
explicit generic card; downloading the library is not a compatibility claim
for every SVG. The original icon files must remain unchanged.

STENCIL FALLBACK
The controller discovers Visio's current built-in content location and installed
language directories; it no longer assumes an English 1033 folder.
You may provide -StencilDirectory for a verified directory containing the
required original Azure .vssx files. A metric (_M) counterpart is also supported.
Universal master names are checked; localized/older editions missing a master
fail explicitly. This is not a guarantee that every Visio edition is compatible.
Legacy builtin icon keys still use the recipient's installed stencils.
Use a model-specific Check to verify only assets required by that model.
Generate native drawing templates locally.

COWORK DIRECT FILE CREATION
Run portable_visio.py using the code-execution tool actually exposed by the host.
Resolve scripts beside SKILL.md, use the environment's actual output path, and
return that host's downloadable .vsdx and .pdf artifact links. Never invent a
tool name or reference a user's PC path from a remote execution environment.
The portable renderer writes an editable Visio package and matching PDF from
one model without calling desktop Visio. It does not replace existing files.

Install requirements-portable.txt through the permitted package mechanism when
the declared libraries are missing. Use portable_visio.py --help for its exact
validation, icon-installation, and rendering commands. The portable icon path
uses catalog iconRef values, not installed Windows stencil master names.
Keep intermediate JSON/scripts internal unless the user asks for them.

Completion means both actual files exist, are nonempty, have three pages in
the required order, and the flowchart describes the main architecture. A PNG
preview or renamed file extension is not a Visio document. If code execution,
library installation, source/icon access, or output download is unavailable,
state that exact blocker without producing a Scout handoff as a substitute.

No Cowork-to-Scout bridge, public endpoint, or unattended desktop runner is
included. Classified data must remain in approved locations;
do not emit confidential requirements into unprotected JSON or image exports.
Sharing or publishing still requires explicit approval.

CONTENTS
SKILL.md: portable agent instructions.
AzureVisio.ps1: native CRUD, model validation/roundtrip, icon catalog, exports.
Architecture-Guide.txt: requirements, input interpretation, and quality workflow.
CRUD-Guide.txt: exact create/read/update/delete contracts, approvals, and limits.
Reference-Workflow.txt: source persistence, resume behavior and visual fidelity.
Reference-Fidelity.ps1: validates a model against a persisted source contract.
Reference-Fidelity.Tests.ps1: synthetic positive and deliberate-drift cases.
Quiet-Workflow.txt: concise chat, grouped research, and automatic Visio startup.
Enterprise-Style.ps1: visible-caption, icon/card ratio and boundary guardrails.
Install-ReferenceAtlas.ps1: official reference asset download/provenance.
architecture-references.json: curated starter index plus browse/source URLs.
Import-Draft.ps1: neutral structured draft extraction, not automatic Azure mapping.
Draft-Import.Tests.ps1: extraction checks using plain PowerShell assertions.
Install-IconLibrary.ps1: official collection downloader/indexer.
icon-sources.json: official source/download URLs and collection freshness.
New-ReferenceModel.ps1: generates the generic two- or three-spoke model.
hub-spoke-reference.json: generic illustrative reference, no tenant data.
environment.example.json: empty optional local mapping.
portable_visio.py: direct editable Visio/PDF generation in a Python environment.
portable_reference.py: portable source-contract fidelity gate for conversions.
requirements-portable.txt: declared dependencies for direct portable rendering.
Developer tests remain in the source repository, not the small runtime ZIP:
Test-AzureVisio.ps1, Draft-Import.Tests.ps1, Reference-Fidelity.Tests.ps1, and
test_portable_visio.py and test_portable_reference.py. This reserves the companion-file budget
for runtime code and guides.
Build-AzureVisioPackage.ps1: allow-listed packaging, excludes environment.json.
README.txt: these installation, compatibility, and use instructions.

MAINTAINER VALIDATION
Use the full source repository for developer tests; runtime ZIPs omit them.
Run Test-AzureVisio.ps1 -PortableOnly -OutputDirectory <new absolute folder>
for filesystem/path checks without Visio.
Use -CrudOnly for the new native update/delete/read-model exercise.
Use -EnterpriseOnly -IconDirectory <catalog directory> for icon-first captions,
invisible anchors, native updates, source-detail preservation and roundtrip.
Optionally supply -ReferenceModelPath <actual model.json> and -ReferencePath
<source contract.json> together. The same runner tests the real reference and
deliberate omission/custom-adapter/replaced-page failures.
For a complete exercise, open Visio and run Test-AzureVisio.ps1 with a NEW
absolute OutputDirectory. This creates only its own example, adds a third
spoke, exercises edits, saves/reopens, and exports. It leaves that example
open. Do not run this repeatedly as ordinary skill use.

To rebuild a new archive from the installed skill folder:
.\Build-AzureVisioPackage.ps1 -SkillDirectory $PWD -OutputPath <new absolute ZIP path>

PUBLIC PRODUCT DOCUMENTATION
https://learn.microsoft.com/en-us/microsoft-scout/use-microsoft-scout
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/cowork-customize
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/cowork-faq
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/cowork-plugin-development
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/use-cowork
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/cowork-local-browser
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/cowork-plugin-development

Compatibility guidance reviewed September 5, 2026. Cowork documents skill ZIP
uploads and companion files, but not direct local PowerShell/Visio COM execution.
Cowork imports have not been exercised by this local integration harness.
