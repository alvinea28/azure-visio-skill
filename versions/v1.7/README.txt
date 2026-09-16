AZURE VISIO 1.7 - AUTOMATIC OFFLINE CREATION

Create a three-page editable Visio and matching PDF in Cowork or Scout.
All 715 official PNG icons are supplied as normal JSON companions.
No runtime ZIP download, SVG choice, pip, network, or Scout handoff.
Required preinstalled libraries: ReportLab and Pillow; the skill cannot grant
execution, data access, output download, or installation permissions.

INSTALL ONCE; DON'T ASK EVERY USER TO DOWNLOAD A ZIP
For colleagues: install the current skill once, then use Cowork's skill detail
page > Share > Specific users in your organization. Re-share after updates.
Recipients obtain the shared skill through Cowork rather than fetching GitHub ZIPs.
For managed deployment: an administrator can deploy a skills-only plugin to
assigned users/groups. Recipient acquisition is automatic; host approvals remain.
No plugin manifest, tenant permissions or admin deployment are silently created
by the Python renderer. See DISTRIBUTION.txt in the source for documented routes.

For an individual first installation, Cowork currently documents
Customize (sidebar or +) > Skills > Add dropdown > Upload skill.
Select azure-visio-1.7.0.zip. SKILL.md and every companion are at the archive root.
Do not upload just SKILL.md, GitHub's source archive, or an old installation ZIP.
Wait for Cowork's installation synchronization and start a NEW task.
The host's initial import/trust approval is separate from runtime setup and
cannot be silently approved by the skill being installed.

Scout can register the same skill and all its companions, then start a new chat.
This is the portable new-drawing path, not a change to existing desktop CRUD.

NORMAL USER PROMPT
Use azure-visio to create this architecture: [requirements].
Deliver the three-page editable Visio and PDF here.

The assistant accesses the installed companions automatically, runs startup,
chooses catalog IDs, creates the model, allocates new output paths and delivers.
It must not ask for an SVG renderer, icon ZIP, catalog path, output filename or
continue/confirmation of routine first-run steps.
Real host approvals, missing essential sources and sharing consent still apply.

AUTOMATIC RUNTIME COMMANDS (FOR THE ASSISTANT)
python -B portable_visio.py startup
python -B portable_visio.py catalog --search "vault"
python -B portable_visio.py render --model ABSOLUTE_MODEL_PATH --bundle

The script finds offline-assets.json beside itself, independent of process cwd.
The manifest hashes catalog.json and Icon-Data-1.json / Icon-Data-2.json. Ordinary
base64 PNG data is decoded in memory; no nested ZIP is fetched or extracted.
All companions are accessed through the host's normal skill-resource tools.
Output gets a unique directory beside the model. Use --output-root with a
host-provided task folder when rendering the bundled example; do not write into
installed resources or ask the user to choose a working directory.
The renderer returns ready, complete, or blocked; never waiting-for-input.
Missing or corrupt companions fail immediately. Don't retry missing downloads.
The geometry/routing phase has a 120-second budget. Other host/tool delays are
outside this renderer's control; this is not a guarantee of total task duration.

WHY PREVIOUS RUNS ASKED OR STOPPED
The public 1.6 download required resvg-py and an external SVG icon catalog.
Instructions told Cowork to install dependencies and obtain icons. On hosts with
no network/package installation, that became setup questions or a blocked task.
The earlier offline candidate was local-only and never replaced that public link.
The renderer had no stdin/input wait loop; an individual reported hang cannot be
diagnosed beyond this without its log. Version 1.7 removes those setup branches,
publishes the bundled assets, and explicitly disallows nonessential input waits.

DELIVERABLES AND LIMITS
Page 1: requested main architecture.
Page 2: applicable hardening proposal, otherwise a substantive controls review.
Page 3: connected flowchart and detailed write-up of page 1.
Native captions, shapes, metadata and glued right-angle connectors stay editable.
PNG glyphs are at 256-pixel longest-edge resolution; do not claim vector pixels.
Visible text supports Windows-1252, not arbitrary fonts/scripts. Boundaries have
semantic hierarchy rather than native Visio container membership.
Python creates new files; it does not control desktop Visio.
No requirements, private customer files or downloaded reference diagrams belong
in a reusable skill package. Follow classification and source usage rights.
Original artwork provenance and Microsoft usage terms are in NOTICE.txt.

DOCUMENTED HOST ROUTES
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/cowork-customize
https://learn.microsoft.com/en-us/microsoft-365/copilot/cowork/cowork-manage-plugins

Independent project, not an official Microsoft or GitHub product.
