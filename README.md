# Azure Visio Skill

Create editable Azure architecture **Visio + PDF** packs with Scout or Copilot Cowork: **main architecture**, **enterprise hardening/review**, and **detailed flow explanation**.

**Current release: [v1.7.0 skill package](versions/v1.7/dist/azure-visio-1.7.0.zip)** | [Setup](versions/v1.7/README.txt) | [Distribution](versions/v1.7/DISTRIBUTION.txt)

**Automated first run:** 715 bundled icons, automatic resource discovery and output folders, no additional ZIP/SVG downloads, package installs, or routine setup questions.

## Install in Microsoft Scout

**Prerequisites:** Scout with skill registration and permitted local file/Python execution, **Python 3.10+**, and preinstalled **ReportLab and Pillow**. v1.7 creates new Visio/PDF files directly with Python; desktop Visio is **not required to generate them**. Use a compatible Visio application to open and edit the resulting `.vsdx`. Existing Scout desktop-Visio editing workflows remain separate.

1. Get the **[v1.7.0 skill ZIP](versions/v1.7/dist/azure-visio-1.7.0.zip)**, not GitHub's **Code > Download ZIP** source archive.
2. Extract it to a local folder, preferably outside OneDrive, and keep **all 11 files together**. For example, open `%LOCALAPPDATA%` in File Explorer and create `AzureVisio\1.7` there. Copy the extracted folder's full path.
3. In a Scout chat, paste this installation prompt, replacing `[full folder path]`:

   > Install the azure-visio skill from [full folder path]. Register its SKILL.md using your skill-management tools and copy every companion file into the registered skill's resource directory, preserving their filenames. Update the existing azure-visio registration if present rather than creating a duplicate, and preserve unrelated skills/settings. Run portable_visio.py startup using permitted Python execution. Use the bundled icons; do not download SVGs or install packages. Report any missing prerequisite clearly.

4. Approve any legitimate Scout file/execution access requests. Once registration completes, start a **new Scout chat** so the updated skill loads.
5. Attach your requirements or architecture reference and use the prompt below. If slash commands are available, `/azure-visio` invokes the skill; otherwise say **"Use azure-visio"**.

If Python or its base libraries are missing, have them provisioned through your organization's approved setup process. The skill does not install dependencies or bypass policy. You do not need to supply an icon pack, SVG files, or an output folder.

## Use in Cowork

The User imports the current package once through **Customize > Skills > Add > Upload skill**, keeping all companions. Start a new task after installation/update.

Cowork completes both files directly with preinstalled **ReportLab and Pillow**, **without a Scout handoff**. Initial host trust/permission approvals still apply; no skill can silently grant them.

## Create your first diagram

Attach your file(s) to the chat, or give Scout an accessible full local path. For OneDrive/SharePoint documents, supply the file link and let the host use its approved Microsoft 365 reading tools.

> Use azure-visio to create an architecture from the attached [filename]. Treat it as [requirements / an existing architecture to preserve]. Keep the supplied components, quantities, regions, boundaries and relationships; do not invent missing connections. Record unresolved assumptions clearly. Deliver an editable three-page Visio and matching PDF: main architecture, enterprise hardening or review, and a detailed explanation of the main flow. Use official icons and right-angle connectors.

The assistant reads the source, creates the internal model, finds the bundled icons, and saves both files in a new output folder. **You do not need to write JSON yourself.** Missing essential information or unreadable/protected sources may still need clarification.

## What can I supply?

| Source example | What to include | Important limit |
| --- | --- | --- |
| Requirements document: `requirements.docx`, `.pdf`, `.md`, or `.txt` | Business goal, users, services, request/data flows, networking, security, region and recovery requirements. | Reading depends on the host's available document tools and your usage rights. |
| Sketch or existing architecture: `.png` or `.jpg` | A legible screenshot/photo with component names, arrows, boundaries and a legend. Say whether to preserve it or propose improvements. | Cropped text, ambiguous arrows and crossings cannot be reliably inferred; this is not a guaranteed OCR importer. |
| Bill of materials: `BOM.xlsx` or `.csv` | Resource/service, SKU, quantity, region, environment, purpose and known connections; include a dependency sheet if available. | A BOM lists inventory and cost, not necessarily topology. Quantities alone do not prove HA or connectivity. |
| Architecture model: `architecture.json` | A model using the skill's schema, such as the **[bundled example model](versions/v1.7/example-model.json)**. | This is the renderer's direct input format. Other JSON inventories must first be interpreted and mapped to that model. |
| Azure ARM export: `template.json` and sanitized parameters | Exported resource definitions, scopes, properties, references and relevant parameter values. | An ARM template is not a complete traffic-flow diagram. `dependsOn` means deployment ordering, not necessarily runtime communication. Templates are read as data, never deployed. |
| Azure topology export or screenshot | Resource IDs/types and explicit connections, or a readable Azure topology screenshot. | A topology view and an ARM deployment template are different sources; neither should be assumed to show all application/data flows. |

For an ARM source, use **Azure portal > resource group > Export template > Download**, then attach the relevant JSON files after removing secrets. See Microsoft's [export guidance and limitations](https://learn.microsoft.com/en-us/azure/azure-resource-manager/templates/export-template-portal). Add a topology screenshot or flow description when the template does not establish the connections.

These are **source references interpreted through the host's permitted tools**, not a promise that the renderer natively imports every Office, image or Azure export format. Multiple files can be supplied together; identify which is authoritative if they disagree. Keep sensitive customer documents in approved locations, not this public repository.

### Example requirements to paste or save as a document

> Design an illustrative Azure web application. Users access Azure App Service over HTTPS. The application reads and writes Azure SQL Database and stores uploaded files in Azure Blob Storage. Include Microsoft Entra ID authentication and Azure Monitor telemetry. Region, SKU, capacity and recovery targets are not yet selected; mark them as open decisions rather than inventing values. Keep any additional hardening recommendations separate from the requested main design.

For a ready-made schema example, attach **[example-model.json](versions/v1.7/example-model.json)** and say: **"Use azure-visio to render this supplied model into editable Visio and PDF files."** It contains a synthetic RAG assistant, not a default design for other requirements.

## Same requirements, two platforms

The same **Fabric IQ Pilot requirements document** was sent to **Copilot Cowork** and **Microsoft Scout**. These two screenshots show the architecture generated by each platform.

### Copilot Cowork

![Fabric IQ Pilot architecture generated by Copilot Cowork](sample%20output/fabric-iq-pilot-cowork.png)

### Microsoft Scout

![Fabric IQ Pilot architecture generated by Microsoft Scout](sample%20output/fabric-iq-pilot-scout.jpeg)

Only the two public-release screenshots are shared for this comparison, not their source requirements, PDFs, or Visio files. Illustrative outputs, not a deployment or compliance certification.

Independent project, not an official Microsoft or GitHub product. Microsoft icons retain their [published terms](https://learn.microsoft.com/en-us/azure/architecture/icons/).
