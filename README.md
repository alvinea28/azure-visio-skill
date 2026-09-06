# Azure Visio Skill

Create editable Azure architecture **Visio + PDF** files with Scout or Copilot Cowork. Each v1.6 pack has three pages: **main architecture**, **conditional enterprise hardening/review**, and **detailed main-flow explanation**.

**[Download v1.6 ZIP](versions/v1.6/dist/azure-visio-1.6.0.zip)** | [v1.6 files](versions/v1.6/) | [v1.5 archive](versions/v1.5/)

Code, guides, tests, downloads, and screenshots are grouped inside each version folder.

## Install

**Scout:** Requires licensed desktop Visio on Windows. Extract the ZIP, keep its files together, and ask:

> Install azure-visio from [folder]. Register SKILL.md and all companion files, and complete the official icon setup.

**Cowork:** Open **Customize** (sidebar or **+**) > **Skills > Add > Upload skill** and upload the v1.6 ZIP, not GitHub's source archive. Cowork creates the files directly with Python, **without a Scout handoff**. Skill upload, Python/dependencies, and file downloads must be allowed by your organization's rollout and policy.

Start a **new chat/task** after installation.

## Try it

> Use azure-visio to draw [your requirements]. Deliver the three-page editable Visio and PDF. Use official icons and right-angle connectors. In Cowork, finish here without a Scout handoff.

[Setup and commands](versions/v1.6/README.txt) | [Editing guide](versions/v1.6/CRUD-Guide.txt)

## Latest sample: AVS and enterprise backup

Azure VMware Solution with multi-region active/passive disaster recovery, enterprise landing zones, security, backup, and monitoring.

[Sample PDF](versions/v1.6/examples/AVS-Enterprise-Backup.pdf) | [Editable Visio sample](versions/v1.6/examples/AVS-Enterprise-Backup.vsdx)

Both files include all three pages: AVS main architecture, enterprise landing zones, and flow/write-up. Illustrative target design, not a deployed or certified solution.

### Enterprise AVS and backup view (page 2)

![AVS enterprise landing zones with security, backup, and monitoring](versions/v1.6/examples/images/avs-enterprise-backup.png)

### Flow and write-up (page 3)

![AVS protection, regional failover, recovery, and return flow with explanatory write-up](versions/v1.6/examples/images/avs-flow-writeup.png)

Independent project, not an official Microsoft or GitHub product. Microsoft icons retain their [published terms](https://learn.microsoft.com/en-us/azure/architecture/icons/).
