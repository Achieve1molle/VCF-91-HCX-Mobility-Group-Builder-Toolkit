# VCF 9.1 HCX Mobility Group Builder Toolkit Rev 2.2

PowerShell 7 and WPF toolkit for preparing, validating, previewing, and saving VMware HCX 9.1 Mobility Group drafts through a controlled two-phase workflow.

**Primary script:** `VCF_9_1_HCX_Mobility_Group_Builder_Toolkit_Rev_2.2.ps1`  
**Target platform:** VMware Cloud Foundation 9 with HCX 9.1  
**Execution environment:** Windows, PowerShell 7, interactive STA session  
**Primary interfaces:** HCX REST API, source vCenter, destination vCenter, VCF PowerCLI, and SPBM

## Purpose

The toolkit separates migration preparation from HCX draft creation:

1. **Phase I** discovers source VMs, validates destination selections, and exports an authoritative CSV.
2. **Phase II** imports that CSV, preserves the validated selections, supports optional final changes, builds payloads, and saves HCX drafts.

Saving a draft does not start a migration. Review every draft in HCX and follow approved change, validation, cutover, and rollback procedures.

## Key capabilities

### Inventory and topology

- HCX 9.1 REST authentication.
- Source and destination vCenter connections.
- Automatic current-session HCX topology, direction, site-pair, and Service Mesh discovery.
- Source VM, NIC, disk, power-state, and vTPM discovery.
- Destination host, cluster, datastore, StoragePod, folder, network, and SPBM policy inventory.

### Phase I

- Case-insensitive VM-name matching.
- Detailed, scrollable, virtualized VM Import Summary.
- Red highlighting for missing, ambiguous, or failed vTPM discovery.
- Gold highlighting for powered-off VMs.
- Automatic vTPM detection from source vCenter. No vTPM input column is required.
- Separate standard and vTPM storage-policy defaults.
- Effective per-VM storage-policy visibility and overrides.
- **Same format as source** as the disk-format default.
- Optional source-to-destination network-mapping CSV.
- Per-VM and per-NIC network overrides.
- Fail-closed validation before CSV creation.

### Phase II

- Restoration of group, compute, folder, storage, policy, disk, vTPM, and network selections.
- Collapsed optional Destination Settings section for approved changes after Phase I.
- Standard and vTPM policy selectors.
- Per-VM datastore, storage-policy, and disk-format overrides.
- Large scrollable VM grid with row virtualization and frozen identifying columns.
- Multi-group support from 1 through 50.
- Local payload preview before HCX submission.
- Authentication refresh before submission and one retry after an authorization failure.
- Sequential draft creation only after all selected payloads build successfully.

### Evidence and safety

- Timestamped per-launch run folders.
- Operational log, transcript, and debug artifacts.
- Payload, request, and response JSON.
- Host and StoragePod audit JSON inside the active run folder.
- Password-free connection profile.
- Password retention only while the application is open, with purge on close.
- Diagnostic sanitization for known secret fields and current password values.

## Requirements

- Windows administrative workstation with an interactive desktop session.
- PowerShell 7 or later.
- STA execution and WPF support.
- VCF.PowerCLI.
- HTTPS connectivity to the source HCX Manager and both vCenter systems.
- Accounts permitted to read required inventories and create HCX Mobility Group drafts.
- Write access to the configured output location.

## Installation and launch

```powershell
Set-Location C:\Script
pwsh.exe -NoProfile -ExecutionPolicy Bypass -STA `
    -File ".\VCF_9_1_HCX_Mobility_Group_Builder_Toolkit_Rev_2.2.ps1"
```

The script includes local self-signing support. Use the approved enterprise signing process where required.

## Phase I workflow

### 1. Connect and load inventory

Enter HCX, source-vCenter, and destination-vCenter connection data. Select **Connect and Load Inventory** and review topology, direction, Service Mesh, and inventory results.

### 2. Optionally import a network-mapping CSV

Use a network-mapping CSV when source and destination network names differ. It may be imported before or after the VM list.

```csv
SourceNetworkName,DestinationNetworkName
Legacy-App-Network,New-App-Network
Legacy-Database-Network,New-Database-Network
```

Accepted source headings include `SourceNetworkName`, `SourceNetwork`, `Source`, and `SourcePortGroup`. Accepted destination headings include `DestinationNetworkName`, `DestinationNetwork`, `Destination`, and `DestinationPortGroup`.

The importer rejects missing fields, duplicates, conflicts, unresolved destination names, and ambiguous destination names.

Mapping precedence:

1. Per-VM override.
2. Imported mapping CSV.
3. Automatic exact-name match.
4. Unresolved.

### 3. Import the VM CSV

```csv
VMName,MobilityGroupNumber
appserver01,1
secureapp01,2
```

`VMName` or `Name` is required. `MobilityGroupNumber` is optional. vTPM is detected from live source-vCenter inventory.

### 4. Review the import summary

The summary reports every VM's inventory state, power state, vTPM state, detection result, and overall finding. Import continues for missing and powered-off VMs, but validation determines whether an included row can be exported.

### 5. Apply global settings

Configure:

- Destination Site.
- Compute.
- Datastore or Datastore Cluster.
- Standard Storage Policy.
- Folder.
- Migration Type.
- vTPM Storage Policy.
- Disk Format.

Disk Format defaults to **Same format as source**.

### 6. Review per-VM values

Review group number, effective policy, policy assignment, compute, folder, storage, disk format, and every NIC mapping. Explicit per-VM policy choices are protected from later global-policy application.

### 7. Validate and export

Select **Validate**. Resolve every failure, then select **Create Mobility Group CSV**.

## Automatic vTPM behavior

Detection order:

1. `Get-VTpm` when available.
2. VM hardware-device inspection as a fallback.

A failed lookup remains unknown and blocks validation. It is not silently treated as `vTPM=0`.

- Detected vTPM VM: vTPM policy.
- VM without vTPM: standard policy.
- Individual policy selection: per-VM override.

## Phase II workflow

1. Import the approved Phase I CSV.
2. Set Base Group Name and group count.
3. Keep optional Destination Settings collapsed unless an approved change is required.
4. Review migration and switchover options.
5. Review the scrollable Phase II VM grid.
6. Validate.
7. Preview all payloads.
8. Review payload and audit JSON.
9. Save drafts to HCX.
10. Verify all drafts in HCX before migration execution.

Phase II supports per-VM changes to destination storage, effective storage policy, and disk format. Phase II overrides are retained when destination defaults are reapplied.

## CSV contracts

### Phase I output

Principal fields include:

- SchemaVersion
- HCXManager
- MobilityGroupNumber
- VMName and VMId
- PowerState
- Source compute, folder, datastore, and network information
- MigrationType
- Destination site, compute, folder, datastore, and storage type
- DestinationStoragePolicy
- vTPM
- VtpmDetectionStatus
- VtpmDetectionSource
- StoragePolicyAssignment
- DiskFormat
- DiskFormatSelectionSource
- DestinationNetworkMappingsJson
- NetworkMappingCsvPath
- ValidationStatus
- ValidatedOn

`DestinationNetworkMappingsJson` preserves adapter, source and destination names, IDs, types, match status, match detail, and mapping source.

## Placement rules

```text
Compute:
  host       host-<integer>
  cluster    domain-c<integer>

Storage:
  datastore  datastore-<integer>
  storagepod group-p<integer>
```

The disk-format default is `Same format as source`, which is converted to `sameAsSource` in the payload.

Standard VMs can inherit group-default storage. vTPM VMs and explicit per-VM policy exceptions can carry per-VM storage data.

## Logging and evidence

Each launch creates:

```text
HCX91-MobilityCSV-Run-YYYYMMDD-HHMMSS
```

Typical contents:

```text
HCX91-MobilityCSV-*.log
HCX91-PowerShell-Transcript-*.log
Debug-Artifacts\*.json
HCX91-<Group>-Payload.json
HCX91-<Group>-Request.json
HCX91-<Group>-Response.json
HCX91-Host-StoragePod-Audit-*.json
HCX91-ComputeId-Normalization-*.json
```

Host and StoragePod audits are written to the active run folder, not the script root.

## Credential handling

Passwords remain in password controls while the application is open so reconnect, authentication refresh, validation, and submission can reuse the current entries. On close, password controls, HCX session data, and authorization headers are cleared. Passwords are not written to the connection profile, CSV, payload, or intended audit output.

## Troubleshooting

### Network-mapping CSV parser error

Use the released Rev 2.2 file. PowerShell variables followed by a colon must be delimited:

```powershell
"CSV line ${line}: $($resolved.Detail)"
```

### Destination network cannot be resolved

Confirm that the destination name exists uniquely in current destination inventory. Use the per-VM **Configure** dialog for an approved exception.

### vTPM detection fails

Confirm source-vCenter hardware-read permissions and review the run log. Unknown vTPM status blocks validation.

### Phase II grid shows too few VMs

Maximize the window and keep Destination Settings collapsed. Use the grid's vertical and horizontal scrollbars.

### Passwords clear during the workflow

Confirm that `VCF_9_1_HCX_Mobility_Group_Builder_Toolkit_Rev_2.2.ps1` is running. Passwords should clear only when the application closes.

### Audit JSON appears in the script root

Confirm the final Rev 2.2 file is in use and that the active run directory is writable.

## Release notes

### Rev 2.2

- Final script name: `VCF_9_1_HCX_Mobility_Group_Builder_Toolkit_Rev_2.2.ps1`.
- Added automatic vTPM discovery and removed manual vTPM input.
- Added standard and vTPM policy tracks.
- Added effective per-VM policy display, override, validation, and export.
- Added Phase II vTPM policy and per-VM storage controls.
- Added detailed, scrollable, color-coded VM import reporting.
- Defaulted disk format to Same format as source in both phases.
- Added Phase I disk-format selection and CSV persistence.
- Added optional source-to-destination network-mapping CSV.
- Added duplicate, conflict, unresolved, and ambiguous mapping validation.
- Added mapping precedence, clear, example, provenance, and per-VM override behavior.
- Corrected PowerShell mapping-validation interpolation.
- Added collapsible optional Phase II Destination Settings.
- Expanded and virtualized the Phase II grid with scrollbars and frozen columns.
- Retained credentials for the application lifetime and purged them on close.
- Routed Host and StoragePod audits into the active run folder.
- Preserved current-session HCX topology discovery, multi-group preview, evidence generation, and sequential draft creation.

## Repository layout

```text
/
├── VCF_9_1_HCX_Mobility_Group_Builder_Toolkit_Rev_2.2.ps1
├── README.md
├── Wiki.md
├── examples/
│   ├── vm-import.example.csv
│   ├── network-mapping.example.csv
│   └── phase1-output.example.csv
├── screenshots/
│   ├── phase1-prepare.png
│   ├── vm-import-summary.png
│   ├── network-mapping.png
│   └── phase2-create.png
└── docs/
    ├── payload-example.json
    └── troubleshooting.md
```

Do not publish production credentials, customer-specific CSV files, connection profiles, payloads, logs, or infrastructure identifiers without approved sanitization.
