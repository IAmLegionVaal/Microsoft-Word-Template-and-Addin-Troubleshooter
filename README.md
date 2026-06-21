# Microsoft Word Template and Add-in Troubleshooter

Created by **Dewald Pretorius**.

`Troubleshooter.ps1` collects template, Normal.dotm, COM add-in, startup-folder, and automation evidence. `Repair.ps1` adds read-only diagnosis and a guarded `ResetNormalTemplate` action.

```powershell
.\Repair.ps1 -Action Diagnose
.\Repair.ps1 -Action ResetNormalTemplate -WhatIf
.\Repair.ps1 -Action ResetNormalTemplate -Confirm
```

Word must be closed. The workflow copies `Normal.dotm` into the repair output folder and renames the original so Word can generate a clean template on next launch. Startup add-ins are inventoried but not changed. Source-reviewed; not runtime-tested against every Word build or add-in set.
