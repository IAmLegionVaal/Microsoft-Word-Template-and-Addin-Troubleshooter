#requires -Version 5.1
<# Created by Dewald Pretorius. #>
[CmdletBinding(SupportsShouldProcess=$true)]
param(
    [ValidateSet('Diagnose','ResetNormalTemplate')]
    [string]$Action='Diagnose',
    [string]$OutputPath=(Join-Path ([Environment]::GetFolderPath('Desktop')) 'Word_Template_Repair')
)
$ErrorActionPreference='Stop'
$NormalTemplate="$env:APPDATA\Microsoft\Templates\Normal.dotm"
$StartupFolder="$env:APPDATA\Microsoft\Word\STARTUP"
New-Item -ItemType Directory -Path $OutputPath -Force|Out-Null
$Stamp=Get-Date -Format 'yyyyMMdd_HHmmss'
$LogPath=Join-Path $OutputPath "Repair_$Stamp.log"
function Write-RepairLog([string]$Message){$Line='{0:u} {1}' -f (Get-Date),$Message;Write-Host $Line;Add-Content -LiteralPath $LogPath -Value $Line}
$State=[ordered]@{Action=$Action;WordRunning=[bool](Get-Process WINWORD -ErrorAction SilentlyContinue);NormalTemplateExists=(Test-Path -LiteralPath $NormalTemplate);StartupAddins=@(Get-ChildItem -LiteralPath $StartupFolder -ErrorAction SilentlyContinue|Select-Object Name,Length,LastWriteTime)}
$State|ConvertTo-Json -Depth 5|Set-Content -LiteralPath (Join-Path $OutputPath "PreRepair_$Stamp.json") -Encoding UTF8
if($Action -eq 'Diagnose'){Write-RepairLog '[COMPLETE] Read-only snapshot saved.';exit 0}
try{
    if(Get-Process WINWORD -ErrorAction SilentlyContinue){throw 'Close Microsoft Word before resetting the Normal template.'}
    if(Test-Path -LiteralPath $NormalTemplate){
        $Backup=Join-Path $OutputPath "Normal_$Stamp.dotm"
        if($PSCmdlet.ShouldProcess($NormalTemplate,"Preserve as $Backup and allow Word to create a clean template")){
            Copy-Item -LiteralPath $NormalTemplate -Destination $Backup -Force
            Rename-Item -LiteralPath $NormalTemplate -NewName "Normal.dotm.disabled-$Stamp"
            Write-RepairLog "[BACKUP] $Backup"
        }
    }
}catch{Write-RepairLog "[FAILED] $($_.Exception.Message)";exit 5}
Write-RepairLog '[COMPLETE] Repair completed.'
exit 0
