# PowerShell: Mail Request

## Install

```powershell
$APP = "mail-request"; $ORG = "pkgstore"; $PFX = "pwsh-"; $URI = "https://raw.githubusercontent.com/${ORG}/${PFX}${APP}/refs/heads/main"; $PSV = "7.4"; if ($PSVersionTable.PSVersion -ge $PSV) { $TS = (Get-Date -UFormat "%s"); $META = Invoke-RestMethod -Uri "${URI}/meta.json"; $META.install.file.ForEach({ if (-not (Test-Path -LiteralPath "$($_.path)")) { New-Item -Path "$($_.path)" -ItemType "Directory" | Out-Null }; if (Test-Path -LiteralPath "$($_.path)\$($_.name)") { Compress-Archive -LiteralPath "$($_.path)\$($_.name)" -DestinationPath "$($_.path)\$($_.name).${TS}.zip" }; Invoke-WebRequest "${URI}/$($_.name)" -OutFile "$($_.path)" }) } else { Write-Host "Please update to PowerShell ${PSV} or later to run this script correctly." }
```

## Resources

- [Documentation (RU)](https://libsys.ru/ru/2025/12/91f3c9a4-e6a8-5403-b42b-7004f234bff2/)
