# PowerShell: Mail Request

## Install

```powershell
$APP = "mail-request"; $ORG = "pkgstore"; $PFX = "pwsh-"; $URI = "https://raw.githubusercontent.com/${ORG}/${PFX}${APP}/refs/heads/main"; $META = Invoke-RestMethod -Uri "${URI}/meta.json"; $META.install.file.ForEach({ if (-not (Test-Path "$($_.path)")) { New-Item -Path "$($_.path)" -ItemType "Directory" | Out-Null }; Invoke-WebRequest "${URI}/$($_.name)" -OutFile "$($_.path)" })
```

## Resources

- [Documentation (RU)](https://libsys.ru/ru/2025/12/91f3c9a4-e6a8-5403-b42b-7004f234bff2/)
