# PowerShell: Mail Request

## Install

```powershell
$App = 'mail-request'; $Ver = 'v0.0.0'; Invoke-Command -ScriptBlock $([scriptblock]::Create((Invoke-WebRequest -Uri 'https://pkgstore.ru/pwsh.install.txt').Content)) -ArgumentList ($args + @($App,$Ver))
```

## Resources

- [Documentation (RU)](https://libsys.ru/ru/2025/12/91f3c9a4-e6a8-5403-b42b-7004f234bff2/)
