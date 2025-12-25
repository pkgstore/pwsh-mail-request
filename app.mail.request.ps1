<#PSScriptInfo
.VERSION      0.2.0
.GUID         07371aa0-dc86-417d-8981-517226c40c6c
.AUTHOR       Kai Kimera
.AUTHOREMAIL  mail@kaikim.ru
.TAGS         windows server mail
.LICENSEURI   https://choosealicense.com/licenses/mit/
.PROJECTURI   https://libsys.ru/ru/2025/12/91f3c9a4-e6a8-5403-b42b-7004f234bff2/
#>

#Requires -Version 7.4

<#
.SYNOPSIS
Sending emails with requests using SMTP.

.DESCRIPTION

.EXAMPLE
.\app.mail.request.ps1 -From 'request@example.com' -Request 'C:\Request\*.txt'

.EXAMPLE
.\app.mail.request.ps1 -From 'request@example.com' -Request 'C:\Request\*.txt' -Cc 'mail@example.net', 'mail@example.biz'

.EXAMPLE
.\app.mail.request.ps1 -From 'request@example.com' -Request 'C:\Request\*.txt' -Bcc 'mail@example.net', 'mail@example.biz'

.LINK
https://libsys.ru/ru/2025/12/91f3c9a4-e6a8-5403-b42b-7004f234bff2/
#>

# -------------------------------------------------------------------------------------------------------------------- #
# CONFIGURATION
# -------------------------------------------------------------------------------------------------------------------- #

param(
  [Parameter(Mandatory)][SupportsWildcards()][string[]]$Request,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][Parameter(Mandatory)][string]$From,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Cc,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Bcc,
  [ValidateSet('Low', 'Normal', 'High')][string]$Priority = 'Normal',
  [string]$LogPath = "${PSScriptRoot}\log.mail.request.txt",
  [string]$LogSize = '50MB',
  [switch]$Attach,
  [switch]$Save,
  [switch]$HTML,
  [switch]$SSL,
  [switch]$BypassCertValid
)

$Cfg = ((Get-Item "${PSCommandPath}").Basename + '.ini')
$P = (Get-Content "${PSScriptRoot}\${Cfg}" | ConvertFrom-StringData)
$N = [Environment]::NewLine
$TS = (Get-Date -UFormat '%F.%H-%M-%S')
$Date = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ssK')
$Request = (Resolve-Path "${Request}" | Select-Object -ExpandProperty 'Path')
if ($null -eq $Request ) { Write-Host 'Request not found!'; exit }

# -------------------------------------------------------------------------------------------------------------------- #
# -----------------------------------------------------< SCRIPT >----------------------------------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #

function Remove-Request([string]$Path) {
  Remove-Item -LiteralPath "${Path}"
}

function Compress-Log([string]$Path, [string]$Size) {
  if ((Get-Item $Path).Length -gt $Size) {
    Compress-Archive -LiteralPath "${Path}" -DestinationPath "${Path}.${TS}.zip" && Remove-Item -LiteralPath "${Path}"
  }
}

function Write-Sep {
  $Sep = switch ( $true ) {
    $HTML   { -join ('<br><br>', '<hr style="border:none;border-top:1px solid #cccccc;width:100%;">') }
    default { -join ("${N}${N}-- ", "${N}") }
  }

  return $Sep
}

function Write-Sign([string]$Sign) {
  $Sign = switch ( $true ) {
    $HTML   { -join ('<ul>', "<li><code>#Request:${Sign}</code></li>", "<li><code>#Date:${Date}</code></li>", '</ul>') }
    default { -join ("#Request:${Sign}${N}", "#Date:${Date}") }
  }

  return $Sign.ToUpper()
}

function Send-Mail {
  try {
    $Request.ForEach({
      $RequestName = (Split-Path -Path "${_}" -LeafBase)
      $RequestBody = ((Get-Content -LiteralPath "${_}" | Select-Object -Skip 2) | Out-String)
      $Mail = (New-Object System.Net.Mail.MailMessage)
      $Mail.Subject = ((Get-Content -LiteralPath "${_}" | Select-String '^(Subject: (.*))').Matches[0].Groups[2].Value)
      $Mail.Body = (-join ($RequestBody, $(Write-Sep), $(Write-Sign "${RequestName}")))
      $Mail.BodyEncoding= [System.Text.Encoding]::UTF8
      $Mail.From = $From
      $Mail.Priority = $Priority
      $Mail.IsBodyHtml = $HTML
      $To = (((Get-Content -LiteralPath "${_}" | Select-String '^(To: (.*))').Matches[0].Groups[2].Value) -split ',')
      $To.ForEach({ $Mail.To.Add($_) })
      $Cc.ForEach({ $Mail.CC.Add($_) })
      $Bcc.ForEach({ $Mail.BCC.Add($_) })
      if ($Attach) { $Mail.Attachments.Add((New-Object System.Net.Mail.Attachment($_))) }
      if ($BypassCertValid) { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }

      $MailName = ($Mail.To | Join-String -Separator ', ')
      $SmtpClient = (New-Object Net.Mail.SmtpClient($P.Server, $P.Port))
      $SmtpClient.EnableSsl = $SSL
      $SmtpClient.Credentials = (New-Object System.Net.NetworkCredential($P.User, $P.Password))
      $SmtpClient.Send($Mail) && Write-Host "Request '${RequestName}' sent to: ${MailName}..."
      $Mail.Dispose()
      $SmtpClient.Dispose()
      if (-not $Save) { Remove-Request "${_}" }
    })
  } catch {
    Write-Error "ERROR: $($_.Exception.Message)"
  }
}

function Start-Script() {
  Start-Transcript -LiteralPath "${LogPath}" -Append; Send-Mail; Stop-Transcript
  Compress-Log -Path "${LogPath}" -Size "${LogSize}"
}; Start-Script
