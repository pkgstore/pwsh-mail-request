<#PSScriptInfo
.VERSION      0.1.0
.GUID         07371aa0-dc86-417d-8981-517226c40c6c
.AUTHOR       Kai Kimera
.AUTHOREMAIL  mail@kaikim.ru
.TAGS         windows server mail
.LICENSEURI   https://choosealicense.com/licenses/mit/
.PROJECTURI
#>

#Requires -Version 7.2

<#
.SYNOPSIS
Sends an email notification using SMTP.

.DESCRIPTION

.EXAMPLE
.\app.mail.request.ps1 -Domain 'example.org' -From 'request@example.com' -Request 'C:\Request\*.txt'

.EXAMPLE
.\app.mail.request.ps1 -Domain 'example.org' -From 'request@example.com' -Request 'C:\Request\*.txt' -Cc 'mail@example.net', 'mail@example.biz'

.EXAMPLE
.\app.mail.request.ps1 -Domain 'example.org' -From 'request@example.com' -Request 'C:\Request\*.txt' -Bcc 'mail@example.net', 'mail@example.biz'

.LINK

#>

# -------------------------------------------------------------------------------------------------------------------- #
# CONFIGURATION
# -------------------------------------------------------------------------------------------------------------------- #

param(
  [Parameter(Mandatory)][string]$Domain,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][Parameter(Mandatory)][string]$From,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Cc,
  [ValidatePattern('^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{1,}$')][string[]]$Bcc,
  [SupportsWildcards()][string[]]$Request,
  [ValidateSet('Low', 'Normal', 'High')][string]$Priority = 'Normal',
  [switch]$Attach,
  [switch]$Save,
  [switch]$HTML,
  [switch]$SSL,
  [switch]$BypassCertValid
)

$Cfg = ((Get-Item "${PSCommandPath}").Basename + '.ini')
$P = (Get-Content -Path "${PSScriptRoot}\${Cfg}" | ConvertFrom-StringData)
$Log = "${PSScriptRoot}\log.mail.request.txt"
$Request = (Resolve-Path "${Request}" | Select-Object -ExpandProperty 'Path'); if ($null -eq $Request ) { exit }

# -------------------------------------------------------------------------------------------------------------------- #
# -----------------------------------------------------< SCRIPT >----------------------------------------------------- #
# -------------------------------------------------------------------------------------------------------------------- #

function Remove-Request([string]$Path) {
  Remove-Item -LiteralPath "${Path}"
}

function Send-Mail {
  try {
    $Request.ForEach({
      $Mail = (New-Object System.Net.Mail.MailMessage)
      $Mail.Subject = ((Get-Content "${_}" | Select-Object -Skip 1 -First 1) -replace '^(Subject: )','')
      $Mail.Body = (Get-Content "${_}" | Select-Object -Skip 2)
      $Mail.BodyEncoding= [System.Text.Encoding]::UTF8
      $Mail.From = $From
      $Mail.Priority = $Priority
      $Mail.IsBodyHtml = $HTML
      $To = (((Get-Content "${_}" | Select-Object -First 1) -replace '^(To: )','') -split ',')
      $To.ForEach({ $Mail.To.Add(-join($_,'@', $Domain)) })
      $Cc.ForEach({ $Mail.CC.Add($_) })
      $Bcc.ForEach({ $Mail.BCC.Add($_) })
      if ($Attach) { $Mail.Attachments.Add((New-Object System.Net.Mail.Attachment($_))) }
      if ($BypassCertValid) { [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true } }

      $RequestName = (Split-Path -Path "${_}" -LeafBase)
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
  Start-Transcript -Path "${Log}"
  Send-Mail
  Stop-Transcript
}; Start-Script
