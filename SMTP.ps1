# Connect to Exchange Online
Set-ExecutionPolicy unrestricted -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline

# Escribe en la consola, recibe 3 parámetros (Texto, ForegroundColor, numero de saltos de linea, segundos de espera)
function WriteConsole {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$Text,
    [ValidateSet("Black", "DarkBlue", "DarkGreen", "DarkCyan", "DarkRed", "DarkMagenta", "DarkYellow", "Gray", "DarkGray", "Blue", "Green", "Cyan", "Red", "Magenta", "Yellow", "White")]
    [string]$ForegroundColor = "White",
    [int]$NewLine = 0,
    [int]$Wait = 0
  )

  Write-Host $Text -ForegroundColor $ForegroundColor

  if ($NewLine -gt 0) { for ($i = 0; $i -lt $NewLine; $i++) { Write-Host "" } }

  if ($Wait -gt 0) { Start-Sleep -Seconds $Wait }
}

function IsUserExist($User) {
  $User = Get-Mailbox -Identity $User -ErrorAction SilentlyContinue
  return $null -ne $User
}

function ShowMenu() {
  Clear-Host
  WriteConsole -Text "--------------------[ SMTP Managment ]--------------------" -ForegroundColor "Yellow" -NewLine 1
  WriteConsole -Text "1- Enable SMTP authentication in the organization."
  WriteConsole -Text "2- Disable SMTP authentication in the organization."
  WriteConsole -Text "3- Show SMTP authentication status in the organization."
  WriteConsole -Text "4- Set default SMTP for all users."
  WriteConsole -Text "5- Enable SMTP for all users."
  WriteConsole -Text "6- Disable SMTP for all users."
  WriteConsole -Text "7- Enable SMTP authentication for specific user."
  WriteConsole -Text "8- Disable SMTP authentication for specific user."
  WriteConsole -Text "9- Show SMTP authentication status for a specific user."
  WriteConsole -Text "10- Send email."
  WriteConsole -Text "0- Exit."
  WriteConsole -Text "------------------------------------------------------------" -ForegroundColor "Yellow"
  return Read-Host "Select an option"
}

# Set SMTP authentication for the organization
function Set-SMTPAuthentication {
  param (
    [bool]$Value
  )

  if ($Value) {
    Set-TransportConfig -SmtpClientAuthenticationDisabled $true
    WriteConsole -Text "`nSMTP authentication has been disabled in the organization." -ForegroundColor "green" -NewLine 2
  }
  else {
    Set-TransportConfig -SmtpClientAuthenticationDisabled $false
    WriteConsole -Text "`nSMTP authentication has been enabled in the organization." -ForegroundColor "green" -NewLine 2
  }
  Pause
}
# Verify if SMTP authentication is enabled
function IsSMTPEnabled() { return !(Get-TransportConfig).SmtpClientAuthenticationDisabled }

# Set SMTP authentication for a specific user
function Set-SMTPForUser {
  [CmdletBinding()]
  param (
    [Parameter(Mandatory = $true)]
    [string]$User,
    [Parameter(Mandatory = $false)]
    [bool]$value = $null
  )
  Set-CASMailbox -Identity $user -SmtpClientAuthenticationDisabled $value
}

# Set SMTP authentication for all users
function Set-SMTPForAllMailboxes {
  param (
    [Parameter(Mandatory = $true)]
    [bool]$value
  )
  $Mailboxes = Get-Mailbox
  foreach ($Mailbox in $Mailboxes) { Set-SMTPForUser -User $Mailbox -value $value }
  if ($value) { WriteConsole -Text "`nSMTP authentication has been disabled for all users." -ForegroundColor "green" -NewLine 2 }
  else { WriteConsole -Text "`nSMTP authentication has been enabled for all users." -ForegroundColor "green" -NewLine 2 }
  Pause
}

# Inherits the organization configuration
function Set-DefaultSMTP() {
  $Mailboxes = Get-Mailbox
  foreach ($Mailbox in $Mailboxes) { Set-SMTPForUser -User $Mailbox.PrimarySmtpAddress }
  WriteConsole -Text "`nSMTP authentication has been set by default for all users." -ForegroundColor "green" -NewLine 2
  Pause
}

# Set SMTP authentication for a specific user
function Set-SMTPForSpecificMailboxes {
  param(
    [Parameter(Mandatory = $true)]
    [bool]$Value
  )
  $User = Read-Host "`nEnter the user email"
  if (IsUserExist($User)) {
    Set-SMTPForUser -User $User -value $Value
    if ($Value) { WriteConsole -Text "`nSMTP authentication has been disabled for the user $User" -ForegroundColor "green" -NewLine 2 }
    else { WriteConsole -Text "`nSMTP authentication has been enabled for the user $User" -ForegroundColor "green" -NewLine 2 }
  }
  else { WriteConsole -Text "`nThe user $User does not exist." -ForegroundColor "red" -NewLine 2 }
  Pause
}

function Show-SMTPStatus() {
  if ($IsSMTPEnabled) { WriteConsole -Text "`nSMTP authentication is enabled in the organization." -ForegroundColor "green" -NewLine 2 }
  else { WriteConsole -Text "`nSMTP authentication is disabled in the organization." -ForegroundColor "red" -NewLine 2 }
  Pause
}

function Show-SMTPStatusForSpecificMailboxes() {
  $User = Read-Host "`nEnter the user email"
  if (IsUserExist($User)) {
    $SMTPStatus = (Get-CASMailbox -Identity $User).SmtpClientAuthenticationDisabled
    if ($SMTPStatus) { WriteConsole -Text "`nSMTP authentication is disabled for the user $User" -ForegroundColor "red" -NewLine 2 }
    else { WriteConsole -Text "`nSMTP authentication is enabled for the user $User" -ForegroundColor "green" -NewLine 2 }
  }
  else { WriteConsole -Text "`nThe user $User does not exist." -ForegroundColor "red" -NewLine 2 }
  Pause
}

function SendMail() {
  $credentials = Get-Credential
  $from = $credentials.UserName
  $to = Read-Host "`nEscriba el correo del destinatario"
  $subject = Read-Host "Escriba el asunto del correo"
  $body = Read-Host "Escriba el cuerpo del correo"
  Send-MailMessage -From $from -To $to -Subject $subject -Body $body -Credential $credentials -UseSsl -SmtpServer smtp.office365.com -Port 587
  Pause
}

do {
  $Option = ShowMenu
  switch ($Option) {
    1 { Set-SMTPAuthentication -Value $false }
    2 { Set-SMTPAuthentication -Value $true }
    3 { Show-SMTPStatus }
    4 { Set-DefaultSMTP }
    5 { Set-SMTPForAllMailboxes -Value $false }
    6 { Set-SMTPForAllMailboxes -Value $true }
    7 { Set-SMTPForSpecificMailboxes -Value $false }
    8 { Set-SMTPForSpecificMailboxes -Value $true }
    9 { Show-SMTPStatusForSpecificMailboxes }
    10 { SendMail }
    0 { WriteConsole -Text "`nExiting..." -ForegroundColor "Yellow" -NewLine 2 -Wait 2 }
    default { WriteConsole -Text "`nInvalid option." -ForegroundColor "red" -NewLine 2 -Wait 3 }
  }
} while ($Option -ne 0)