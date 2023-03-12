Set-ExecutionPolicy unrestricted -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline

function WriteConsole($params) {

  $Text = $params[0]
  $ForegroundColor = $params[1]
  $NewLine = $params[2]
  $Wait = $params[3]

  if ($Text -and $ForegroundColor -and $NewLine) {
    Write-Host $Text -ForegroundColor $ForegroundColor
    for ($i = 0; $i -lt $NewLine; $i++) { Write-Host "" }
  }

  elseif ($Text -and $ForegroundColor) { Write-Host $Text -ForegroundColor $ForegroundColor }

  elseif ($Text) { Write-Host $Text }

  if ($Wait) { Start-Sleep -Seconds $Wait }
}

function IsUserExist($User) {
  $User = Get-Mailbox -Identity $User -ErrorAction SilentlyContinue
  return $null -ne $User
}

function ShowMenu() {
  Clear-Host
  WriteConsole("--------------------[ SMTP Managment ]--------------------", "green", 1, 0)
  WriteConsole @("1- Habilitar la autenticación SMTP en la organización", "white")
  WriteConsole @("2- Deshabilitar la autenticación SMTP en la organización", "white")
  WriteConsole @("3- Mostrar el estado de la autenticación SMTP en la organización", "white")
  WriteConsole @("4- Configurar SMTP por defecto para todos los usuarios", "white")
  WriteConsole @("5- Habilitar SMTP para todos los usuarios", "white")
  WriteConsole @("6- Deshabilitar SMTP para todos los usuarios", "white")
  WriteConsole @("7- Habilitar la autenticación SMTP para buzones específicos", "white")
  WriteConsole @("8- Deshabilitar la autenticación SMTP para buzones específicos", "white")
  WriteConsole @("9- Mostrar el estado de la autenticación SMTP para un usuario específico", "white")
  WriteConsole @("10- Enviar correo", "white")
  WriteConsole @("0- Salir", "white", 1)
  $Option = Read-Host "Seleccione una opción"
  return $Option
}

function EnableSMTP() {
  Set-TransportConfig -SmtpClientAuthenticationDisabled $false
  WriteConsole @("La autenticación SMTP ha sido habilitada en la organización", "green", 2, 0)
  Pause
}

function DisableSMTP() {
  Set-TransportConfig -SmtpClientAuthenticationDisabled $true
  WriteConsole @("La autenticación SMTP ha sido deshabilitada en la organización", "green", 2, 0)
  Pause
}

function IsSMTPEnabled() { return !(Get-TransportConfig).SmtpClientAuthenticationDisabled }

function setSMTP($params) {
  $user = $params[0]
  $value = $params[1]
  Set-CASMailbox -Identity $user -SmtpClientAuthenticationDisabled $value
}

function EnableSMTPForAllMailboxes() {
  $Mailboxes = Get-Mailbox
  foreach ($Mailbox in $Mailboxes) {
    setSMTP @($Mailbox.PrimarySmtpAddress, $false)
  }
  WriteConsole @("La autenticación SMTP ha sido habilitada para todos los usuarios", "green", 2, 0)
  Pause
}

function DisableSMTPForAllMailboxes() {
  $Mailboxes = Get-Mailbox
  foreach ($Mailbox in $Mailboxes) {
    setSMTP @($Mailbox.PrimarySmtpAddress, $true)
  }
  WriteConsole @("La autenticación SMTP ha sido deshabilitada para todos los usuarios", "green", 2, 0)
  Pause
}

function SetDefaultSMTP() {
  $Mailboxes = Get-Mailbox
  foreach ($Mailbox in $Mailboxes) {
    setSMTP @($Mailbox.PrimarySmtpAddress, $null)
  }
  WriteConsole @("La autenticación SMTP ha sido configurada por defecto para todos los usuarios", "green", 2, 0)
  Pause
}

function EnableSMTPForSpecificMailboxes() {
  $User = Read-Host "Escriba el correo del usuario"
  if (IsUserExist($User)) {
    setSMTP($User, $false)
    WriteConsole @("La autenticación SMTP ha sido habilitada para el usuario $User", "green", 2, 0)
  }
  else { WriteConsole @("El usuario $User no existe", "red", 2, 0) }
  Pause
}

function DisableSMTPForSpecificMailboxes() {
  $User = Read-Host "Escriba el correo del usuario"
  if (IsUserExist($User)) {
    setSMTP($User, $true)
    WriteConsole @("La autenticación SMTP ha sido deshabilitada para el usuario $User", "green", 2, 0)
  }
  else { WriteConsole @("El usuario $User no existe", "red", 2, 0) }
  Pause
}

function ShowSMTPStatus() {
  $SMTPStatus = IsSMTPEnabled
  if ($SMTPStatus) { WriteConsole @("La autenticación SMTP está habilitada en la organización", "green", 2, 0) }
  else { WriteConsole @("La autenticación SMTP está deshabilitada en la organización", "red", 2, 0) }
  Pause
}

function ShowSMTPStatusForSpecificMailboxes() {
  $User = Read-Host "Escriba el correo del usuario"
  if (IsUserExist($User)) {
    $SMTPStatus = (Get-CASMailbox -Identity $User).SmtpClientAuthenticationDisabled
    if ($SMTPStatus) { WriteConsole @("La autenticación SMTP está deshabilitada para el usuario $User", "red", 2, 0) }
    else { WriteConsole @("La autenticación SMTP está habilitada para el usuario $User", "green", 2, 0) }
  }
  else { WriteConsole @("El usuario $User no existe", "red", 2, 0) }
  Pause
}

function SendMail() {
  $credentials = Get-Credential
  $from = $credentials.UserName
  $to = Read-Host "Escriba el correo del destinatario"
  $subject = Read-Host "Escriba el asunto del correo"
  $body = Read-Host "Escriba el cuerpo del correo"
  Send-MailMessage -From $from -To $to -Subject $subject -Body $body -Credential $credentials -UseSsl -SmtpServer smtp.office365.com -Port 587
  Pause
}

do {
  $Option = ShowMenu
  switch ($Option) {
    1 { EnableSMTP }
    2 { DisableSMTP }
    3 { ShowSMTPStatus }
    4 { SetDefaultSMTP }
    5 { EnableSMTPForAllMailboxes }
    6 { DisableSMTPForAllMailboxes }
    7 { EnableSMTPForSpecificMailboxes }
    8 { DisableSMTPForSpecificMailboxes }
    9 { ShowSMTPStatusForSpecificMailboxes }
    10 { SendMail }
    0 { break }
    default { WriteConsole @("Opcion no valida", "red", 2, 0) }
  }
} while ($Option -ne 0)