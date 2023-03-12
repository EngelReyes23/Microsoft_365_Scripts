# Connect Exchange
Set-ExecutionPolicy unrestricted -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline

#region ---------------------------------------- Functions ---------------------------------------- #
# Escribe en la consola, recibe 3 parámetros (Texto, ForegroundColor, numero de saltos de linea, segundos de espera)
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

# Valida si el usuario existe
function IsUserExist($User) {
  $User = Get-Mailbox -Identity $User -ErrorAction SilentlyContinue
  return $null -ne $User
}

# Valida si el buzón está inactivo
function IsInactiveMailbox($User) {
  $User = Get-Mailbox -InactiveMailboxOnly -Identity $User -ErrorAction SilentlyContinue
  return $null -ne $User
}


# Muestra los buzones inactivos
function ListInactiveMailboxes() {
  Clear-Host
  WriteConsole("--------------------[ List Inactive Mailboxes ]--------------------", "Green")
  Get-Mailbox -InactiveMailboxOnly | Format-List DisplayName, MicrosoftOnlineServicesID, PrimarySMTPAddress, WhenMailboxCreated, WhenSoftDeleted
  WriteConsole("--------------------------------------------------------------------------------", "white", 2)
  Pause
}

# Muestra los buzones activos
function ListActiveMailboxes() {
  Clear-Host
  WriteConsole("--------------------[ List Active Mailboxes ]--------------------", "Green")
  Get-Mailbox | Format-List DisplayName, MicrosoftOnlineServicesID, PrimarySMTPAddress, WhenMailboxCreated
  WriteConsole("--------------------------------------------------------------------------------", "white", 2)
  Pause
}

# Solicita la dirección del buzon inactivo
function GetInactiveMailboxAddress() {
  $InactiveMailboxAddress = read-host "Ingrese la dirección del buzón inactivo"

  # Si el buzon inactivo no existe, solicita nuevamente la dirección del buzon inactivo
  while (!(IsInactiveMailbox $InactiveMailboxAddress)) {
    WriteConsole("The Inactive Mailbox does not exist", "Red", 2)
    Pause
    ListInactiveMailboxes | Out-Host
    $InactiveMailboxAddress = read-host "Ingrese la dirección del buzón inactivo"
  }

  return $InactiveMailboxAddress
}


# Solicita la dirección del buzón activo
function GetActiveMailboxAddress() {
  $ActiveMailboxAddress = read-host "Ingrese la dirección del buzón destino"

  while (!(IsUserExist $ActiveMailboxAddress)) {
    WriteConsole("The Active Mailbox does not exist", "Red", 2)
    Pause
    $ActiveMailboxAddress = read-host "Ingrese la dirección del buzón destino"
  }

  return $ActiveMailboxAddress
}

# Solicita y valida la contraseña para el nuevo buzón (Recover Inactive Mailbox)
function GetPassword() {
  $Password = read-host "Ingrese la contraseña del nuevo buzón (Dejar en blanco para establecer la contraseña por defecto)"

  if ($Password -eq "") {
    $Password = "P@ssw0rd"
    WriteConsole("The Password is: $Password", "Green", 2)
  }
  else {
    $PasswordConfirm = read-host "Confirme la contraseña del nuevo buzón"

    while ($Password -ne $PasswordConfirm) {
      WriteConsole("The Password and Password Confirm does not match", "Red", 2)

      $Password = read-host "Ingrese la contraseña del nuevo buzón"

      while ($Password -eq "") {
        WriteConsole("The Password is required", "Red", 2)
        $Password = read-host "Ingrese la contraseña del nuevo buzón"
      }

      $PasswordConfirm = read-host "Confirme la contraseña del nuevo buzón"
    }
  }

  $ResetPasswordOnNextLogon = read-host "¿Desea que el usuario cambie la contraseña al iniciar sesión por primera vez? (S/N)"

  $ResetPasswordOnNextLogon = $ResetPasswordOnNextLogon -eq "S" -or $ResetPasswordOnNextLogon -eq "s"

  return [PSCustomObject]@{
    Password                 = $Password
    ResetPasswordOnNextLogon = $ResetPasswordOnNextLogon
  }
}

# Solicita la información para crear un nuevo buzón
function GetNewMailboxData() {
  $FirstName = read-host "Ingrese el primer nombre del nuevo buzón"

  while ($FirstName -eq "") {
    WriteConsole("The First Name is required", "Red", 3)
    $FirstName = read-host "Ingrese el primer nombre del nuevo buzón"
  }

  $LastName = read-host "Ingrese el apellido del nuevo buzón"

  while ($LastName -eq "") {
    WriteConsole("The Last Name is required", "Red", 3)
    $LastName = read-host "Ingrese el apellido del nuevo buzón"
  }

  $DisplayName = $FirstName + " " + $LastName

  $Name = $FirstName + $LastName

  WriteConsole("The Display Name is: $DisplayName", "yellow")
  WriteConsole("Luego puede ser modificado en el centro de administración de Office 365", "yellow")

  $MicrosoftOnlineServicesID = read-host "Ingrese el UPN del nuevo buzón"

  # Si el UPN no es válido, solicita nuevamente el UPN
  while ($MicrosoftOnlineServicesID -notmatch "^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$") {
    WriteConsole("The Microsoft Online Services ID (UPN) is not valid", "Red")
    Pause
    $MicrosoftOnlineServicesID = read-host "Ingrese el UPN del nuevo buzón"
  }

  # Si el UPN ya existe, solicita nuevamente el UPN
  while ($null -ne (Get-Mailbox -Identity $MicrosoftOnlineServicesID -ErrorAction SilentlyContinue)) {
    WriteConsole("The Microsoft Online Services ID (UPN) already exists", "Red")
    Pause
    $MicrosoftOnlineServicesID = read-host "Ingrese el UPN del nuevo buzón"
  }

  # Solicita la contraseña para el nuevo buzón
  $PasswordData = GetPassword

  return [PSCustomObject]@{
    Name                      = $Name
    LastName                  = $LastName
    FirstName                 = $FirstName
    DisplayName               = $DisplayName
    Password                  = $PasswordData.Password
    MicrosoftOnlineServicesID = $MicrosoftOnlineServicesID
    ResetPasswordOnNextLogon  = $PasswordData.ResetPasswordOnNextLogon
  }
}

function RestoreInactiveMailbox() {
  ListInactiveMailboxes
  WriteConsole("--------------------[ Restore Inactive Mailbox ]--------------------", "Green", 2)

  $InactiveMailboxAddress = GetInactiveMailboxAddress
  $ActiveMailboxAddress = GetActiveMailboxAddress
  $inactiveMailbox = Get-Mailbox -InactiveMailboxOnly -Identity $InactiveMailboxAddress

  # Agregue LegacyExchangeDN del buzón inactivo como una dirección de proxy X500 al buzón de destino
  $LegacyExchangeDN = $inactiveMailbox.LegacyExchangeDN
  $proxy = "X500:$LegacyExchangeDN"
  Set-Mailbox $ActiveMailboxAddress -EmailAddresses @{Add = $proxy }

  $RestoreFolder = read-host "¿Desea restaurar el contenido dentro de una carpeta? (S/N)"
  # Convertir la respuesta a booleano
  $RestoreFolder = $RestoreFolder -eq "S" -or $RestoreFolder -eq "s"

  if ($RestoreFolder) {
    $FolderName = read-host "Ingrese el nombre de la carpeta"

    while ($FolderName -eq "") {
      WriteConsole("The Folder Name is required", "Red", 2)
      $FolderName = read-host "Ingrese el nombre de la carpeta"
    }

    New-MailboxRestoreRequest -SourceMailbox $inactiveMailbox.DistinguishedName -TargetMailbox $ActiveMailboxAddress -TargetRootFolder $FolderName -AllowLegacyDNMismatch | Out-Host
  }
  else { New-MailboxRestoreRequest -SourceMailbox $inactiveMailbox.DistinguishedName -TargetMailbox $ActiveMailboxAddress -AllowLegacyDNMismatch | Out-Host }

  # Obtener la solicitud de restauración
  $RestoreRequest = Get-MailboxRestoreRequest -TargetMailbox $ActiveMailboxAddress

  # Espera a que la solicitud de restauración se complete
  while ($RestoreRequest.Status -ne "Completed") {
    WriteConsole("Waiting for the Restore Request to complete", "Yellow", 0, 5)
    $RestoreRequest = Get-MailboxRestoreRequest -TargetMailbox $ActiveMailboxAddress

    if ($RestoreRequest.Status -eq "Failed") {
      WriteConsole("The Restore Request has failed", "Red", 2)
      Pause
      break
    }
  }

  <# Una vez completada la solicitud de restauración, puede quitar opcionalmente legacyExchangeDN del buzón inactivo del buzón de destino.
  Salir de LegacyExchangeDN desde el buzón inactivo no afectará al buzón de destino. #>
  Set-Mailbox $ActiveMailboxAddress -EmailAddresses @{remove = $proxy }

  WriteConsole("The Inactive Mailbox has been restored", "Green", 2)
  Pause
}

function RecoverInactiveMailbox() {
  ListInactiveMailboxes
  WriteConsole("--------------------[ Recover Inactive Mailbox ]--------------------", "Green", 2)

  $InactiveMailboxAddress = GetInactiveMailboxAddress
  $InactiveMailboxData = Get-Mailbox -InactiveMailboxOnly -Identity $InactiveMailboxAddress

  # Solicita la información para crear un nuevo buzón
  $NewMailboxData = GetNewMailboxData

  WriteConsole("`nNew Mailbox Data", "Green")
  WriteConsole("------------------------------------------------------------", "white")
  $NewMailboxData | Out-Host
  WriteConsole("------------------------------------------------------------", "white", 2)

  # Crea el nuevo buzón con la información ingresada y la información del buzón inactivo
  New-Mailbox -InactiveMailbox $InactiveMailboxData.DistinguishedName -Name $NewMailboxData.Name -FirstName $NewMailboxData.FirstName -LastName $NewMailboxData.LastName -DisplayName $NewMailboxData.DisplayName -MicrosoftOnlineServicesID $NewMailboxData.MicrosoftOnlineServicesID -Password (ConvertTo-SecureString -String $NewMailboxData.Password -AsPlainText -Force) -ResetPasswordOnNextLogon $NewMailboxData.ResetPasswordOnNextLogon

  WriteConsole("The Inactive Mailbox has been recovered", "Green", 2)
  Pause
}

function mainMenu() {
  Clear-Host
  WriteConsole("--------------------[ Exchange Online Inactive Mailbox Management ]--------------------", "Green", 2)
  WriteConsole("1- Show Inactive Mailboxes`n2- Show Active Mailboxes`n3- Restore Inactive Mailbox`n4- Recover Inactive Mailbox`n5- Identify retentions in inactive mailboxes`n6- Eliminate inactive mailboxes`n7- Exit", "White", 2)
}

#TODO: Terminar de implementar todas las funciones de aqui hacia abajo
# Menu para identificar las retenciones en los buzones inactivos
function IdentifyRetentionsMenu() {
  Clear-Host
  WriteConsole("--------------------[ Identify Retentions ]--------------------", "Green", 2)
  WriteConsole("1- Identify retentions for a specific inactive mailbox`n2- Identify retentions for all inactive mailboxes`n3- Exit", "White", 2)
}

# Identifica las retenciones para un buzon inactivo especifico
function GetRetentionPoliciesForInactiveMailbox($InactiveMailboxAddress) {
  Clear-Host
  WriteConsole("----------[ Retenciones para el buzón inactivo $InactiveMailboxAddress ]----------", "Green", 2)
  Get-Mailbox -Identity $InactiveMailboxAddress -InactiveMailboxOnly | Format-List DisplayName, Name, DistinguishedName, ExchangeGuid, IsInactiveMailbox, LitigationHoldEnabled, LitigationHoldDuration, LitigationHoldDate, LitigationHoldOwner, InPlaceHolds, ComplianceTagHoldApplied
  WriteConsole("----------------------------------------------------------------------------------------------------", "white", 2)
  Pause
}

# Identifica las retenciones para todos los buzones inactivos
function GetRetentionPoliciesForAllInactiveMailboxes() {
  Clear-Host
  WriteConsole("-------------------[ Retenciones para todos los buzones inactivos ]--------------------", "Green", 2)
  Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Select-Object DisplayName, Name, DistinguishedName, ExchangeGuid, IsInactiveMailbox, LitigationHoldEnabled, LitigationHoldDuration, LitigationHoldDate, LitigationHoldOwner, InPlaceHolds, ComplianceTagHoldApplied
  WriteConsole("----------------------------------------------------------------------------------------------------", "white", 2)
  Pause
}

# function para la opcion de ver retencion de un buzon inactivo especifico
function IdentifyRetentionsForInactiveMailbox() {
  ListInactiveMailboxes
  WriteConsole("--------------------[ Identify Retentions ]--------------------", "Green", 2)
  $InactiveMailboxAddress = GetInactiveMailboxAddress
  GetRetentionPoliciesForInactiveMailbox $InactiveMailboxAddress
}

# function para la opcion de ver retencion de todos los buzones inactivos
function IdentifyRetentionsForAllInactiveMailboxes() {
  WriteConsole("--------------------[ Identify Retentions ]--------------------", "Green", 2)
  GetRetentionPoliciesForAllInactiveMailboxes
}

# Menu para eliminar los buzones inactivos
function EliminateInactiveMailboxesMenu() {
  Clear-Host
  WriteConsole("--------------------[ Eliminate Inactive Mailboxes ]--------------------", "Green", 2)
  WriteConsole("1- Eliminate inactive mailbox for a specific inactive mailbox`n2- Eliminate inactive mailbox for all inactive mailboxes`n3- Exit", "White", 2)
}

# Elimina el buzón inactivo especificado
function EliminateInactiveMailbox($InactiveMailboxAddress) {
  Clear-Host
  WriteConsole("----------[ Eliminando el buzón inactivo $InactiveMailboxAddress ]----------", "Green", 2)
  Remove-Mailbox -Identity $InactiveMailboxAddress -Confirm:$false -Force
  WriteConsole("----------------------------------------------------------------------------------------------------", "white", 2)
  Pause
}

# Elimina todos los buzones inactivos
function EliminateAllInactiveMailboxes() {
  Clear-Host
  WriteConsole("-------------------[ Eliminando todos los buzones inactivos ]--------------------", "Green", 2)
  Get-Mailbox -InactiveMailboxOnly -ResultSize Unlimited | Remove-Mailbox -Confirm:$false -Force
  WriteConsole("----------------------------------------------------------------------------------------------------", "white", 2)
  Pause
}
#endregion -------------------------------------------------------------------------------- #


#region ---------------------------------------- Main ---------------------------------------- #
do {
  mainMenu
  $option = read-host "Select an option"
  switch ($option) {
    1 { ListInactiveMailboxes }
    2 { ListActiveMailboxes }
    3 { RestoreInactiveMailbox }
    4 { RecoverInactiveMailbox }
    5 {
      do {
        IdentifyRetentionsMenu
        $option = read-host "Select an option"
        switch ($option) {
          1 { IdentifyRetentionsForInactiveMailbox }
          2 { IdentifyRetentionsForAllInactiveMailboxes }
          3 { WriteConsole("Back to main menu", "Green", 2, 1) }
          default { WriteConsole("Invalid option", "Red", 2, 2) }
        }
      } while ($option -ne 3)
    }
    # 6 {
    #   EliminateInactiveMailboxesMenu
    #   $option = read-host "Select an option"
    #   do {
    #     switch ($option) {
    #       1 { EliminateInactiveMailbox }
    #       2 { EliminateAllInactiveMailboxes }
    #       3 { WriteConsole("Back to main menu", "Green", 2, 1) }
    #       default { WriteConsole("Invalid option", "Red", 2, 2) }
    #     }
    #   } while ($option -ne 3)
    # }
    7 {
      WriteConsole("Bye Bye", "Green", 2, 1)
      exit
    }
    default { WriteConsole("Invalid option", "Red", 2, 2) }
  }
} while ($option -ne 5)
#endregion -------------------------------------------------------------------------------- #