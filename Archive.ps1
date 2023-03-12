<#
Script para el manejo de archivado en buzones de Exchange Online (Microsoft 365)
Autor: Engel Reyes <engel.reyes@concentrix.com>
#>

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

function IsArchiveEnabled($mailbox) {
  $user = Get-Mailbox -Identity $mailbox -ErrorAction SilentlyContinue
  return $user.ArchiveStatus -eq "Active"
}

function IsAutoExpandArchiveEnabled($mailbox) {
  $user = Get-Mailbox -Identity $mailbox -ErrorAction SilentlyContinue
  return $user.AutoExpandingArchiveStatus -eq "Active"
}

# Activar el archive para un usuario ( recibe el parametro $mailbox )
function EnableArchive($mailbox) {

  # Si el usuario existe, se activa el archive
  if (IsUserExist($mailbox)) {

    # Si el archive no esta activado, se activa
    if (!(IsArchiveEnabled($mailbox))) {
      Enable-Mailbox -Identity $mailbox -Archive
      WriteConsole("Archive enabled for $mailbox", "Green", 2)
      Pause
    }

    # Si el archive ya esta activado, se muestra un mensaje de error
    else {
      WriteConsole("Archive already enabled for $mailbox", "Red", 2)
      Pause
    }
  }

  # Si el usuario no existe, se muestra un mensaje de error
  else {
    WriteConsole("User $mailbox does not exist", "Red", 2)
    Pause
  }
}

function ShowAutoExpandArchiveInfo() {
  WriteConsole("Después de activar el archivado de expansión automática para su organización o para un usuario específico, no se puede desactivar. Los administradores tampoco pueden ajustar la cuota de almacenamiento para el archivado de expansión automática.
`nEl archivado de expansión automática impide recuperar o restaurar un buzón inactivo. Esto significa que si habilita el archivado de expansión automática para un buzón de correo y el buzón se deja inactivo en una fecha posterior, no podrá recuperar el buzón inactivo (convertirlo en un buzón activo) ni restaurarlo (combinando el contenido con un buzón existente).
`nSi el archivado de expansión automática está habilitado en un buzón inactivo, la única manera de recuperar datos es mediante la herramienta de búsqueda de contenido de la portal de cumplimiento Microsoft Purview para exportar los datos del buzón e importarlos a otro buzón", "Yellow", 2)
}

function EnableAutoExpandArchive($mailbox) {

  if (IsUserExist($mailbox)) {

    # verifica si tiene archive activado
    if (!(IsArchiveEnabled($mailbox))) {
      WriteConsole("Archive is not enabled for $mailbox", "Red", 2)
      Pause
      return
    }

    if (!(IsAutoExpandArchiveEnabled($mailbox))) {
      Enable-Mailbox $mailbox -AutoExpandingArchive
      WriteConsole("Auto expanding archive enabled for $mailbox", "Green", 2)
      Pause
    }
    else {
      WriteConsole("Auto expanding archive already enabled for $mailbox", "Red", 2)
      Pause
    }
  }
  else {
    WriteConsole("User $mailbox does not exist", "Red", 2)
    Pause
  }
}

# activa auto expandir el archive para todos los usuarios
function EnableAutoExpandArchiveAll() {
  # confirma si se desea activar el auto expandir el archive para todos los usuarios
  $confirm = Read-Host "Are you sure you want to enable auto expanding archive for all users? (y/n)"
  if ($confirm -eq "y") {
    Set-OrganizationConfig -AutoExpandingArchive
    WriteConsole("Auto expanding archive enabled for all users", "Green", 2)
    Pause
  }
}

# Fuerza el movimiento de los elementos de correo electrónico de un buzón al buzon de archivado
function ForceMoveToArchive($mailbox) {
  if (IsUserExist($mailbox)) {
    if (IsArchiveEnabled($mailbox)) {
      Start-ManagedFolderAssistant -Identity $mailbox -Fullcrawl
      WriteConsole("Forced move to archive for $mailbox", "Green", 2)
      Pause
    }
    else {
      WriteConsole("Archive is not enabled for $mailbox", "Red", 2)
      Pause
    }
  }
  else {
    WriteConsole("User $mailbox does not exist", "Red", 2)
    Pause
  }
}

function DisableArchive($mailbox) {

  # Si el usuario existe, se desactiva el archive
  if (IsUserExist($mailbox)) {

    # Si el archive esta activado, se desactiva
    if (IsArchiveEnabled($mailbox)) {
      Disable-Mailbox -Identity $mailbox -Archive
      WriteConsole("Archive disabled for $mailbox", "Green", 2)
      Pause
    }

    # Si el archive no esta activado, se muestra un mensaje de error
    else {
      WriteConsole("Archive already disabled for $mailbox", "Red", 2)
      Pause
    }
  }
  # Si el usuario no existe, se muestra un mensaje de error
  else {
    WriteConsole("User $mailbox does not exist", "Red", 2)
    Pause
  }
}

# Activar el archive para todos los usuarios
function EnableArchiveAll() {
  # confirma si se desea activar el archive para todos los usuarios
  $confirm = Read-Host "Are you sure you want to enable archive for all users? (y/n)"
  if ($confirm -eq "y") {
    # activa el archive para todos los usuarios, filtrando si no tienen el archive activado
    Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ArchiveStatus -eq "None" } | Enable-Mailbox -Archive
    WriteConsole("Archive enabled for all users", "Green", 2)
    Pause
  }
}

# Desactivar el archive para todos los usuarios
function DisableArchiveAll() {
  # confirma si se desea desactivar el archive para todos los usuarios
  $confirm = Read-Host "Are you sure you want to disable archive for all users? (y/n)"
  if ($confirm -eq "y") {
    # desactiva el archive para todos los usuarios, filtrando si tienen el archive activado
    Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ArchiveStatus -eq "Active" } | Disable-Mailbox -Archive
    WriteConsole("Archive disabled for all users", "Green", 3)
    Pause
  }
}

# Lista todos los usuarios con el archive activado
function ListArchiveEnabled() {
  Clear-Host
  WriteConsole("--------------------[ Listing all users with archive enabled ]--------------------", "Green", 2)
  # lista todos los usuarios con el archive activado
  Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ArchiveStatus -eq "Active" } | Format-List DisplayName, UserPrincipalName, MicrosoftOnlineServicesID, ArchiveStatus, AutoExpandingArchiveEnabled, WhenCreated, GUID
  WriteConsole("--------------------------------------------------------------------------------", "white", 2)
  Pause
}

# Lista todos los usuarios con el archive desactivado
function ListArchiveDisabled() {
  Clear-Host
  WriteConsole("--------------------[ Listing all users with archive disabled ]--------------------", "Green", 2)
  # lista todos los usuarios con el archive desactivado
  Get-Mailbox -ResultSize Unlimited | Where-Object { $_.ArchiveStatus -eq "None" } | Format-List DisplayName, UserPrincipalName, MicrosoftOnlineServicesID, ArchiveStatus, AutoExpandingArchiveEnabled, WhenCreated, GUID
  WriteConsole("--------------------------------------------------------------------------------", "white", 2)
  Pause
}

# Listar el estado del archive para todos los usuarios
function ListArchiveStatus() {
  Clear-Host
  WriteConsole("--------------------[ Listing all users with archive status ]--------------------", "Green", 2)
  # lista el estado del archive para todos los usuarios
  Get-Mailbox -ResultSize Unlimited | Format-List DisplayName, UserPrincipalName, MicrosoftOnlineServicesID, ArchiveStatus, AutoExpandingArchiveEnabled, WhenCreated, GUID
  WriteConsole("--------------------------------------------------------------------------------", "white", 2)
  Pause
}

# Show main menu
function ShowMenu() {
  Clear-Host
  WriteConsole("--------------------[ Exchange Online Archive Management ]--------------------", "Green", 2)
  Write-Host "1- Enable Archive `n2- Disable Archive `n3- List Archive Status `n4- Enable Auto Expand Archive `n5- Force Move to Archive `n6- Exit"
}

# Show sub menu Enable Archive
function ShowSubMenuEnableArchive() {
  Clear-Host
  WriteConsole("--------------------[ Enable Archive ]--------------------", "Green", 2)
  Write-Host "1- Enable Archive for a user `n2- Enable Archive for all users `n3- Back"
}

# Show sub menu Disable Archive
function ShowSubMenuDisableArchive() {
  Clear-Host
  WriteConsole("--------------------[ Disable Archive ]--------------------", "Green", 2)
  Write-Host "1- Disable Archive for a user `n2- Disable Archive for all users `n3- Back"
}

# Show sub menu List Archive Status
function ShowSubMenuListArchiveStatus() {
  Clear-Host
  WriteConsole("--------------------[ List Archive Status ]--------------------", "Green", 2)
  Write-Host "1- List Archive Enabled `n2- List Archive Disabled `n3- List Archive Status for all users `n4- Back"
}

# Show sub menu Enable Auto Expand Archive
function ShowSubMenuEnableAutoExpandArchive() {
  Clear-Host
  WriteConsole("--------------------[ Enable Auto Expand Archive ]--------------------", "Green", 2)
  Write-Host "1- Enable Auto Expand Archive for a user `n2- Enable Auto Expand Archive for all users `n3- Back"
}

# Show sub menu Force Move to Archive
function ShowSubMenuForceMoveToArchive() {
  Clear-Host
  WriteConsole("--------------------[ Force Move to Archive ]--------------------", "Green", 2)
  Write-Host "1- Force Move to Archive for a user `n2- Back"
}
#endregion -------------------------------------------------------------------------------- #

#region ---------------------------------------- Main ---------------------------------------- #

do {
  ShowMenu
  $option = Read-Host "Option"
  switch ($option) {
    1 {
      do {
        ShowSubMenuEnableArchive
        $option = Read-Host "Option"
        switch ($option) {
          1 {
            $mailbox = Read-Host "Mailbox address"
            EnableArchive($mailbox)
          }
          2 { EnableArchiveAll }
          3 { WriteConsole("Back to main menu", "Green", 2, 1) }
          default {
            WriteConsole("Invalid option", "Red", 2, 2)
            Clear-Host
          }
        }
      } while ($option -ne 3)
    }

    2 {
      do {
        ShowSubMenuDisableArchive
        $option = Read-Host "Option"
        switch ($option) {
          1 {
            $mailbox = Read-Host "Mailbox address"
            DisableArchive($mailbox)
          }
          2 { DisableArchiveAll }
          3 { WriteConsole("Back to main menu", "Green", 2, 1) }
          default { WriteConsole("Invalid option", "Red", 2, 2) }
        }
      } while ($option -ne 3)
    }

    3 {
      do {
        ShowSubMenuListArchiveStatus
        $option = Read-Host "Option"
        switch ($option) {
          1 { ListArchiveEnabled }
          2 { ListArchiveDisabled }
          3 { ListArchiveStatus }
          4 { WriteConsole("Back to main menu", "Green", 2, 1) }
          default { WriteConsole("Invalid option", "Red", 2, 2) }
        }
      } while ($option -ne 4)
    }

    4 {
      ShowAutoExpandArchiveInfo

      # confirma si se desea continuar
      $confirm = Read-Host "Do you want to continue? (y/n)"

      if ($confirm -eq "y") {
        do {
          ShowSubMenuEnableAutoExpandArchive
          $option = Read-Host "Option"
          switch ($option) {
            1 {
              $mailbox = Read-Host "Mailbox address"
              EnableAutoExpandArchive($mailbox)
            }
            2 { EnableAutoExpandArchiveAll }
            3 { WriteConsole("Back to main menu", "Green", 2, 1) }
            default { WriteConsole("Invalid option", "Red", 2, 2) }
          }
        } while ($option -ne 3)
      }
      else { WriteConsole("Back to main menu", "Green", 2) }
    }

    5 {
      do {
        ShowSubMenuForceMoveToArchive
        $option = Read-Host "Option"
        switch ($option) {
          1 {
            $mailbox = Read-Host "Mailbox address"
            ForceMoveToArchive($mailbox)
          }
          2 { WriteConsole("Back to main menu", "Green", 2, 1) }
          default { WriteConsole("Invalid option", "Red", 2, 2) }
        }
      } while ($option -ne 2)
    }

    6 { WriteConsole("Bye Bye", "Green", 2) }
    default {
      WriteConsole("Invalid option", "Red", 2, 2)
    }
  }
} while ($option -ne 6)
#endregion -------------------------------------------------------------------------------- #