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

function CheckDynamicDistributionList($DDGIdentity) {
  $DDG = Get-DynamicDistributionGroup -Identity $DDGIdentity -ErrorAction SilentlyContinue
  if (!$DDG) {
    WriteConsole -Text "`nDynamic Distribution List does not exist." -ForegroundColor "red" -NewLine 2
    Pause
    return $false
  }
  return $true
}

function GetDynamicDistributionListName() {
  return Read-Host "`nEnter the name of the Dynamic Distribution List."
}

function ShowMembers() {
  $DDGIdentity = GetDynamicDistributionListName
  if (CheckDynamicDistributionList $DDGIdentity) {
    WriteConsole -Text "--------------------[ Dynamic Distribution List Members ]--------------------" -ForegroundColor "green" -NewLine 2
    Get-DynamicDistributionGroupMember -Identity $DDGIdentity | Format-Table DisplayName, PrimarySmtpAddress
    WriteConsole -text "-----------------------------------------------------------------------------" -NewLine 2
    Pause
  }
}

function UpdateDynamicDistributionList() {
  ShowNotes
  $DDGIdentity = GetDynamicDistributionListName
  if (CheckDynamicDistributionList $DDGIdentity) {
    Set-DynamicDistributionGroup -Identity $DDGIdentity -ForceMembershipRefresh
    WriteConsole -Text "`nDynamic Distribution List updated successfully." -ForegroundColor "green" -NewLine 2
    Pause
  }
}

function DeleteDynamicDistributionList($DDGIdentity) {
  $DDGIdentity = GetDynamicDistributionListName
  if (CheckDynamicDistributionList $DDGIdentity) {
    Remove-DynamicDistributionGroup -Identity $DDGIdentity
    WriteConsole -Text "`nDynamic Distribution List deleted successfully." -ForegroundColor "green" -NewLine 2
    Pause
  }
}

function ShowNotes() {
  WriteConsole -Text "`nLos grupos de distribución dinámicos son grupos de distribución cuya pertenencia se calcula periódicamente en función de propiedades de destinatario específicas que se usan como filtros (filtros predefinidos para filtros personalizados)" -ForegroundColor "yellow" -NewLine 2
  WriteConsole -Text "Puede ejecutar el comando refresh solo después de que haya transcurrido más de una hora desde la última actualización de pertenencia." -ForegroundColor "yellow" -NewLine 2
  Pause
}

function ShowMenuOptions() {
  Clear-Host
  WriteConsole -Text "--------------------[ Dynamic Distribution List Management ]--------------------" -ForegroundColor "Yellow" -NewLine 1
  WriteConsole -Text "1- Show Members."
  WriteConsole -Text "2- Update Dynamic Distribution List."
  WriteConsole -Text "3- Delete Dynamic Distribution List."
  WriteConsole -Text "4- Exit"
  WriteConsole -Text "--------------------------------------------------------------------------------" -ForegroundColor "Yellow"
  return Read-Host "Select an option"
}

do {
  $option = ShowMenuOptions
  switch ($option) {
    1 { ShowMembers }
    2 { UpdateDynamicDistributionList }
    3 { DeleteDynamicDistributionList }
    4 { WriteConsole -Text "`nExiting..." -ForegroundColor "Yellow" -NewLine 2 -Wait 2 }
    default { WriteConsole -Text "`nInvalid option" -ForegroundColor "red" -NewLine 2 -Wait 3 }
  }
} while ($option -ne 4)