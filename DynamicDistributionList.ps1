# Connect Exchange
Set-ExecutionPolicy unrestricted -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name ExchangeOnlineManagement
Connect-ExchangeOnline

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

function CheckDynamicDistributionList($DDGIdentity) {
  $DDG = Get-DynamicDistributionGroup -Identity $DDGIdentity -ErrorAction SilentlyContinue
  if (!$DDG) {
    WriteConsole("Dynamic Distribution List does not exist", "red", 2)
    Pause
    return $false
  }
  return $true
}

function GetDynamicDistributionListName() {
  $DDGIdentity = Read-Host "Enter the name of the Dynamic Distribution List"
  return $DDGIdentity
}

function ShowMembers($DDGIdentity) {
  Clear-Host
  WriteConsole("--------------------[ Dynamic Distribution List Members ]--------------------", "green", 2)
  Get-DynamicDistributionGroupMember -Identity $DDGIdentity | Format-Table DisplayName, PrimarySmtpAddress
  WriteConsole("--------------------------------------------------------------------------------", "white")
  Pause
}

function UpdateDynamicDistributionList($DDGIdentity) {
  Set-DynamicDistributionGroup -Identity $DDGIdentity -ForceMembershipRefresh
  WriteConsole("Dynamic Distribution List updated successfully", "green", 2)
  Pause
}

function DeleteDynamicDistributionList($DDGIdentity) {
  Remove-DynamicDistributionGroup -Identity $DDGIdentity
  WriteConsole("Dynamic Distribution List deleted successfully", "green", 2)
  Pause
}

function ShowNotes() {
  Clear-Host
  WriteConsole("Los grupos de distribución dinámicos son grupos de distribución cuya pertenencia se calcula periódicamente en función de propiedades de destinatario específicas que se usan como filtros (filtros predefinidos para filtros personalizados)", "yellow", 2)
  WriteConsole("Puede ejecutar el comando refresh solo después de que haya transcurrido más de una hora desde la última actualización de pertenencia.", "yellow")
  Pause
}

function ShowMenuOptions() {
  Clear-Host
  WriteConsole("--------------------[ Dynamic Distribution List Management ]--------------------", "green", 2)
  WriteConsole @("1- Show Members")
  WriteConsole @("2- Update Dynamic Distribution List")
  WriteConsole @("3- Delete Dynamic Distribution List")
  WriteConsole @("4- Exit")
  $option = Read-Host "Select an option"
  return $option
}

do {
  $option = ShowMenuOptions
  switch ($option) {
    1 {
      $DDGIdentity = GetDynamicDistributionListName
      if (CheckDynamicDistributionList $DDGIdentity) { ShowMembers $DDGIdentity }
    }
    2 {
      ShowNotes
      $DDGIdentity = GetDynamicDistributionListName
      if (CheckDynamicDistributionList $DDGIdentity) { UpdateDynamicDistributionList $DDGIdentity }
    }
    3 {
      $DDGIdentity = GetDynamicDistributionListName
      if (CheckDynamicDistributionList $DDGIdentity) { DeleteDynamicDistributionList $DDGIdentity }
    }
    4 { WriteConsole("Bye", "green", 2) }
  }
} while ($option -ne 4)