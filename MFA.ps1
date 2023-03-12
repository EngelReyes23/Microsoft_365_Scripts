Set-ExecutionPolicy unrestricted -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name MSOnline
Connect-MsolService

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

# Sets the MFA requirement state
function Set-MfaState {
  [CmdletBinding()]
  param(
    [Parameter(ValueFromPipelineByPropertyName = $True)]
    $ObjectId,
    [Parameter(ValueFromPipelineByPropertyName = $True)]
    $UserPrincipalName,
    [ValidateSet("Disabled", "Enabled", "Enforced")]
    $State
  )
  Process {
    Write-Verbose ("Setting MFA state for user '{0}' to '{1}'." -f $ObjectId, $State)
    $Requirements = @()
    if ($State -ne "Disabled") {
      $Requirement =
      [Microsoft.Online.Administration.StrongAuthenticationRequirement]::new()
      $Requirement.RelyingParty = "*"
      $Requirement.State = $State
      $Requirements += $Requirement
    }
    Set-MsolUser -ObjectId $ObjectId -UserPrincipalName $UserPrincipalName `
      -StrongAuthenticationRequirements $Requirements
  }
}

#----------------------------------------------------------------------------------------------------#

#region For specific user
# Enable MFA for specific user
function Enable-MfaForUser {
  $UserPrincipalName = Read-Host "Enter user's UPN"
  $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
  if ($null -ne $User) {
    WriteConsole @("User $($User.DisplayName) will have MFA enabled.", "green", 1, 0)
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State Enabled
  }
  else { WriteConsole("User not found.", "red", 1, 0) }
  Pause
}

# Disable MFA for specific user
function Disable-MfaForUser {
  $UserPrincipalName = Read-Host "Enter user's UPN"
  $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
  if ($null -ne $User) {
    WriteConsole @("User $($User.DisplayName) will have MFA disabled.", "green", 1, 0)
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State Disabled
  }
  else { WriteConsole("User not found.", "red", 1, 0) }
  Pause
}

# Enforce MFA for specific user
function Enforce-MfaForUser {
  $UserPrincipalName = Read-Host "Enter user's UPN"
  $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
  if ($null -ne $User) {
    WriteConsole @("User $($User.DisplayName) will have MFA enforced.", "green", 1, 0)
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State Enforced
  }
  else { WriteConsole("User not found.", "red", 1, 0) }
  Pause
}
#endregion

#----------------------------------------------------------------------------------------------------#

#region For All Users

# Enable MFA for all users
function Enable-MfaForAllUsers {
  $Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
  foreach ($User in $Users) {
    WriteConsole @("Enabling MFA for $($User.DisplayName)", "green")
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State Enabled
  }
  Pause
}

# Disable MFA for all users
function Disable-MfaForAllUsers {
  $Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
  foreach ($User in $Users) {
    WriteConsole @("Disabling MFA for $($User.DisplayName)", "green")
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State Disabled
  }
  Pause
}

# Enforce MFA for all users
function Enforce-MfaForAllUsers {
  $Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
  foreach ($User in $Users) {
    WriteConsole @("Enforcing MFA for $($User.DisplayName)", "green")
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State Enforced
  }
  Pause
}
#endregion

#----------------------------------------------------------------------------------------------------#

#region Main

# Show menu of options
do {
  Clear-Host
  WriteConsole("--------------------[ MFA Management ]--------------------", "white", 1, 0)
  WriteConsole("1- Enable MFA for all users.", "white")
  WriteConsole("2- Disable MFA for all users.", "white")
  WriteConsole("3- Enforce MFA for all users.", "white")
  WriteConsole("4- Enable MFA for a specific user.", "white")
  WriteConsole("5- Disable MFA for a specific user.", "white")
  WriteConsole("6- Enforce MFA for a specific user.", "white")
  WriteConsole("7- Exit", "white")
  $option = Read-Host "Select an option"

  switch ($option) {
    1 { Enable-MfaForAllUsers }
    2 { Disable-MfaForAllUsers }
    3 { Enforce-MfaForAllUsers }
    4 { Enable-MfaForUser }
    5 { Disable-MfaForUser }
    6 { Enforce-MfaForUser }
    7 { break }
    default {
      WriteConsole("Invalid option.", "red", 1, 0)
      Pause
    }
  }
} while ($option -ne 7)
#endregion