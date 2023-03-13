# Connect to MsolService
Set-ExecutionPolicy unrestricted -Force
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-Module -Name MSOnline
Connect-MsolService

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
# Set MFA state for specific user
function Set-MfaStateForUser {
  param(
    [string]$UserPrincipalName,
    [ValidateSet('Enabled', 'Disabled', 'Enforced')]
    [string]$State
  )
  $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
  if ($null -ne $User) {
    $StateVerb = if ($State -eq 'Enabled') { 'enabling' } elseif ($State -eq 'Disabled') { 'disabling' } elseif ($State -eq 'Enforced') { 'enforcing' }
    WriteConsole -Text "User $($User.DisplayName) will have MFA $StateVerb." -ForegroundColor "green" -NewLine 1
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State $State
  }
  else { WriteConsole -Text "`nUser not found." -ForegroundColor "red" -NewLine 1 }
  Pause
}

# Request the user's email and receive the status of MFA by parameter
function Set-MfaStateForUserPrompt {
  param(
    [ValidateSet('Enabled', 'Disabled', 'Enforced')]
    [string]$State
  )
  $UserPrincipalName = Read-Host "`nEnter user's UPN"
  Set-MfaStateForUser -UserPrincipalName $UserPrincipalName -State $State
}

# Get MFA status for specific user
function Get-MfaStatusForUser {
  $UserPrincipalName = Read-Host "`nEnter user's UPN"
  $User = Get-MsolUser -UserPrincipalName $UserPrincipalName
  if ($null -ne $User) {
    $MfaStatus = Get-MsolUser -UserPrincipalName $UserPrincipalName | Select-Object -ExpandProperty StrongAuthenticationRequirements
    if ($null -eq $MfaStatus) { WriteConsole -Text "`nMFA is disabled for user $($User.DisplayName)." -ForegroundColor "red" -NewLine 1 }
    elseif ($MfaStatus.State -eq "Enabled") { WriteConsole -Text "`nMFA is enabled for user $($User.DisplayName)." -ForegroundColor "green" -NewLine 1 }
    elseif ($MfaStatus.State -eq "Enforced") { WriteConsole -Text "`nMFA is enforced for user $($User.DisplayName)." -ForegroundColor "yellow" -NewLine 1 }
  }
  else { WriteConsole -Text "`nUser not found." -ForegroundColor "red" -NewLine 1 }
  Pause
}
#endregion

#----------------------------------------------------------------------------------------------------#

#region For All Users
# Set MFA state for all users
function Set-MfaForAllUsers {
  param (
    [Parameter(Mandatory = $true)]
    [ValidateSet('Enabled', 'Disabled', 'Enforced')]
    [string]$State
  )
  $Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
  foreach ($User in $Users) {
    WriteConsole -Text "Setting MFA state to '$State' for $($User.DisplayName)" -ForegroundColor "green"
    Set-MfaState -ObjectId $User.ObjectId -UserPrincipalName $User.UserPrincipalName -State $State
  }
  Pause
}

# Get MFA status for all users
function Get-MfaStatusForAllUsers {
  $Users = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
  foreach ($User in $Users) {
    $MfaStatus = Get-MsolUser -UserPrincipalName $User.UserPrincipalName | Select-Object -ExpandProperty StrongAuthenticationRequirements
    if ($null -eq $MfaStatus) { WriteConsole -Text "MFA is disabled for user $($User.DisplayName)." -ForegroundColor "red" }
    elseif ($MfaStatus.State -eq "Enabled") { WriteConsole -Text "MFA is enabled for user $($User.DisplayName)." -ForegroundColor "green" }
    elseif ($MfaStatus.State -eq "Enforced") { WriteConsole -Text "MFA is enforced for user $($User.DisplayName)." -ForegroundColor "yellow" }
  }
  Pause
}
#endregion

#----------------------------------------------------------------------------------------------------#

#region Main
# Show menu of options
do {
  Clear-Host
  WriteConsole -Text "--------------------[ MFA Management ]--------------------" -ForegroundColor "yellow" -NewLine 1
  WriteConsole -Text "1- Enable MFA for all users."
  WriteConsole -Text "2- Disable MFA for all users."
  WriteConsole -Text "3- Enforce MFA for all users."
  WriteConsole -Text "4- Get MFA status for all users."
  WriteConsole -Text "5- Enable MFA for specific user."
  WriteConsole -Text "6- Disable MFA for specific user."
  WriteConsole -Text "7- Enforce MFA for specific user."
  WriteConsole -Text "8- Get MFA status for specific user."
  WriteConsole -Text "9- Exit."
  WriteConsole -Text "----------------------------------------------------------" -ForegroundColor "yellow" -NewLine 1
  $option = Read-Host "Select an option"

  switch ($option) {
    1 { Set-MfaForAllUsers -State "Enabled" }
    2 { Set-MfaForAllUsers -State "Disabled" }
    3 { Set-MfaForAllUsers -State "Enforced" }
    4 { Get-MfaStatusForAllUsers }
    5 { Set-MfaStateForUserPrompt -State "Enabled" }
    6 { Set-MfaStateForUserPrompt -State "Disabled" }
    7 { Set-MfaStateForUserPrompt -State "Enforced" }
    8 { Get-MfaStatusForUser }
    9 { WriteConsole -Text "`nExiting..." -ForegroundColor "yellow" -NewLine 1 -Wait 2 }
    default { WriteConsole -Text "`nInvalid option." -ForegroundColor "red" -NewLine 1 -Wait 3 }
  }
} while ($option -ne 9)
#endregion