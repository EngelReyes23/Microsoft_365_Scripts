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

function IsUserExist($User) {
  $User = Get-Mailbox -Identity $User -ErrorAction SilentlyContinue
  return $null -ne $User
}

function ShowRetentionAction() {
  WriteConsole @("--------------------[ Retention Action ]--------------------", "yellow", 2)
  WriteConsole @("1- Move To Archive")
  WriteConsole @("2- Move To Deleted Items")
  WriteConsole @("3- Delete And Allow Recovery")
  WriteConsole @("4- Delete And Prevent Recovery")
  $RetentionAction = Read-Host "Retention Action"

  switch ($RetentionAction) {
    1 { $RetentionAction = "MoveToArchive" }
    2 { $RetentionAction = "MoveToDeletedItems" }
    3 { $RetentionAction = "DeleteAndAllowRecovery" }
    4 { $RetentionAction = "DeleteAndPreventRecovery" }
  }
  WriteConsole @("------------------------------------------------------------", "white", 2)
  return $RetentionAction
}

function ShowType() {
  WriteConsole @("--------------------[ Type ]--------------------", "yellow", 2)
  WriteConsole @("1- Personal")
  WriteConsole @("2- PublicFolder")
  WriteConsole @("3- Archive")
  $Type = Read-Host "Type"

  switch ($Type) {
    1 { $Type = "Personal" }
    2 { $Type = "PublicFolder" }
    3 { $Type = "Archive" }
  }
  WriteConsole @("------------------------------------------------------------", "white", 2)
  return $Type
}

function CreateRetentionPolicyTag() {
  Clear-Host
  WriteConsole @("--------------------[ Retention Policy Tag ]--------------------", "green", 2)
  $params = @{}
  $params.RetentionPolicyTagName = Read-Host "Name Retention Policy Tag (No spaces)"
  $params.Comment = Read-Host "Comment (Optional - No spaces))"
  $params.Type = ShowType
  $params.AgeLimitForRetention = Read-Host "Age Limit For Retention (Days)"
  $params.RetentionEnabled = Read-Host "Retention Enabled (True = 1, False = 0)"
  $params.RetentionEnabled = [bool]([int]$params.RetentionEnabled)
  $params.RetentionAction = ShowRetentionAction

  New-RetentionPolicyTag $params.RetentionPolicyTagName -Type $params.Type -Comment $params.Comment -RetentionEnabled $params.RetentionEnabled -AgeLimitForRetention $params.AgeLimitForRetention -RetentionAction $params.RetentionAction

  WriteConsole @("Retention Policy Tag Created", "green", 2)
  Pause
}

function ShowRetentionPolicyTag() {
  WriteConsole @("--------------------[ Retention Policy Tag List ]--------------------", "green", 2)
  $RetentionPolicyTag = Get-RetentionPolicyTag
  foreach ($item in $RetentionPolicyTag) {
    WriteConsole @($item.Name, "white", 0)
  }
  WriteConsole @("-------------------------------------------------------------------", "white", 2)
  Pause
}

function ShowRetentionPolicy() {
  WriteConsole @("--------------------[ Retention Policy List ]--------------------", "green", 2)
  $RetentionPolicy = Get-RetentionPolicy
  foreach ($item in $RetentionPolicy) {
    WriteConsole @($item.Name, "white", 0)
  }
  WriteConsole @("-------------------------------------------------------------------", "white", 2)
  Pause
}

function CreateRetentionPolicy() {
  Clear-Host
  WriteConsole @("--------------------[ Retention Policy ]--------------------", "green", 2)
  $params = @{}
  $params.NameRetentionPolicy = Read-Host "Name Retention Policy (No spaces)"
  ShowRetentionPolicyTag
  $params.RetentionPolicyTagLinks = Read-Host "Retention Policy Tag Links"

  New-RetentionPolicy $params.NameRetentionPolicy -RetentionPolicyTagLinks $params.RetentionPolicyTagLinks

  WriteConsole @("Retention Policy Created", "green", 2)
  Pause
}

function DeleteRetentionPolicyTag() {
  Clear-Host
  WriteConsole @("--------------------[ Delete Retention Policy Tag ]--------------------", "green", 2)
  ShowRetentionPolicyTag
  $RetentionPolicyTag = Read-Host "Retention Policy Tag"
  Remove-RetentionPolicyTag $RetentionPolicyTag
  WriteConsole @("Retention Policy Tag Deleted", "green", 2)
  Pause
}

function DeleteRetentionPolicy() {
  Clear-Host
  WriteConsole @("--------------------[ Delete Retention Policy ]--------------------", "green", 2)
  ShowRetentionPolicy
  $RetentionPolicy = Read-Host "Retention Policy"
  Remove-RetentionPolicy $RetentionPolicy
  WriteConsole @("Retention Policy Deleted", "green", 2)
  Pause
}

function SetRetentionPolicyToUser() {
  Clear-Host
  WriteConsole @("--------------------[ Set Retention Policy ]--------------------", "green", 2)
  $User = Read-Host "User"
  if (IsUserExist $User) {
    ShowRetentionPolicy
    $RetentionPolicy = Read-Host "Retention Policy"
    Set-Mailbox $User -RetentionPolicy $RetentionPolicy
    WriteConsole @("Retention Policy Assigned", "green", 2)
    Pause
  }
  else {
    WriteConsole @("User not found", "red", 2)
    Pause
  }
}

function ShowMenu() {
  Clear-Host
  WriteConsole @("--------------------[ Retention Policy Management (MRM) ]--------------------", "green", 2)
  WriteConsole @("1- Retention Policy Tag")
  WriteConsole @("2- Retention Policy")
  WriteConsole @("3- Set Retention Policy to User")
  WriteConsole @("4- Exit")
  $Option = Read-Host "Option"
  return $Option
}

function ShowRetentionPolicyTagMenu() {
  Clear-Host
  WriteConsole @("--------------------[ Retention Policy Tag ]--------------------", "green", 2)
  WriteConsole @("1- Create Retention Policy Tag")
  WriteConsole @("2- Show Retention Policy Tag List")
  WriteConsole @("3- Delete Retention Policy Tag")
  WriteConsole @("4- Exit")
  $Option = Read-Host "Option"
  return $Option
}

function ShowRetentionPolicyMenu() {
  Clear-Host
  WriteConsole @("--------------------[ Retention Policy ]--------------------", "green", 2)
  WriteConsole @("1- Create Retention Policy")
  WriteConsole @("2- Show Retention Policy List")
  WriteConsole @("3- Delete Retention Policy")
  WriteConsole @("4- Exit")
  $Option = Read-Host "Option"
  return $Option
}

# Main
do {
  $Option = ShowMenu
  switch ($Option) {
    1 {
      do {
        $Option1 = ShowRetentionPolicyTagMenu
        switch ($Option1) {
          1 { CreateRetentionPolicyTag }
          2 { ShowRetentionPolicyTag }
          3 { DeleteRetentionPolicyTag }
        }
      } while ($Option1 -ne 4)
    }
    2 {
      do {
        $Option2 = ShowRetentionPolicyMenu
        switch ($Option2) {
          1 { CreateRetentionPolicy }
          2 { ShowRetentionPolicy }
          3 { DeleteRetentionPolicy }
        }
      } while ($Option2 -ne 4)
    }
    3 { SetRetentionPolicyToUser }
  }
} while ($Option -ne 4)