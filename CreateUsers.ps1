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

# Función para crear el archivo CSV
function Create-UserCSV {
  $CSVFile = Read-Host "`nEnter the path and name of the CSV file to create"
  # Crea un objeto con las propiedades que quieres que tenga el archivo CSV
  $CSVObject = New-Object PSObject -Property @{
    "Name"              = ""
    "LastName"          = ""
    "DisplayName"       = ""
    "UserPrincipalName" = ""
    "Password"          = ""
    "LicenseAssignment" = ""
  }
  # Crea el archivo CSV y agrega el objeto
  $CSVObject | Export-Csv -Path "$CSVFile.csv" -NoTypeInformation
  WriteConsole -Text "`nThe CSV file has been created successfully" -ForegroundColor "Green" -NewLine 2
  Pause
}

# Función para leer el archivo CSV y crear los usuarios
function Import-UserCSV {
  $CSVFile = Read-Host "`nEnter the path and name of the CSV file to import"
  # Lee el archivo CSV y crea un usuario para cada fila
  $Users = Import-Csv -Path $CSVFile
  foreach ($User in $Users) {

    # verifica que el usuario no exista
    $UserExists = Get-MsolUser -UserPrincipalName $User.UserPrincipalName -ErrorAction SilentlyContinue
    if ($UserExists) {
      WriteConsole -Text "`nThe user $User.UserPrincipalName already exists" -ForegroundColor "Red" -NewLine 2
      continue
    }

    # Crea el usuario con los valores del archivo CSV
    New-MsolUser -DisplayName $User.DisplayName -UserPrincipalName $User.UserPrincipalName -Password $User.Password -FirstName $User.Name -LastName $User.LastName

    # Asigna la licencia especificada en el archivo CSV
    if ($User.LicenseAssignment) { Set-MsolUserLicense -UserPrincipalName $User.UserPrincipalName -AddLicenses $User.LicenseAssignment }

    WriteConsole -Text "`nThe user $User.UserPrincipalName has been created successfully" -ForegroundColor "Green" -NewLine 2
  }
  Pause
}

# Menú de opciones
do {
  Clear-Host
  WriteConsole -Text "-------------------[ Azure AD Bulk User Creation ]-------------------" -ForegroundColor "Yellow" -NewLine 1
  WriteConsole -Text "1- Create a CSV file for users."
  WriteConsole -Text "2- Import a CSV file and create users."
  WriteConsole -Text "3- Exit."
  WriteConsole -Text "---------------------------------------------------------------------" -ForegroundColor "Yellow"
  $Option = Read-Host "Enter an option"

  switch ($Option) {
    "1" { Create-UserCSV }
    "2" { Import-UserCSV }
    "3" { WriteConsole -Text "`nExiting..." -ForegroundColor "Yellow" -NewLine 2 -Wait 2 }
    default { WriteConsole -Text "`nInvalid option, please try again" -ForegroundColor "Red" -Wait 3 }
  }
} while ($Option -ne "3")