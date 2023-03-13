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

function ShowMenu {
  while ($true) {
    Clear-Host
    WriteConsole -Text "--------------------[ Change Office Product Key ]--------------------" -ForegroundColor "Yellow" -NewLine 1
    WriteConsole -Text "1- Office 2019 or 2016: 32-bit Office and 32-bit Windows."
    WriteConsole -Text "2- Office 2019 or 2016: 32-bit Office and 64-bit Windows."
    WriteConsole -Text "3- Office 2019 or 2016: 64-bit Office and 64-bit Windows."
    WriteConsole -Text "4- Exit"
    WriteConsole -Text "---------------------------------------------------------------------" -ForegroundColor "Yellow"
    return Read-Host "Enter your choice"
  }
}

function ShowSubMenu {
  while ($true) {
    WriteConsole -Text "1- Show Office Product Key."
    WriteConsole -Text "2- Remove Office Product Key."
    WriteConsole -Text "3- Change Office Product Key."
    WriteConsole -Text "4- Exit"
    WriteConsole -Text "-------------------------------------------------------------------------------------------------" -ForegroundColor "Yellow"
    return Read-Host "Enter your choice"
  }
}

function GetOfficeProductKey {
  return Read-Host "Enter the Office Product Key"
}

<#
  Office 2019 or 2016: 32-bit Office and 32-bit Windows

    cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus

    cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /unpkey:XXXXX

    cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
  #>
function Option1 {
  do {
    clear-host
    WriteConsole -Text "--------------------[ Office 2019 or 2016: 32-bit Office and 32-bit Windows ]--------------------" -ForegroundColor "Yellow" -NewLine 1
    $option = ShowSubMenu

    switch ($option) {
      "1" {
        cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus
        Pause
      }

      "2" {
        $OfficeProductKey = Read-Host "Enter the last 5 digits of the Office Product Key"
        cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /unpkey:$OfficeProductKey
        Pause
      }

      "3" {
        $OfficeProductKey = GetOfficeProductKey
        cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /inpkey:$OfficeProductKey
        Pause
      }

      "4" { WriteConsole -Text "`nExiting..." -ForegroundColor "Yellow" -NewLine 2 -Wait 1 }

      Default { WriteConsole -Text "`nInvalid option" -ForegroundColor "Red" -NewLine 2 -Wait 3 }
    }
  } while ($option -ne "4")
}

<#
  Office 2019 or 2016: 32-bit Office and 64-bit Windows

    cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /dstatus

    cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /unpkey:XXXXX

    cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
  #>
function Option2 {
  do {
    clear-host
    WriteConsole -Text "--------------------[ Office 2019 or 2016: 32-bit Office and 64-bit Windows ]--------------------" -ForegroundColor "Yellow" -NewLine 1
    $option = ShowSubMenu

    switch ($option) {
      "1" {
        cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /dstatus
        Pause
      }

      "2" {
        $OfficeProductKey = Read-Host "Enter the last 5 digits of the Office Product Key"
        cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /unpkey:$OfficeProductKey
        Pause
      }

      "3" {
        $OfficeProductKey = GetOfficeProductKey
        cscript "C:\Program Files (x86)\Microsoft Office\Office16\OSPP.VBS" /inpkey:$OfficeProductKey
        Pause
      }

      "4" { WriteConsole -Text "`nExiting..." -ForegroundColor "Yellow" -NewLine 2 -Wait 1 }

      Default { WriteConsole -Text "`nInvalid option" -ForegroundColor "Red" -NewLine 2 -Wait 3 }
    }
  } while ($option -ne "4")
}

<#
  Office 2019 or 2016: 64-bit Office and 64-bit Windows

    cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus

    cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /unpkey:XXXXX

    cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /inpkey:XXXXX-XXXXX-XXXXX-XXXXX-XXXXX
  #>
function Option3 {
  do {
    clear-host
    WriteConsole -Text "--------------------[ Office 2019 or 2016: 64-bit Office and 64-bit Windows ]--------------------" -ForegroundColor "Yellow" -NewLine 1
    $option = ShowSubMenu

    switch ($option) {
      "1" {
        cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /dstatus
        Pause
      }

      "2" {
        $OfficeProductKey = Read-Host "Enter the last 5 digits of the Office Product Key"
        cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /unpkey:$OfficeProductKey
        Pause
      }

      "3" {
        $OfficeProductKey = GetOfficeProductKey
        cscript "C:\Program Files\Microsoft Office\Office16\OSPP.VBS" /inpkey:$OfficeProductKey
        Pause
      }

      "4" { WriteConsole -Text "`nExiting..." -ForegroundColor "Yellow" -NewLine 2 -Wait 1 }

      Default { WriteConsole -Text "`nInvalid option" -ForegroundColor "Red" -NewLine 2 -Wait 3 }
    }
  } while ($option -ne "4")
}

do {
  $option = ShowMenu

  switch ($option) {
    "1" { Option1 }
    "2" { Option2 }
    "3" { Option3 }
    "4" { WriteConsole -Text "`nExiting..." -ForegroundColor "Yellow" -NewLine 2 -Wait 2 }
    Default { WriteConsole -Text "`nInvalid option" -ForegroundColor "Red" -NewLine 2 -Wait 3 }
  }
} while ($option -ne "4")