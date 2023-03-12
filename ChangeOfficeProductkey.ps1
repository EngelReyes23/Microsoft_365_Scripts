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

function ShowMenu {
  while ($true) {
    Clear-Host
    WriteConsole("--------------------[ Change Office Product Key ]--------------------", "green", 2)
    WriteConsole("1- Office 2019 or 2016: 32-bit Office and 32-bit Windows", "white")
    WriteConsole("2- Office 2019 or 2016: 32-bit Office and 64-bit Windows", "white")
    WriteConsole("3- Office 2019 or 2016: 64-bit Office and 64-bit Windows", "white")
    WriteConsole("4- Exit", "white", 2)
    return Read-Host "Enter your choice"
  }
}

function ShowSubMenu {
  while ($true) {
    WriteConsole("1- Show Office Product Key", "white")
    WriteConsole("2- Remove Office Product Key", "white")
    WriteConsole("3- Change Office Product Key", "white")
    WriteConsole("4- Exit", "white", 2)
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
    WriteConsole("--------------------[ Office 2019 or 2016: 32-bit Office and 32-bit Windows ]--------------------", "green", 2)
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

      "4" { WriteConsole("Exit", "white", 2) }

      Default {
        WriteConsole("Invalid option", "red", 2)
        Pause
      }
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
    WriteConsole("--------------------[ Office 2019 or 2016: 32-bit Office and 64-bit Windows ]--------------------", "green", 2)
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

      "4" { WriteConsole("Exit", "white", 2) }

      Default {
        WriteConsole("Invalid option", "red", 2)
        Pause
      }
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
    WriteConsole("--------------------[ Office 2019 or 2016: 64-bit Office and 64-bit Windows ]--------------------", "green", 2)
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

      "4" { WriteConsole("Exit", "white", 2) }

      Default {
        WriteConsole("Invalid option", "red", 2)
        Pause
      }
    }
  } while ($option -ne "4")
}

do {
  $option = ShowMenu

  switch ($option) {
    "1" { Option1 }
    "2" { Option2 }
    "3" { Option3 }
    "4" { WriteConsole("Bye", "white", 2) }
    Default {
      WriteConsole("Invalid option", "red", 2)
      Pause
    }
  }
} while ($option -ne "4")