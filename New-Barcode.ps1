<#
.SYNOPSIS
  Create Barcodes
.DESCRIPTION
  This Function can create Barcodes, DataMatrix Codes, QR Codes and more. Its a powershell Wrapper around the zint.exe cli

.EXAMPLE
  New-Barcode -BarcodeType "CODE128" -Content "1234567890"
.PARAMETER BarcodeType 
  Type of the Barcode
.PARAMETER Content 
  Data whats saved in the Barcode
.PARAMETER NoTitle
  Switch on when the title should be hidden
.PARAMETER Scale
  Scale of the Barcode
.PARAMETER OutputPath
  Path of the Output
.PARAMETER FileName
  Filename of the Output file
.PARAMETER BackgroundColor
  Background color of the Barcode
.PARAMETER ForegroundColor 
  Foreground color of the Barcode
.PARAMETER Format
  File Format of the Output
.PARAMETER Height
  Hight of the Barcode (only for 1D Codes)
.PARAMETER DotSize 
  Dotsize only for QR/Datamatrix Codes
.PARAMETER Rotation 
  Rotation of the Barcode
.PARAMETER ECI
  ECI-MOde for special character sets
.PARAMETER Version
  Version of the QR Code (0 = Auto)
.PARAMETER ZintPath
  Path to the Zint.exe
#>


function New-Barcode {
  [CmdletBinding()]
  param (
      [Parameter(Mandatory=$true)]
      [ValidateSet(
        "CODE11", "CODE39", "CODE93", "CODE128", "CODE128B",
        "DATAMATRIX", "QRCODE", "AZTEC",
        "GS1_128", "GS1_128_CC"
    )]
    
      [string]$BarcodeType,

      [Parameter(Mandatory=$true)]
      [string]$Content,

      [Parameter(Mandatory=$false)]
      [switch]$NoTitle,

      [Parameter(Mandatory=$false)]
      [ValidateRange(1,200)]
      [int]$Scale = 10,

      [Parameter(Mandatory=$false)]
      [string]$OutputPath = $PSScriptRoot,

      [Parameter(Mandatory=$false)]
      [string]$FileName = "$($BarcodeType)_$($Content)",

      [Parameter(Mandatory=$false)]
      [string]$BackgroundColor = "FFFFFF",

      [Parameter(Mandatory=$false)]
      [string]$ForegroundColor = "000000",

      [Parameter(Mandatory=$false)]
      [ValidateSet("png","svg","eps")]
      [string]$Format = "png",      

      [Parameter(Mandatory=$false)]
      [ValidateRange(1,2000)]
      [int]$Height = 50,

      [Parameter(Mandatory=$false)]
      [ValidateRange(1,20)]
      [int]$DotSize = 1,

      [Parameter(Mandatory=$false)]
      [ValidateSet(0,90,180,270)]
      [int]$Rotation = 0,
    
      [Parameter(Mandatory=$false)]
      [string]$ECI = "",

      [Parameter(Mandatory=$false)]
      [int]$Version  = 0,
    
      [Parameter(Mandatory=$false)]
      $ZintPath = ".\zint.exe"        
  )

  if (-Not (Test-Path $ZintPath)) {
      Write-Error "Zint wurde nicht gefunden!"
      return
  }

  $OutputFile = "$OutputPath\$FileName.$Format"

  $cmd = "$ZintPath -b $BarcodeType -o `"$OutputFile`" -d `"$Content`" --fg $ForegroundColor --bg $BackgroundColor --rotate $Rotation"

  if ($NoTitle) { $cmd += " --notext" }

  if ($Height -ne 50) { $cmd += " --height $Height" }

  if ($BarcodeType -in @("QRCODE", "DATAMATRIX", "AZTEC", "MICROQR", "GRIDMATRIX")) { 
      $cmd += " --scale $Scale" 
  }
  if ($DotSize -ne 1) { $cmd += " --dotsize $DotSize" }

  if ($BarcodeType -eq "QRCODE" -and $Version -ne 0) { 
      $cmd += " --vers $Version" 
  }

  if ($ECI -ne "") { $cmd += " --eci $ECI" }


  Invoke-Expression $cmd

  if (Test-Path $OutputFile) {
      $Output = [PSCustomObject]@{
        Name = $FileName
        FilePath = $OutputFile
      }
      #return $Output
  } else {
      Write-Error "Fehler: Barcode konnte nicht erstellt werden."
  }
}
