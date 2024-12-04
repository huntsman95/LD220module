# HP LD220 VFD PowerShell Module

## Summary
Interface with the HP LD220 VFD POS Display via USB Serial

## Usage

```powershell
Import-Module LD220

# Connect to the VFD
Connect-LD220 -PortName COM15
# -- or for auto-detection --
$USBCOM = Get-PnpDevice -Class Ports `
  | Where-Object { 'USB\VID_03F0&PID_3524' -in $_.HardwareID } `
  | Select-Object -expand friendlyname `
  | Select-String '(COM\d+)'
$USBCOMPORT = $USBCOM.Matches.Groups[1].Value
Connect-LD220 -PortName $USBCOMPORT

# Write text to the VFD
Write-LD220 "Hello World"

# Disconnect from the VFD
Disconnect-LD220
```

## List of all commands
```powershell
Connect-LD220
Disconnect-LD220
Send-LD220Command
Write-LD220Text
Write-LD220TextXY
Write-LD220RAWXY
Invoke-LD220TypedTextEffect
Show-LD220Time
Clear-LD220Screen
Initialize-LD220Screen
Invoke-LD220Blink
Stop-LD220Blink
Set-LD220CursorDisplayMode
Write-LD220ProgressBar
Set-LD220CodePage
Invoke-LD220Marquee
```