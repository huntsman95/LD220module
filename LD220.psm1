function Connect-LD220 {
    param (
        [Parameter(Mandatory = $false)]
        [string]$COM = 'COM15'
    )
    $global:port = New-Object System.IO.Ports.SerialPort $COM, 9600, None, 8, one
    $port.open()
}

function Disconnect-LD220 {
    $port.close()
}

function Send-LD220Command {
    param (
        [Parameter(Mandatory = $true)]
        [byte[]]$Command
    )
    $port.Write($Command, 0, $Command.Length)
}


function Write-LD220Text {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Text
        ,
        [Parameter(Mandatory = $false)]
        [switch]$Clear
    )
    if ($Clear) {
        Clear-LD220Screen
    }
    $textBytes = [System.Text.Encoding]::ASCII.GetBytes($Text)
    # return $payload
    Send-LD220Command -Command $textBytes
}

function Write-LD220TextXY {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 20)]
        [int]$ColStart
        ,
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 2)]
        [int]$Row
        ,
        [Parameter(Mandatory = $true)]
        [string]$Text
    )
    $textBytes = [System.Text.Encoding]::ASCII.GetBytes($Text)
    $payload = @([byte[]]@(0x1F, 0x24, $ColStart, $Row) + $textBytes) -as [byte[]]
    # return $payload
    Send-LD220Command -Command $payload
}

function Write-LD220RAWXY {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 20)]
        [int]$ColStart
        ,
        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 2)]
        [int]$Row
        ,
        [Parameter(Mandatory = $true)]
        [byte[]]$Text
    )
    $payload = @([byte[]]@(0x1F, 0x24, $ColStart, $Row) + $Text) -as [byte[]]
    # return $payload
    Send-LD220Command -Command $payload
}


function Invoke-LD220TypedTextEffect {
    param (
        [string]$Text,
        [int]$Delay = 50
    )
    $port.Write([byte[]]@(0x0c), 0, 1) #Clear Screen
    $port.Write([byte[]]@(0x1F, 0x02), 0, 2) #Vertical Scroll Mode
    $port.Write(@(0x1F, 0x43, 0x01), 0, 3) #Show cursor
    foreach ($char in $Text.ToCharArray()) {
        Write-LD220Text -Text $char
        Start-Sleep -Milliseconds $Delay
    }
    # Write-Host
    $port.Write([byte[]]@(0x1F, 0x01), 0, 2) #Overwrite Scroll Mode
    Start-Sleep -Seconds 2
    $port.Write(@(0x1F, 0x43, 0x00), 0, 3) #Hide cursor
}


function Show-LD220Time {
    $hour = (Get-Date).Hour
    $minute = (Get-Date).Minute
    Send-LD220Command -Command ([byte[]]@(0x1F, 0x54, $hour, $minute)) #Set time to 12:54 and count
}

function Clear-LD220Screen {
    Send-LD220Command -Command ([byte[]]@(0x0c))
}

function Initialize-LD220Screen {
    Send-LD220Command -Command ([byte[]]@(27, 64), 0, 2)
}

function Invoke-LD220Blink {
    param(
        [int]$IntervalMS = 50
    )
    $blinkInterval = [math]::Round($IntervalMS / 9)
    Send-LD220Command -Command 0x1F, 0x45, $blinkInterval, 0, 3
}

function Stop-LD220Blink {
    Send-LD220Command -Command 0x1F, 0x45, 0, 0, 3
}

function Set-LD220CursorDisplayMode {
    param (
        [ValidateSet('Off', 'On')]
        [string]$Mode = 'Off'
    )
    enum DISPLAYMODE {
        OFF = 0
        ON = 1
    }
    Send-LD220Command -Command 0x1F, 0x43, ([byte][DISPLAYMODE]::$Mode)
}

function Write-LD220ProgressBar {
    param (
        [Parameter(Mandatory = $true)]
        [ValidateRange(0, 100)]
        [int]$Percent
    )

    $width = 20

    $progressCharacter = 219
    $emptyCharacter = 32

    $fullBlocks = [Math]::Floor($Percent / (100 / $width))
    $emptyBlocks = $width - $fullBlocks
    $bar = [System.Collections.Generic.List[byte]]::new()

    for ($i = 0; $i -lt $fullBlocks; $i++) {
        [void]$bar.Add($progressCharacter)
    }
    for ($i = 0; $i -lt $emptyBlocks; $i++) {
        [void]$bar.Add($emptyCharacter)
    }
    $bar = $bar.ToArray()
    Write-LD220RAWXY -ColStart 1 -Row 2 -Text $bar
}

function Set-LD220CodePage {
    param (
        [ValidateSet('Default', 'Katakana')]
        [string]$Codepage = 'Default'
    )
    enum CODEPAGE {
        DEFAULT = 0
        KATAKANA = 1
    }
    Send-LD220Command -Command 0x1B, 0x74, ([byte][CODEPAGE]::$Codepage)
}

function Invoke-LD220Marquee {
    param (
        [string]$Text
        ,
        [int]$LineNumber = 1
        ,
        [int]$Interval = 80
        ,
        [int]$Padding = 20
    )

    [char]$padChar = [char]0x20
    [char[]]$pad = @()
    for ($i = 0; $i -lt $Padding; $i++) {
        $pad += $padChar
    }
    $Text = (($pad) -join '') + $Text
    $marquee = [byte[]]::new(20)
    $textBytes = [System.Text.Encoding]::UTF8.GetBytes($Text)

    $offset = 0
    while ($true) {
        for ($i = 0; $i -lt $marquee.Length; $i++) {
            $marquee[$i] = $textBytes[($offset + $i) % $textBytes.Length]
        }
        $marqueeTxt = [System.Text.Encoding]::ASCII.GetString($marquee)
        Write-LD220TextXY -ColStart 1 -Row $LineNumber -Text $marqueeTxt
        Start-Sleep -Milliseconds $Interval
        $offset = ($offset + 1) % $textBytes.Length
    }
}