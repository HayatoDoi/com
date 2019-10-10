#!/usr/bin/env pwsh

$params = @{
    "newLine" = @{
        "value" = [string] "CR" # CR(default), CR+LF, LF
        "value_regex" = [string] "^((CR)|(CR\+LF)|(LF))$"
    }
    "encoding" = @{
        "value" = [string] "UTF-8" # UTF-8(default), UTF-16, Shift_JIS
        "value_regex" = [string] "^((UTF\-8)|(UTF\-16)|(Shift_JIS))$"
    }
    "portSpeed" = @{
        "value" = [int] 9600
        "value_regex" = [string] "^\d*$"
    }
    "comNum" = @{
        "value" = [string] ""
        "value_regex" = [string] "^((com)|(COM))\d+$"
    }
}

[double] $version = 1.3
[string] $copyright = "(c) 2019, Hayato Doi."
function sirialStart() {

    $c = New-Object System.IO.Ports.SerialPort $params["comNum"]["value"], $params["portSpeed"]["value"], ([System.IO.Ports.Parity]::None)

    $c.DtrEnable = $true
    $c.RtsEnable = $true

    # no handshake
    $c.Handshake=[System.IO.Ports.Handshake]::None

    # set new line
    if ($params["newLine"]["value"] -eq "CR+LF") {
        $c.NewLine = "`r`n"
    }
    elseif ($params["newLine"]["value"] -eq "LF") {
        $c.NewLine = "`n"
    }
    else {
        $c.NewLine = "`r"
    }

    # set test code
    $c.Encoding=[System.Text.Encoding]::GetEncoding($params["encoding"]["value"])
    
    # Register serial receive event
    # when received, output to console 
    [Boolean] $inFunStrFlag = $FALSE
    [String] $funStr = ""
    $d = Register-ObjectEvent -InputObject $c -EventName "DataReceived" -Action {
        param (
            [System.IO.Ports.SerialPort]$sender,
            [System.EventArgs]$e
        )
        $buf = $sender.ReadExisting()
        $buf.ToCharArray() | foreach {
            if ($inFunStrFlag -eq $TRUE -or [int]$_ -eq 27){
                $inFunStrFlag = $TRUE
                if ( [int]$_ -eq 27 ){
                    $funStr += "ESC"
                }
                else {
                    $funStr += $_
                }
                if ($funStr -match "^ESC\[\d*(D|J|C)$"){
                    [string] $noEscStr = $funStr -replace "^ESC\["
                    [int] $num = $noEscStr -replace "(D|J|C)"
                    [string] $dORj = $noEscStr -replace "\d*"
                    $cursorLeft =  [System.Console]::CursorLeft
                    $cursorTop = [System.Console]::CursorTop
                    if ($dORj -eq "C") {
                        [System.Console]::SetCursorPosition($cursorLeft + $num, $cursorTop)
                    }
                    elseif ($dORj -eq "D") {
                        [System.Console]::SetCursorPosition($cursorLeft - $num, $cursorTop)
                    }
                    # $dORj -eq "J"
                    else {
                        if ($num -eq 0) {
                            Write-Host -NoNewline (" " * ([System.Console]::WindowHeight - $cursorLeft) )
                            [System.Console]::SetCursorPosition($cursorLeft, $cursorTop)
                        }
                    }
                    $inFunStrFlag = $FALSE
                    $funStr = ""
                }
            }
            else {
                Write-Host -NoNewline $_
            }
        }
    }
    Try {
        # open com port
        $c.Open()
    }
    Catch {
        Write-Host $_.Exception.Message
        exit
    }
    Try {
        for (;;) {
            if ([Console]::KeyAvailable){
                $keyinfo = [Console]::ReadKey($true)
                if ($keyinfo.Key -eq "UpArrow" ) {
                    $c.Write((0x1B,0x5B,0x41),0,3)
                }
                elseif ($keyinfo.Key -eq "DownArrow") {
                    $c.Write((0x1B,0x5B,0x42),0,3)
                }
                elseif ($keyinfo.Key -eq "RightArrow") {
                    $c.Write((0x1B,0x5B,0x43),0,3)
                }
                elseif ($keyinfo.Key -eq "LeftArrow") {
                    $c.Write((0x1B,0x5B,0x44),0,3)
                }
                else{
                    $c.Write($keyinfo.KeyChar)
                }
            }
        }
    }
    Finally {
        Write-Host "`nbye."
        $c.Close()
        Unregister-Event $d.Name
        Remove-Job $d.Id
    }
}
function comList {
    (Get-WmiObject -query "SELECT * FROM Win32_PnPEntity" | Where {$_.Name -Match "COM\d+"}).name
}
function setParam {
    param (
        [string] $key,
        [string] $value
    )
    foreach($param in $params.GetEnumerator()){
        if ($key -eq $param.Key -and $value -match $param.Value["value_regex"]){
            $param.Value["value"] = $value
            return
        }
    }
    throw "error"
}
function checkParam {
    foreach($param in $params.GetEnumerator()){
        if (!($param.Value["value"] -match $param.Value["value_regex"])){
            throw "error"
        }
    }
}
function debugPrintPrams {
    Write-Host "Key       Value"
    foreach($param in $params.GetEnumerator()){
        Write-Host $param.key"       "$param.Value["value"]
    }
}
function loadIniFile {
    param (
        [string] $filepath
    )
    Get-Content $filepath | `
        where-object {($_ -notmatch '^\s*$') -and (!($_.TrimStart().StartsWith(";")))} | foreach {
            $param = $_.split("=", 2)
            try {
                setParam $param[0] $param[1]
            }
            catch {
                error "can't load config file"
            }
    }
}
function printVersion {
    Write-Host("com command. ver {0:f2}" -f $version)
    Write-Host "$copyright"
}
function printHelp {
    [int] $totalWidth = 30
    Write-Host "Usage: com [COM_PORT] [OPTION]... "
    Write-Host ""
    Write-Host "[COM_PORT]"
    Write-Host "You can get it by `"com ls`" command."
    Write-Host ""
    Write-Host "[OPTION]"
    Write-Host -NoNewline "ls, list".PadRight($totalWidth)
    Write-Host "Show com port list."
    Write-Host -NoNewline "--newLine <new line>".PadRight($totalWidth)
    Write-Host "Set new line. ( CR(default), CR+LF, LF )"
    Write-Host -NoNewline "--encoding <encoding>".PadRight($totalWidth)
    Write-Host "Set encoding. ( UTF-8(default), UTF-16, Shift_JIS )"
    Write-Host -NoNewline "--portSpeed <portSpeed>".PadRight($totalWidth)
    Write-Host "Set port speed. ( 9600(default) )"
    Write-Host -NoNewline "-c, --config <cofnig file>".PadRight($totalWidth)
    Write-Host "Load config file."
    Write-Host -NoNewline "-v, --version".PadRight($totalWidth)
    Write-Host "Show version."
    Write-Host -NoNewline "-h, --help".PadRight($totalWidth)
    Write-Host "Show help."
    Write-Host ""
    Write-Host "[Example]"
    Write-Host "> com ls"
    Write-Host "USB Serial Port (COM3)"
    Write-Host "> com COM3"
    Write-Host "# start sirial..."
    Write-Host "".PadRight($totalWidth, "-")
    Write-Host "$copyright"
}

function error {
    param (
        [string] $message
    )
    Write-Host $message
    exit 1
}

function argumentError {
    error "argument error.`nTry 'com --help' for more information."
}
#--- main start ---
# args check
for ($i = 0; $i -lt $args.Count; $i++) {
    if($args[$i] -match $params["comNum"]["value_regex"]){
        $params["comNum"]["value"] = $args[$i]
    }
    elseif ($args[$i] -match "^((\-\-newLine)|(\-\-encoding)|(\-\-portSpeed))$") {
        if ($i++ -ge $args.Count){
            argumentError
        }
        try {
            setParam $args[$i-1].Remove(0, 2) $args[$i]
        }
        catch {
            argumentError
        }
    }
    elseif ($args[$i] -match "^((\-\-config)|(\-c))$") {
        if ($i++ -ge $args.Count){
            argumentError
        }
        loadIniFile $args[$i]
        # $i++
    }
    elseif ($args[$i] -match "^((list)|(ls))$") {
        if ($args.Count -ne 1) {
            argumentError
        }
        comList
        exit
    }
    elseif ($args[$i] -match "((^\-\-version)|(\-v))$") {
        if ($args.Count -ne 1) {
            argumentError
        }
        printVersion
        exit
    }
    elseif ($args[$i] -match "((^\-\-help)|(\-h))$") {
        if ($args.Count -ne 1) {
            argumentError
        }
        printHelp
        exit
    }
    else {
        argumentError
    }
}

try {
    checkParam
}
catch {
    argumentError   
}

sirialStart