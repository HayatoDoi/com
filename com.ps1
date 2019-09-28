
# default params
[string] $newLine  = "CR" # CR(default), CR+LF, LF
[string] $encoding = "UTF-8" # UTF-8(default), UTF-16, Shift_JIS
[int] $portSpeed   = 9600

[double] $version = 1.0
[string] $copyright = "(c) 2019, Hayato Doi."
function sirial() {
    param (
        $comNum
    )

    $c = New-Object System.IO.Ports.SerialPort $comNum, $portSpeed, ([System.IO.Ports.Parity]::None)

    $c.DtrEnable = $true
    $c.RtsEnable = $true

    # no handshake
    $c.Handshake=[System.IO.Ports.Handshake]::None

    # set new line
    if ($newLine -eq "CR+LF") {
        $c.NewLine = "`r`n"
    }
    elseif ($newLine -eq "LF") {
        $c.NewLine = "`n"
    }
    else {
        $c.NewLine = "`r"
    }

    # set test code
    $c.Encoding=[System.Text.Encoding]::GetEncoding($encoding)
    
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
                if ($funStr -match "^ESC\[\d*(D|J)$"){
                    [string] $noEscStr = $funStr -replace "^ESC\["
                    [int] $num = $noEscStr -replace "(D|J)"
                    [string] $dORj = $noEscStr -replace "\d*"
                    $cursorLeft =  [System.Console]::CursorLeft
                    $cursorTop = [System.Console]::CursorTop
                    if ($dORj -eq "D") {
                        [System.Console]::SetCursorPosition($cursorLeft - $num, $cursorTop)
                        Write-Host -NoNewline (" " * $num)
                        [System.Console]::SetCursorPosition($cursorLeft - $num, $cursorTop)
                    }
                    # $dORj -eq "D"
                    else {
                        # todo: fix
                        [System.Console]::SetCursorPosition($cursorLeft, $cursorTop - $num)
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
function printVersion {
    Write-Host("com command. ver {0:f2}" -f $version)
    Write-Host "$copyright"
}
#--- main start ---
# args check
If($args.Length -eq 1){
    if($args[0] -match "^com\d$"){
        sirial $args[0]
        exit
    }
    elseif ($args[0] -match "^((list)|(ls))$") {
        comList
        exit
    }
    elseif ($args[0] -match "^version$") {
        printVersion
        exit
    }
}
Write-Host "argument error"
