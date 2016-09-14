#region About
# Stript to control ATEM with X-keys 60
#
# Requires SwitcherLib from https://ianmorrish.wordpress.com/v-ise/atem/
# Requires VISE_Xkeys from https://ianmorrish.wordpress.com/v-ise/x-keys/
# Optional: Json editor for XKeys config file from http://bit.ly/2cuZggS
#
# Button layout
#
# 79  71      55  47  39  31      15   7
# 78  70      54  46  38  30      14   6
# 77  69      53  45  37  29      13   5
# 76  68      52  44  36  28      12   4
# 75  67      51  43  35  27      11   3
# --------------------------------------
# 73  65  57  49  41  33  25  17   9   1
# 72  64  56  48  40  32  24  16   8   0
# For Red background, 2xColumn number
#endregion

#region ATEM setup
Switch ($env:computername){
    VideoPC {add-type -path 'C:\Users\imorrish\Documents\WindowsPowerShell\Modules\ATEM\SwitcherLib.dll'
            add-type -path 'C:\Users\imorrish\onedrive\powershell\x-keys\VISE_Xkeys.dll'}
    Default {add-type -path 'documents\windowspowershell\SwitcherLib.dll'
            add-type -path 'documents\windowspowershell\VISE_Xkeys.dll'}

}
[bool]$Global:Debug = $false
$Global:atem = New-Object SwitcherLib.Switcher("192.168.1.8")
$atem.Connect()
$me=$atem.GetMEs()
$me1=$me[0]
try{$me2=$me[1]; $Global:haveME2 = $true}
catch{$Global:haveME2 = $false}
$Global:activeME = $me1
$Global:txMode = $activeME.TransitionStyle

$Global:Program = $activeME.Program
$Global:Preview = $activeME.Preview
$Global:CurrentTransition = $activeME.TransitionStyle.ToString()

$MediaPlayers = $atem.GetMediaPlayers()
$Global:MP1=$MediaPlayers[0]
$Global:MP2=$MediaPlayers[1]

$Global:Aux = $atem.GetAuxInputs()
#Global valuses used in event
$global:loopColor = $false

$ColorGen= $atem.GetColorInputs()
$ColorGen1 = $ColorGen[0]
$ColorGen1.Luma =.5
$ColorGen1.Saturation = 1
$global:InputList = @(1,3,4,6,3020)
$global:InputCurrent = 1
$global:TimeOnInput = 1
$global:LoopInputs = $false
#$global:LoopInputs = $true
#endregion

#region X-keys setup
$Global:Xkeys = new-object VISE_Xkeys.xkeys
$Devices = $Xkeys.GetDevices()
$Devices
$Global:xkeysActiveDevice = $Devices[0].DeviceID 
$Global:Xkeys.NewDevice($xkeysActiveDevice)
$numericKeys = @{55=1;47=2;39=3;31=4;54=5;46=6;38=7;30=8;53=9;45=10;37=11;29=12;52=13;44=14;36=15;28=16;51=17;43=18;35=19;27=20}

$ProgramKeys = @{ 73=1; 65=2; 57=3; 49=4; 41=5; 33=6; 25=7; 17=8}
$previewKeys = @{ 72=1; 64=2; 56=3; 48=4; 40=5; 32=6; 24=7; 16=8}
[string]$Global:KeypadMode=""
if($Global:Program -gt 0 -And $Global:Program -lt 8){
    #turn on new program led
    $Global:xkeys.SendData($Global:xkeysActiveDevice,81 - ($Global:Program*8),1)
}
if($Global:Preview -gt 0 -And $Global:Preview -lt 8){
      #turn on new Preview led
      $Global:xkeys.SendData($Global:xkeysActiveDevice,80 - ($Global:Preview*8),1)
}
#endregion

#region Xkeys actions
function LoadXkeys(){
    $xkFile = ConvertFrom-Json (get-content "C:\Users\imorrish\OneDrive\PowerShell\BA-ISE\xk60file.json" -raw)
    $Global:xkCommands = @{}
    #$xkFile | get-member -MemberType NoteProperty | ForEach-Object{ConvertFrom-Json $_.vlaue} | ForEach-Object{$xkCommands.add($_.name,$xkFile."$($_.name)")}
    foreach($key in $xkFile.keys){
        $xkCommands.add($key.KeyID,$key.Script)
    }
}
LoadXKeys

function HandleXkey($keyId) 
{
    if($Global:Debug -eq $true){write-host "Key pressed - $($keyId) "}
    if($xkCommands.ContainsKey($keyId))
    {
        try{invoke-expression  $xkCommands.Get_Item($keyId)}
        catch{write-host "error: $($error)"}
    }
    else {write-host "Key not defined $($keyID)"}
}
Unregister-Event -SourceIdentifier KeyPressed -ErrorAction SilentlyContinue
$MyEvent = Register-ObjectEvent -InputObject $Xkeys -EventName KeyPressed -SourceIdentifier KeyPressed -Action {HandleXkey($event.sourceEventArgs.KeyID)} 
#endregion


#region Timer
$timer = New-Object System.Timers.Timer
$timer.Interval = 500 
$timer.AutoReset = $true
$sourceIdentifier = "TimerJob"
$timerAction = { 
    #update leds
    #Program
    $CurrentProgram = $Global:activeME.Program
    if($Global:Program -ne $CurrentProgram){
        if($Global:Debug -eq $true){write-host "Program changed from $($Global:Program) to $($CurrentProgram) "}
        
        if($Global:Program -gt 0 -And $Global:Program -lt 8){
            #turn off current LED
            $xkeys.SendData($Global:xkeysActiveDevice,161 - ($Global:Program*8),0)
            if($CurrentProgram -gt 0 -And $CurrentProgram -lt 8){
                #turn on new program led
                $Global:xkeys.SendData($Global:xkeysActiveDevice,161 - ($CurrentProgram*8),1)
            }
        }
        if($Global:Program = 3010){$Global:xkeys.SendData($Global:xkeysActiveDevice,155,0)}
        if($Global:Program = 3020){$Global:xkeys.SendData($Global:xkeysActiveDevice,147,0)}
        $Global:Program = $CurrentProgram
    }
    else{
            #no program led was on
            if($CurrentProgram -gt 0 -And $CurrentProgram -lt 8){
                #turn on new program led
                $Global:xkeys.SendData($Global:xkeysActiveDevice,161 - ($CurrentProgram*8),1)
                $Global:Program = $CurrentProgram
            }
    }
    #Preview
    $CurrentPreview = $Global:activeME.Preview
    if($Global:Preview -ne $CurrentPreview){
        if($Global:Debug -eq $true){write-host "Preview changed from $($Global:Preview) to $($CurrentPreview) "}
        if($Global:Preview -gt 0 -And $Global:Preview -lt 8){
            #turn off current LED
            $xkeys.SendData($Global:xkeysActiveDevice,80 - ($Global:Preview*8),0)
            if($CurrentPreview -gt 0 -And $CurrentPreview -lt 8){
                #turn on new Preview led
                $Global:xkeys.SendData($Global:xkeysActiveDevice,80 - ($CurrentPreview*8),1)
            }
        }
        if($Global:Preview = 3010){$Global:xkeys.SendData($Global:xkeysActiveDevice,75,0)}
        if($Global:Preview = 3020){$Global:xkeys.SendData($Global:xkeysActiveDevice,67,0)}
        $Global:Preview = $CurrentPreview
    }
    else{
            #no program led was on
            if($CurrentPreview -gt 0 -And $CurrentPreview -lt 9){
                #turn on new preview led
                $Global:xkeys.SendData($Global:xkeysActiveDevice,80 - ($CurrentPreview*8),1)
                $Global:Preview = $CurrentPreview
            }
    }
        
    if($Global:loopColor -eq $true){
        $Global:ColorGen1.Hue =$Global:hue
        $Global:hue++
        if($Global:hue -eq 357){$Global:hue=1}
    }
    if($LoopInputs -eq $true){
        if($Global:TimeOnInput -eq 10){
            #get next input
            $Global:InputCurrent++
            if($Global:InputCurrent -gt $Global:InputList.Count){
                $Global:InputCurrent=1
            }
            $Global:activeME.Preview = $Global:InputList[$Global:InputCurrent]
            $Global:activeME.Cut()
            $Global:TimeOnInput=0
        }
        $Global:TimeOnInput++
    }
    #Set transition LED
    $METransitionStyle = $Global:activeME.TransitionStyle.ToString()
    if ($METransitionStyle -ne $Global:CurrentTransition){
        switch ($METransitionStyle){
            bmdSwitcherTransitionStyleMix{
                $Global:xkeys.SendData($Global:xkeysActiveDevice,8,0)
                $Global:xkeys.SendData($Global:xkeysActiveDevice,9,0)
                $Global:CurrentTransition = "bmdSwitcherTransitionStyleMix"

            }
            bmdSwitcherTransitionStyleDIP{
                $Global:xkeys.SendData($Global:xkeysActiveDevice,8,0)
                $Global:xkeys.SendData($Global:xkeysActiveDevice,9,1)
                $Global:CurrentTransition = "bmdSwitcherTransitionStyleDIP"

            }
            bmdSwitcherTransitionStyleWipe{
                $Global:xkeys.SendData($Global:xkeysActiveDevice,8,1)
                $Global:xkeys.SendData($Global:xkeysActiveDevice,9,0)
                $Global:CurrentTransition = "bmdSwitcherTransitionStyleWipe"

            }
        }
    }

}

Unregister-Event $sourceIdentifier -ErrorAction SilentlyContinue
$timer.stop()
$start = Register-ObjectEvent -InputObject $timer -SourceIdentifier $sourceIdentifier -EventName Elapsed -Action $timeraction
$timer.start()
#endregion

#region Xkey functions
function USKAutoTransition(){
        $me1.TransitionSelection=2
        Start-Sleep -Milliseconds 1
        $me1.AutoTransition()
        Start-Sleep -Milliseconds 10 #give it a chance to start
        Start-Sleep 2
        #Key is now onair so remove from next transition (Turn on BKGD)
        $me1.TransitionSelection=1 
}

function ToggleActiveME(){
    if($Global:activeME -eq $me1){
        $Global:activeME = $me2;
        $Global:xkeys.SendData($Global:xkeysActiveDevice,95,1);
        $Global:xkeys.SendData($Global:xkeysActiveDevice,15,0)
    }
    else {
        $Global:activeME = $me1;
        $Global:xkeys.SendData($Global:xkeysActiveDevice,15,1)
        $Global:xkeys.SendData($Global:xkeysActiveDevice,95,0)
    }
}
function changeTransition($newTransitionStyle){

    switch($Global:activeME.TransitionStyle){
        bmdSwitcherTransitionStyleMix{
            $Global:activeME.TransitionStyle=$newTransitionStyle
            break;
        }
        bmdSwitcherTransitionStyleDIP{
            if($newTransitionStyle="Dip"){
                $Global:activeME.TransitionStyle="Mix"
            }
            else{
                $Global:activeME.TransitionStyle="Dip"
            }
            break;
        }
        bmdSwitcherTransitionStyleWipe{
            if($newTransitionStyle="Wipe"){
                $Global:activeME.TransitionStyle="Mix"
            }
            else{
                $Global:activeME.TransitionStyle="Wipe"
            }
            break;
        }
    }
}
function keypad([int]$number){
    switch($Global:KeypadMode){
        Macro{
            Write-host "Running Macro $($number)"
            $atem.RunMacro($number-1)
            $Global:KeypadMode=""
            $Global:xkeys.SendData(1,150,0)
            break;
        }
        MP1{
            $MP1.MediaStill = $number-1
            $Global:KeypadMode=""
            if ($Global:Program -eq 3010){$Global:xkeys.SendData(1,155,1)}
            elseif ($Global:Preview -eq 3010){$Global:xkeys.SendData(1,75,1)}
            else{$Global:xkeys.SendData(1,75,0);$Global:xkeys.SendData(1,155,1)}
            
        }
        MP2{
             $MP2.MediaStill = $number-1
             $Global:KeypadMode=""
        }
        Aux1{
            if($aux.Count -gt 0){$aux[0].Input = $number;}
            $Global:xkeys.SendData(1,78,0)
        }
        Aux2{
            if($aux.Count -gt 1){$aux[1].Input = $number;}
            $Global:xkeys.SendData(1,77,0)
        }
        Aux3{
            if($aux.Count -gt 0){$aux[2].Input = $number;}
             $Global:xkeys.SendData(1,76,0)
        }
    }
}
#endregion