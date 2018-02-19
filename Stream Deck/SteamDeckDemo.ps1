#
# ATEM Legato Stream Deck controler
#
# By Ian Morrish
# https://ianmorrish.wordpress.com
#
Set-Location $env:USERPROFILE\documents
#region ATEM Setup
Try{
$ATEMipAddress = (Get-ItemProperty -path 'HKCU:\Software\Blackmagic Design\ATEM Software Control').ipAddress
add-type -path 'windowspowershell\SwitcherLib.dll'
$Global:atem = New-Object SwitcherLib.Switcher($ATEMipAddress)
$atem.Connect()
}
catch{
write-host "Can't connect to ATEM on $($ATEMipAddrss)."
Write-Host "ATEM controle software must be installed and have connected to switcher at least one time"
Write-Host "switcherlib.dll and StreamDeckSharp.dll must be in [user]\documents\WindowsPowerShell"
}
$me=$atem.GetMEs()
$Global:me1=$me[0]
$Global:activeME = $me1
$Global:Program = $activeME.Program
$Global:Preview = $activeME.Preview

$MediaPlayers = $atem.GetMediaPlayers()
$Global:MP1=$MediaPlayers[0]
$Global:MP2=$MediaPlayers[1]

#endregion

#region Stream Deck setup
add-type -path 'WindowsPowerShell\StreamDeckSharp.dll'
$deckInterface = [StreamDeckSharp.StreamDeck]::FromHID()
#$deckInterface | Get-Member
#$deckInterface.NumberOfKeys
#$deckInterface.ShowLogo()

$buttonMapping = @(9,8,7,6,14,13,12,11)

$imgCut = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\Cut.png")
$imgFade = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\Fade.png")
$imgCrop = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\Crop.png")
$imgMacro = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\Macro.png")
$imgMedia = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\Media.png")
$imgResize = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\Resize.png")
$1Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\1White.png")
$1Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\1Red.png")
$1Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\1Green.png")
$2Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\2White.png")
$2Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\2Red.png")
$2Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\2Green.png")
$3Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\3White.png")
$3Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\3Red.png")
$3Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\3Green.png")
$4Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\4White.png")
$4Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\4Red.png")
$4Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\4Green.png")
$5Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\5White.png")
$5Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\5Red.png")
$5Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\5Green.png")
$6Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\6White.png")
$6Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\6Red.png")
$6Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\6Green.png")
$7Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\7White.png")
$7Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\7Red.png")
$7Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\7Green.png")
$8Label = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\8White.png")
$8Active = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\8Red.png")
$8Preview = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\8Green.png")
$POSH = [StreamDeckSharp.StreamDeckKeyBitmap]::FromFile(".\Pictures\icon\PowerShell72x72.png")

$ButtonInit = @($POSH,$imgMacro,$imgMedia,$imgCrop,$imgResize,`
                $imgFade,$4Label,$3Label,$2Label,$1Label,`
                $imgCut,$8Label,$7Label,$6Label,$5Label
)
$ProgramButtons = @($1Active,$2Active,$3Active,$4Active,$5Active,$6Active,$7Active,$8Active)
$PreviewButtons = @($1Preview,$2Preview,$3Preview,$4Preview,$5Preview,$6Preview,$7Preview,$8Preview)
$BlankButtons = @($1Label,$2Label,$3Label,$4Label,$5Label,$6Label,$7Label,$8Label)
$deckInterface.SetBrightness(90)

for ($i = 0; $i -lt $deckInterface.NumberOfKeys; $i++){
    $deckInterface.SetKeyBitmap($i, $ButtonInit[$i].CloneBitmapData())
}
#endregion

#region set active program and preview
if($Global:Program -gt 0 -And $Global:Program -lt 8){
    #turn on new program led
    $deckInterface.SetKeyBitmap($buttonMapping[$Global:Program-1],$ProgramButtons[$Global:Program-1].CloneBitmapData())
}
if($Global:Preview -gt 0 -And $Global:Preview -lt 8){
      #turn on new Preview led
      $deckInterface.SetKeyBitmap($buttonMapping[$Global:Preview-1],$PreviewButtons[$Global:Preview-1].CloneBitmapData())
}
#endregion
#region Utilities
function USKAutoTransition(){
        $me1.TransitionSelection=2
        Start-Sleep -Milliseconds 1
        $me1.AutoTransition()
        Start-Sleep -Milliseconds 10 #give it a chance to start
        Start-Sleep 2
        #Key is now onair so remove from next transition (Turn on BKGD)
        $me1.TransitionSelection=1 
}

#endregion

# Key down actions

function inputevent($key){
    if($key.IsDown)
        {
        switch($key.key)
            {
                5{$Global:me1.AutoTransition()}
                10{$Global:me1.Cut()}
                4{}
                3{}
                2{}
                1{}
                0{}
                9{$me1.Preview=1}
                8{$me1.Preview=2}
                7{$me1.Preview=3}
                6{$me1.Preview=4}
                14{$me1.Preview=5}
                13{$me1.Preview=6}
                12{$me1.Preview=7}
                11{$me1.Preview=8}
            }
        }
}

#Set up key event

Unregister-Event -SourceIdentifier buttonPress -ErrorAction SilentlyContinue #incase we are re-running the script
$KeyEvent = Register-ObjectEvent -InputObject $deckInterface -EventName KeyPressed -SourceIdentifier buttonPress -Action {inputevent($eventArgs)}

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
            if($Global:Program -gt 0 -And $Global:Program -lt 9){
            #turn off current LED
            $deckInterface.SetKeyBitmap($buttonMapping[$Global:Program-1],$BlankButtons[$Global:Program-1].CloneBitmapData())
            if($CurrentProgram -gt 0 -And $CurrentProgram -lt 9){
                #turn on new program led
                $deckInterface.SetKeyBitmap($buttonMapping[$CurrentProgram-1],$ProgramButtons[$CurrentProgram-1].CloneBitmapData())
            }
        }

        $Global:Program = $CurrentProgram
    }
    else{
            #no program led was on
            if($CurrentProgram -gt 0 -And $CurrentProgram -lt 9){
                #turn on new program led
                $deckInterface.SetKeyBitmap($buttonMapping[$CurrentProgram-1],$ProgramButtons[$CurrentProgram-1].CloneBitmapData())
                $Global:Program = $CurrentProgram
            }
    }
    #Preview
    $CurrentPreview = $Global:activeME.Preview
    if($Global:Preview -ne $CurrentPreview){
        
        if($Global:Preview -gt 0 -And $Global:Preview -lt 9){
            #turn off current LED
            $deckInterface.SetKeyBitmap($buttonMapping[$Global:Preview-1],$BlankButtons[$Global:Preview-1].CloneBitmapData())
            if($CurrentPreview -gt 0 -And $CurrentPreview -lt 9){
                #turn on new Preview led
                $deckInterface.SetKeyBitmap($buttonMapping[$CurrentPreview-1],$PreviewButtons[$CurrentPreview-1].CloneBitmapData())
            }
        }

        $Global:Preview = $CurrentPreview
    }
    else{
            #no program led was on
            if($CurrentPreview -gt 0 -And $CurrentPreview -lt 9){
                #turn on new preview led
                $deckInterface.SetKeyBitmap($buttonMapping[$CurrentPreview-1],$PreviewButtons[$CurrentPreview-1].CloneBitmapData())
                $Global:Preview = $CurrentPreview
            }
    }
}

# Start the timer
Unregister-Event $sourceIdentifier -ErrorAction SilentlyContinue
$timer.stop()
$start = Register-ObjectEvent -InputObject $timer -SourceIdentifier $sourceIdentifier -EventName Elapsed -Action $timeraction
$timer.start()
#endregion
