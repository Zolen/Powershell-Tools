$voice = New-Object -ComObject SAPI.SPVoice
$voice.Rate = -3
 
function invoke-speech{
    param([Parameter(ValueFromPipeline=$true)][string] $say )
    process{
        $voice.Speak($say) | out-null;    
        }
    }
new-alias -name out-voice -value invoke-speech;