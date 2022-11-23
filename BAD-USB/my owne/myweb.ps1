$s=New-Object -ComObject SAPI.SpVoice
$s.Speak("wanna see something really cool?")
$popmessage = "Hello bal-eter"


$readyNotice = New-Object -ComObject Wscript.Shell;$readyNotice.Popup($popmessage)