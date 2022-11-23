$s=New-Object -ComObject SAPI.SpVoice
$s.Speak("wanna see something really cool?")
[System.Windows.MessageBox]::Show('Would  you like to play a game?','Game input','YesNoCancel','Error')
