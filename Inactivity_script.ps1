Foreach($i in 1..480)
{
$wshell = New-Object -ComObject Wscript.Shell
$wshell.SendKeys("%{SCROLLLOCK}")
Start-Sleep -s 58
}
