Start-Sleep -Seconds 5
Add-Type -AssemblyName System.Windows.Forms
$n=new-object System.Windows.Forms.NotifyIcon
$n.Icon=[System.Drawing.SystemIcons]::Information
$n.Visible=$true
$n.ShowBalloonTip(5000,'test','hello',[System.Windows.Forms.ToolTipIcon]::Info)
Start-Sleep -Seconds 6
$n.Dispose()
