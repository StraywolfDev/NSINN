[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
$objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon 
$objNotifyIcon.Icon = "Tray_Notifications\Chip.ico"
$objNotifyIcon.BalloonTipIcon = "None" 
$objNotifyIcon.BalloonTipText = "NSINN Running" 
$objNotifyIcon.BalloonTipTitle = "NSINN"
$objNotifyIcon.Visible = $True 
$objNotifyIcon.ShowBalloonTip(5000)