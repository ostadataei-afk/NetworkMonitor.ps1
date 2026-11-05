Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$window=New-Object Windows.Window
$window.Title=""
$window.WindowStyle="None"
$window.ResizeMode="NoResize"
$window.Width=370
$window.Height=260
$window.WindowStartupLocation="Manual"
$window.Background=[Windows.Media.Brushes]::Black
$root=New-Object Windows.Controls.Grid
$root.Margin=12
$r1=New-Object Windows.Controls.RowDefinition
$r1.Height="*"
$r2=New-Object Windows.Controls.RowDefinition
$r2.Height="Auto"
$r3=New-Object Windows.Controls.RowDefinition
$r3.Height="Auto"
$root.RowDefinitions.Add($r1)
$root.RowDefinitions.Add($r2)
$root.RowDefinitions.Add($r3)
$chartBorder=New-Object Windows.Controls.Border
$chartBorder.Background=[Windows.Media.Brushes]::Black
$chartBorder.BorderBrush=[Windows.Media.Brushes]::DarkSlateGray
$chartBorder.BorderThickness=1
$chartBorder.CornerRadius=8
$chartCanvas=New-Object Windows.Controls.Canvas
$chartBorder.Child=$chartCanvas
[Windows.Controls.Grid]::SetRow($chartBorder,0)
$root.Children.Add($chartBorder)|Out-Null
$rateGrid=New-Object Windows.Controls.Grid
$rateColumn1=New-Object Windows.Controls.ColumnDefinition
$rateColumn1.Width="*"
$rateColumn2=New-Object Windows.Controls.ColumnDefinition
$rateColumn2.Width="*"
$rateGrid.ColumnDefinitions.Add($rateColumn1)
$rateGrid.ColumnDefinitions.Add($rateColumn2)
$downloadValue=New-Object Windows.Controls.TextBlock
$downloadValue.Text="0 B/s"
$downloadValue.FontSize=12
$downloadValue.FontWeight="Bold"
$downloadValue.Foreground=[Windows.Media.Brushes]::DeepSkyBlue
$downloadValue.HorizontalAlignment="Center"
$downloadValue.Margin="0,4,0,0"
[Windows.Controls.Grid]::SetColumn($downloadValue,0)
$rateGrid.Children.Add($downloadValue)|Out-Null
$uploadValue=New-Object Windows.Controls.TextBlock
$uploadValue.Text="0 B/s"
$uploadValue.FontSize=12
$uploadValue.FontWeight="Bold"
$uploadValue.Foreground=[Windows.Media.Brushes]::Orange
$uploadValue.HorizontalAlignment="Center"
$uploadValue.Margin="0,4,0,0"
[Windows.Controls.Grid]::SetColumn($uploadValue,1)
$rateGrid.Children.Add($uploadValue)|Out-Null
[Windows.Controls.Grid]::SetRow($rateGrid,1)
$root.Children.Add($rateGrid)|Out-Null
$sessionGrid=New-Object Windows.Controls.Grid
$sc1=New-Object Windows.Controls.ColumnDefinition
$sc1.Width="*"
$sc2=New-Object Windows.Controls.ColumnDefinition
$sc2.Width="*"
$sessionGrid.ColumnDefinitions.Add($sc1)
$sessionGrid.ColumnDefinitions.Add($sc2)
$sessionDownload=New-Object Windows.Controls.TextBlock
$sessionDownload.Text="DOWNLOADED: 0 B"
$sessionDownload.FontSize=11
$sessionDownload.Foreground=[Windows.Media.Brushes]::LightGray
$sessionDownload.Margin="0,8,0,0"
[Windows.Controls.Grid]::SetColumn($sessionDownload,0)
$sessionGrid.Children.Add($sessionDownload)|Out-Null
$sessionUpload=New-Object Windows.Controls.TextBlock
$sessionUpload.Text="UPLOADED: 0 B"
$sessionUpload.FontSize=11
$sessionUpload.Foreground=[Windows.Media.Brushes]::LightGray
$sessionUpload.Margin="0,8,0,0"
[Windows.Controls.Grid]::SetColumn($sessionUpload,1)
$sessionGrid.Children.Add($sessionUpload)|Out-Null
[Windows.Controls.Grid]::SetRow($sessionGrid,2)
$root.Children.Add($sessionGrid)|Out-Null
$window.Content=$root
$trayIcon=New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon=[System.Drawing.SystemIcons]::Information
$trayIcon.Text="Network Monitor"
$trayMenu=New-Object System.Windows.Forms.ContextMenuStrip
$exitItem=$trayMenu.Items.Add("Exit")
$trayIcon.ContextMenuStrip=$trayMenu
$exitItem.Add_Click({$window.Close()})
$trayIcon.Visible=$true
$window.ShowInTaskbar=$false
$script:monitorVisible=$false
$trayIcon.Add_MouseClick({
    param($sender,$eventArgs)
    if($eventArgs.Button -eq [System.Windows.Forms.MouseButtons]::Left){
        $script:monitorVisible=-not $script:monitorVisible
        if($script:monitorVisible -and -not $window.IsVisible){
            $workArea=[Windows.SystemParameters]::WorkArea
            $window.Left=$workArea.Right-$window.Width-12
            $window.Top=$workArea.Bottom-$window.Height-12
            $window.Opacity=0
            $window.Show()
        }
        $visibilityTimer.Start()
    }
})
function Get-NetworkBytes{
$nic=[System.Net.NetworkInformation.NetworkInterface]::GetAllNetworkInterfaces()|Where-Object{$_.Name -eq "Ethernet"}|Select-Object -First 1
if($null -eq $nic){return [PSCustomObject]@{Received=0L;Sent=0L}}
try{$s=$nic.GetIPStatistics();return [PSCustomObject]@{Received=[long]$s.BytesReceived;Sent=[long]$s.BytesSent}}catch{return [PSCustomObject]@{Received=0L;Sent=0L}}
}
function Format-Bytes([double]$bytes){
if($bytes -ge 1GB){return "{0:N2} GB" -f ($bytes/1GB)}
if($bytes -ge 1MB){return "{0:N2} MB" -f ($bytes/1MB)}
if($bytes -ge 1KB){return "{0:N2} KB" -f ($bytes/1KB)}
return "{0:N0} B" -f $bytes
}
$historyDownload=New-Object System.Collections.Generic.List[double]
$historyUpload=New-Object System.Collections.Generic.List[double]
$first=Get-NetworkBytes
$script:previousReceived=$first.Received
$script:previousSent=$first.Sent
$script:sessionDownloadTotal=0L
$script:sessionUploadTotal=0L
$script:currentDownload=0.0
$script:currentUpload=0.0
function Add-GaugeText([string]$text,[double]$x,[double]$y,[double]$fontSize,[Windows.Media.Brush]$brush,[bool]$bold=$false){
$block=New-Object Windows.Controls.TextBlock
$block.Text=$text;$block.FontSize=$fontSize;$block.Foreground=$brush
if($bold){$block.FontWeight='Bold'}
[Windows.Controls.Canvas]::SetLeft($block,$x);[Windows.Controls.Canvas]::SetTop($block,$y)
$chartCanvas.Children.Add($block)|Out-Null
}
function Draw-Gauge([double]$centerX,[double]$centerY,[double]$radius,[double]$value,[double]$scale,[string]$label,[Windows.Media.Brush]$accent){
$background=New-Object Windows.Shapes.Ellipse
$background.Width=($radius*2);$background.Height=($radius*2)
$background.Fill=[Windows.Media.Brushes]::Black;$background.Stroke=[Windows.Media.Brushes]::DarkSlateGray;$background.StrokeThickness=1
[Windows.Controls.Canvas]::SetLeft($background,$centerX-$radius);[Windows.Controls.Canvas]::SetTop($background,$centerY-$radius)
$chartCanvas.Children.Add($background)|Out-Null
$arcPoints=New-Object Windows.Media.PointCollection
for($step=0;$step -le 48;$step++){$angle=(135+270*$step/48)*[Math]::PI/180;$arcPoints.Add([Windows.Point]::new(($centerX+$radius*0.78*[Math]::Cos($angle)),($centerY+$radius*0.78*[Math]::Sin($angle))))}
$arc=New-Object Windows.Shapes.Polyline;$arc.Points=$arcPoints;$arc.Stroke=[Windows.Media.Brushes]::DarkSlateGray;$arc.StrokeThickness=4
$chartCanvas.Children.Add($arc)|Out-Null
for($tick=0;$tick -le 10;$tick++){
$angle=(135+270*$tick/10)*[Math]::PI/180;$outer=$radius*0.78
$inner=$radius*$(if($tick%2 -eq 0){0.63}else{0.68})
$line=New-Object Windows.Shapes.Line
$line.X1=$centerX+$outer*[Math]::Cos($angle);$line.Y1=$centerY+$outer*[Math]::Sin($angle)
$line.X2=$centerX+$inner*[Math]::Cos($angle);$line.Y2=$centerY+$inner*[Math]::Sin($angle)
$line.Stroke=[Windows.Media.Brushes]::Gray;$line.StrokeThickness=1;$chartCanvas.Children.Add($line)|Out-Null
if($tick%2 -eq 0){$number=$scale*$tick/10;if($scale -ge 1GB){$tickText=('{0:0.#}G' -f ($number/1GB))}elseif($scale -ge 1MB){$tickText=('{0:0.#}M' -f ($number/1MB))}elseif($scale -ge 1KB){$tickText=('{0:0.#}K' -f ($number/1KB))}else{$tickText=('{0:0.#}' -f $number)};$labelRadius=$radius*0.53;Add-GaugeText $tickText ($centerX+$labelRadius*[Math]::Cos($angle)-9) ($centerY+$labelRadius*[Math]::Sin($angle)-8) 9 ([Windows.Media.Brushes]::WhiteSmoke) $true}
}
$ratio=[Math]::Min(1.0,[Math]::Max(0.0,$value/$scale));$needleAngle=(135+270*$ratio)*[Math]::PI/180
$needle=New-Object Windows.Shapes.Line;$needle.X1=$centerX;$needle.Y1=$centerY
$needle.X2=$centerX+$radius*0.57*[Math]::Cos($needleAngle);$needle.Y2=$centerY+$radius*0.57*[Math]::Sin($needleAngle)
$needle.Stroke=$accent;$needle.StrokeThickness=4;$chartCanvas.Children.Add($needle)|Out-Null
$hub=New-Object Windows.Shapes.Ellipse;$hub.Width=9;$hub.Height=9;$hub.Fill=$accent
[Windows.Controls.Canvas]::SetLeft($hub,$centerX-4.5);[Windows.Controls.Canvas]::SetTop($hub,$centerY-4.5);$chartCanvas.Children.Add($hub)|Out-Null
}
function Draw-Gauges{
$chartCanvas.Children.Clear();$w=$chartCanvas.ActualWidth;$h=$chartCanvas.ActualHeight
if($w -lt 60 -or $h -lt 60){return};$radius=[Math]::Min(($w/4)-8,($h/2)-5);if($radius -lt 25){return}
Draw-Gauge ($w*0.25) ($h*0.52) $radius $script:currentDownload (10MB) 'DOWNLOAD' ([Windows.Media.Brushes]::DeepSkyBlue)
Draw-Gauge ($w*0.75) ($h*0.52) $radius $script:currentUpload (5MB) 'UPLOAD' ([Windows.Media.Brushes]::Orange)
}
$timer=New-Object Windows.Threading.DispatcherTimer
$timer.Interval=[TimeSpan]::FromSeconds(1)
$visibilityTimer=New-Object Windows.Threading.DispatcherTimer
$visibilityTimer.Interval=[TimeSpan]::FromMilliseconds(30)
$visibilityTimer.Add_Tick({
    if($script:monitorVisible){
        $window.Opacity=[Math]::Min(1.0,$window.Opacity+0.06)
    }elseif($window.IsVisible){
        $window.Opacity=[Math]::Max(0.0,$window.Opacity-0.06)
        if($window.Opacity -le 0){$window.Hide()}
    }
    if(($script:monitorVisible -and $window.Opacity -ge 1) -or (-not $script:monitorVisible -and -not $window.IsVisible)){$visibilityTimer.Stop()}
})
$timer.Add_Tick({
try{
$current=Get-NetworkBytes
$download=[Math]::Max(0L,$current.Received-$script:previousReceived)
$upload=[Math]::Max(0L,$current.Sent-$script:previousSent)
$script:currentDownload=[double]$download
$script:currentUpload=[double]$upload
$script:previousReceived=$current.Received
$script:previousSent=$current.Sent
$script:sessionDownloadTotal+=$download
$script:sessionUploadTotal+=$upload
$historyDownload.Add([double]$download)
$historyUpload.Add([double]$upload)
while($historyDownload.Count -gt 10){$historyDownload.RemoveAt(0)}
while($historyUpload.Count -gt 10){$historyUpload.RemoveAt(0)}
$trayText="Down: $(Format-Bytes $download)/s | Up: $(Format-Bytes $upload)/s"
$trayIcon.Text=$trayText.Substring(0,[Math]::Min(63,$trayText.Length))
$downloadValue.Text="$(Format-Bytes $download)/s"
$uploadValue.Text="$(Format-Bytes $upload)/s"
$sessionDownload.Text="DOWNLOADED: $(Format-Bytes $script:sessionDownloadTotal)"
$sessionUpload.Text="UPLOADED: $(Format-Bytes $script:sessionUploadTotal)"
Draw-Gauges
}catch{$errorText="Monitor error: $($_.Exception.Message)";$trayIcon.Text=$errorText.Substring(0,[Math]::Min(63,$errorText.Length))}
})
$chartCanvas.Add_SizeChanged({Draw-Gauges})
$window.Add_Closed({$timer.Stop();$visibilityTimer.Stop();$trayIcon.Visible=$false;$trayIcon.Dispose();$trayMenu.Dispose()})
$timer.Start()
$window.Show()
$window.Hide()
$visibilityTimer.Start()
[Windows.Threading.Dispatcher]::Run()

