#Получаем время в нужном формате
$year=(get-date).year
$datetime = Get-Date -Format "yyyy-MM-dd HH-mm-ss";
# Путь для хранения отчетов
$reportPath = "C:\DiskSpace_Report\";

#Имя файла лога
$logname="DiskSpaceRpt-$datetime.log";
$logreport=$reportpath + $logname

# Имя файла отчета
$reportName = "DiskSpaceRpt-$datetime.html";
$diskReport = $reportPath + $reportName
function Write-Log {
    param(
    [parameter(Mandatory=$true)]
    [string]$Text,
    [parameter(Mandatory=$true)]
    [ValidateSet("WARNING","ERROR","INFO")]
    [string]$Type
    )

    [string]$logMessage = [System.String]::Format("[$(Get-Date)] -"),$Type, $Text
    Add-Content -Path $logreport -Value $logMessage
}

#проверка наличия папки
if(!(test-path -path $reportPath)) {
    
    New-item -ItemType Directory -path $Reportpath -Force -ErrorAction Stop
    write-host "Directory succesfully created"
    Write-Log  -Text "Directory succesfully created" -Type INFO
}
else {
    write-host "Directory is already exist"
    Write-Log -Text "Directory is already exist" -type INFO
}

#Блок проверки и создания вложенности папок \папка хранения отчета\год\месяц
if(!(dir -directory "$reportPath\$year"))
{
    New-Item -Path "$reportpath\$year" -ItemType Directory -Force
    [datetime]$NewYear = "1/1"
    $MonthCount = 1
    while($MonthCount -le 12)
    {
        $MonthCount_MM = ("{0:D2}" -f $MonthCount).ToString()
        $Month_MMMM = $NewYear.AddMonths($MonthCount-1).ToString("MMMM")
        $MonthCount_MM = $MonthCount_MM + ' ' + $Month_MMMM
        New-Item -Path "$reportPath\$Year\$MonthCount_MM" -ItemType Directory -Force
        $MonthCount++
    }
}
#get-childitem C:\DiskSpace_Report\* -include *.html | remove-item -Recurse -force
#get-childitem C:\DiskSpace_Report\* -include *.log | remove-item -Recurse -force
# Параметры предупреждений в %
$all = 101
$percentWarning = 30
$percentCritcal = 15



$redColor = "#FF4500"
$orangeColor = "#FBB917"
$whiteColor = "#CCCCCC"

$i = 0;

# Список компьютеров для отчета
#$computers = 'a-sec-ksc-01', 'A-SEC-MSCA', 'A-SEC-TMS', 'B-AvCA', 'B-AvsubCA', 'secure', 'a-sec-ksc';
$computers = (get-adcomputer -filter {operatingsystem -like '*server*'}).name;

#E-mail settings
$SMTPServer = "10.70.2.222"
#$SMTPPort = "25"
$Username = "DiskMonitor@gis.by"
$to = "it-group@gis.by"

$subject = "Servers Disks Space monitoring"
$doctype ='<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" 
"http://www.w3.org/TR/html4/strict.dtd">'
Add-Content $diskReport $doctype
$titleDate = Get-Date -Format "yyyy-MM-dd HH-mm-ss"
$header = "
<html>
<head>
<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1'>
<title>DiskSpace Report</title>
<STYLE TYPE='text/css'>
td {
    font-family: Tahoma;
    font-size: 13px;
    border-top: 1px solid #999999;
    border-right: 1px solid #999999;
    border-bottom: 1px solid #999999;
    border-left: 1px solid #999999;
    padding-top: 0px;
    padding-right: 0px;
    padding-bottom: 0px;
    padding-left: 0px;
    }
    body {
    margin-left: 5px;
    margin-top: 5px;
    margin-right: 0px;
    margin-bottom: 10px;
    table {
    border: thin solid #000000;
    }
</style>
</head>
<body>
<table width='50%'>
<tr bgcolor='#CCCCCC'>
<td colspan='7' height='25' align='center'>
<font face='tahoma' color='#003399' size='4'><strong>Servers DiskSpace Report for $titledate</strong></font>
</td>
</tr>
</table>
"

Add-Content $diskReport $header

$tableHeader = "
<table width='50%'><tbody>
<tr bgcolor=#CCCCCC>
<td width='10%' align='center'>Server</td>
<td width='5%' align='center'>Drive</td>
<td width='10%' align='center'>Freespace %</td>
<td width='10%' align='center'>Free Space</td>
<td width='10%' align='center'>Total Capacity</td>
<td width='10%' align='center'>Used Capacity</td>
</tr>
"

Add-Content $diskReport $tableHeader

foreach($computer in $Computers)
{
if (test-connection $computer -count 1 -Quiet){
    Write-Log -Text "Collecting data from $computer.gis.corp" -Type INFO
$disks =Get-WmiObject -ComputerName $Computer -Class Win32_LogicalDisk -Filter "DriveType=3"
$computer = $computer.toupper()
foreach($disk in $disks)
{
$deviceID = $disk.DeviceID;
#$volName = $disk.VolumeName;
[float]$size = $disk.Size;
[float]$freespace = $disk.FreeSpace;
$percentFree = [Math]::Round(($freespace / $size) * 100, 2);
$sizeGB = [Math]::Round($size / 1073741824, 2);
$sizeTB = [Math]::Round($size/1tb, 2);
$freeSpaceGB = [Math]::Round($freespace / 1073741824, 2);
$freeSpaceTB = [Math]::Round($freespace/1tb, 2)
$usedSpaceGB = [Math]::Round($sizeGB - $freeSpaceGB, 2);
$usedSpaceTB = [Math]::Round($sizeTB - $freeSpaceTB, 2);
$color = $whiteColor;
#блок условий выбора цвета
if($percentFree -lt $all)
{
$color = $whiteColor

if($percentFree -lt $percentCritcal)
{
$color = $redColor
}
if ($percentFree -lt $percentWarning -and $percentFree -gt $percentCritcal)
{
    $color= $orangeColor
}
#блок условий дописи GB/TB
 if ($sizeGB -gt 1000)
  {
    $sizeGB="$sizeTB TB"
  } else {
      $sizeGB= "$sizeGB GB"
      }
  if($freeSpacegb -gt 1000)
  {
      $freeSpacegb="$freeSpaceTB TB"
  } else {
      $freeSpacegb= "$freeSpaceGB GB"
  }
  if ($usedSpaceGB -gt 1000)
  {
      $usedSpaceGB="$usedSpaceTB TB"
  } else {
      $usedSpaceGB="$usedSpaceGB GB"
  }

$dataRow = "
<tr>
<td width='15%' bgcolor=`'$color`'>$computer</td>
<td width='5%' align='center' bgcolor=`'$color`'>$deviceID</td>
<td width='5%' bgcolor=`'$color`' align='center'>$percentFree</td>
<td width='15%' align='center' bgcolor=`'$color`'>$freeSpaceGB</td>
<td width='10%' align='center' bgcolor=`'$color`'>$sizeGB</td>
<td width='15%' align='center' bgcolor=`'$color`'>$usedSpaceGB</td>
</tr>
"
Add-Content $diskReport $dataRow;
Write-Host -ForegroundColor DarkYellow "$computer $deviceID,percentage,free_space = $percentFree";
$i++
}
}
}
else {
    $OfflineRow= "
    <tr>
    <td width='15%' bgcolor=#CCCCCC>$computer</td>
    <td width='5%' align='center' bgcolor=#CCCCCC>-offline-</td>
    <td width='5%' align='center' bgcolor=#CCCCCC>-offline-</td>
    <td width='15%' align='center' bgcolor=#CCCCCC>-offline-</td>
    <td width='10%' align='center' bgcolor=#CCCCCC>-offline-</td>
    <td width='15%' align='center' bgcolor=#CCCCCC>-offline-</td>
    </tr>
    "
    Add-Content $diskReport $OfflineRow;
    Write-Log -Text "$computer.gis.corp is offline" -Type ERROR
}
}

$tableDescription = "
</table><br><table width='20%'>
<tr bgcolor='White'>
<td width='10%' align='center' bgcolor=`'$whitecolor`'>Normal disk space</td>
<td width='10%' align='center' bgcolor=`'$orangecolor`'>Warning less than 30% free space</td>
<td width='10%' align='center' bgcolor=`'$redcolor`'>Critical less than 15% free space</td>
</tr>
"

Add-Content $diskReport $tableDescription
Add-Content $diskReport "</body></html>"



if ($i -gt 0)
{
    Write-Host "Sending email notification"
    Write-Log -Text "Sending Email notification" -Type INFO

#$file = "$reportPath\DiskSpaceRpt-$datetime.html"
#$logfile ="$reportPath\DiskSpaceRpt-$datetime.log"





$bodyreport = Get-Content "$diskreport" -Raw
Send-MailMessage -to $to -Subject $subject -From $Username -Attachments $logreport -BodyAsHtml $bodyreport -SmtpServer $SMTPServer
<#$message = New-Object System.Net.Mail.MailMessage
$message.subject = $subject
$message.body = $html
$message.to.add($to)
$message.from = $username 
$message.attachments.add($file)

$smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort);
$smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password);
$smtp.send($message)#>
}
$FileDir = dir -directory "$reportpath\$Year"
foreach ($dir in $filedir)
{
 if ($dir -like "*$MonthCount_MM*") {
    Move-Item -path "C:\DiskSpace_Report\*html" -Destination "$reportpath\$Year\$dir" -Force
    Move-Item -path "C:\DiskSpace_Report\*log" -Destination "$reportpath\$Year\$dir" -Force
 }
}
