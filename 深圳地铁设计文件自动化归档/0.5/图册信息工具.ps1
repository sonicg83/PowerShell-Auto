    #ERASE ALL THIS AND PUT XAML BELOW between the @" "@ 
$inputXML = @"
<Window x:Name="MainWindow" x:Class="PowerShellGUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShellGUI"
        mc:Ignorable="d"
        WindowStartupLocation="CenterScreen"
        
        Title="读取图册信息工具 Ver0.5" Height="321.385" Width="690" ResizeMode="NoResize">
    <Grid Margin="0,0,0,18">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="73*"/>
            <ColumnDefinition Width="33*"/>
            <ColumnDefinition Width="452*"/>
            <ColumnDefinition Width="17*"/>
            <ColumnDefinition Width="109*"/>
        </Grid.ColumnDefinitions>
        <TextBlock Height="24" Margin="29,19,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="12" Text="待归档文件所在文件夹：" Grid.ColumnSpan="3" HorizontalAlignment="Left" Width="217"/>
        <Button x:Name="Confirm" Content="生成数据表" HorizontalAlignment="Left" Height="38" Margin="309,225,0,0" VerticalAlignment="Top" Width="80" RenderTransformOrigin="0.533,1.422" Grid.Column="2"/>
        <Button x:Name="Cancel" Content="取消" HorizontalAlignment="Right" Height="38" Margin="0,225,68,0" VerticalAlignment="Top" Width="76" Grid.Column="2" Grid.ColumnSpan="3"/>
        <TextBox x:Name="TextBox_Folder" IsReadOnly="True" HorizontalAlignment="Left" Height="38" Margin="69,49,0,0" TextWrapping="Wrap" Text=" " VerticalAlignment="Top" Width="347" VerticalContentAlignment="Center" Grid.ColumnSpan="3" IsReadOnlyCaretVisible="True"/>
        <Button x:Name="Find_Folder" Content="浏览文件夹..." HorizontalAlignment="Left" Margin="354.762,49,0,0" VerticalAlignment="Top" Width="87" RenderTransformOrigin="0.342,-3.105" Height="38" Grid.Column="2"/>
        <TextBox x:Name="TextBox_Cover" IsReadOnly="True" HorizontalAlignment="Left" Height="38" Margin="69,135,0,0" TextWrapping="Wrap" Text=" " VerticalAlignment="Top" Width="347" VerticalContentAlignment="Center" Grid.ColumnSpan="3" IsReadOnlyCaretVisible="True"/>
        <TextBlock Height="24" Margin="29,106,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="12" Text="载入封面文件：" Grid.ColumnSpan="3" HorizontalAlignment="Left" Width="217"/>
        <Button x:Name="Find_Cover" Content="浏览文件..." HorizontalAlignment="Left" Margin="354.762,134,0,0" VerticalAlignment="Top" Width="87" RenderTransformOrigin="0.342,-3.105" Height="38" Grid.Column="2"/>
        
    </Grid>
</Window> 
  
"@ 
  
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
  
    $reader=(New-Object System.Xml.XmlNodeReader $xaml) 
  try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch [System.Management.Automation.MethodInvocationException] {
    Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
    write-host $error[0].Exception.Message -ForegroundColor Red
    if ($error[0].Exception.Message -like "*button*"){
        write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"}
}
catch{#if it broke some other way <span class="wp-smiley wp-emoji wp-emoji-bigsmile" title=":D">:D</span>
    Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
        }
  
#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================
  
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
  
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
  
#Get-FormVariables
  
#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================

$BaseFolder = ""
if ($MyInvocation.MyCommand.CommandType -eq "ExternalScript")
{ 
   $BaseFolder = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition 
}
else
{ 
   $BaseFolder = Split-Path -Parent -Path ([Environment]::GetCommandLineArgs()[0]) 
   if (!$ScriptPath){ $ScriptPath = "." } 
}

."$BaseFolder\Functions.ps1"
$configfile = "$BaseFolder\config.ini"
$configs = ReadIni -inipath $configfile

Function Show-OpenFileDialog
{
  param
  (
    [string]$StartFolder = "C:\",
    [string]$Title,
    [string]$Filter = "CAD文件|*.dwg"
  )

  $dialog = New-Object -TypeName Microsoft.Win32.OpenFileDialog

  $dialog.Title = $Title
  $dialog.InitialDirectory = $StartFolder
  $dialog.Filter = $Filter

  $result = $dialog.ShowDialog()
  if ($result -eq $true)
  {
    return $dialog.FileName
  }
  else
  {
   return $null
  }
}

Function Show-OpenFolderDialog
{
  param
  ( 
    #[string]$StartFolder = $defaultpath,
    [string]$Title = "请选择归档项目图纸所在的文件夹"
  )

 $dialog = New-Object -TypeName System.Windows.Forms.FolderBrowserDialog
 #$dialog.SelectedPath = $StartFolder
 $dialog.Description=$Title
  $result = $dialog.ShowDialog()
  if ($result -eq $true)
  {
    return $dialog.SelectedPath
  }
  else
  {
   return $null
  }
}


$Form.add_Loaded({

$WPFConfirm.IsEnabled = $false

})

$WPFCancel.add_Click({
$Form.Close()
exit
})

$WPFFind_Folder.add_Click({

 $text = Show-OpenFolderDialog 
 if($text -ne $null)
 {
  $WPFTextBox_Folder.Text = $text
 }

})

$WPFFind_Cover.add_Click({
if($WPFTextBox_Folder.Text -ne " ")
{
 $text = Show-OpenFileDialog -StartFolder $WPFTextBox_Folder.Text -Title "选择封面文件"
 if($text -ne $null)
 {
 $WPFTextBox_Cover.Text = $text
 }
}
else
{
 $text = Show-OpenFileDialog -Title "选择封面文件"
 if($text -ne $null)
 {
 $WPFTextBox_Cover.Text = $text
 }
}

})


$WPFTextBox_Folder.add_TextChanged({
if($WPFTextBox_Cover.Text -ne " ")
{
 $WPFConfirm.IsEnabled = $true
}

}) 

$WPFTextBox_Cover.add_TextChanged({
if($WPFTextBox_Folder.Text -ne " ")
{
 $WPFConfirm.IsEnabled = $true
}

})


$WPFConfirm.add_Click({

$allstart = Get-Date
$workfolder = $WPFTextBox_Folder.Text
$coverfile = $WPFTextBox_Cover.Text
$TempScrPath_Sheetset = "$BaseFolder\temp\tempscr_sheetset.scr"
$TempScrPath_Sheet = "$BaseFolder\temp\tempscr_sheet.scr"
$TempSplitfiles = "{0}\TempSplitfiles" -f $workfolder
$acccore = $configs.accoreconsolepath
$AcadAddin_GetLayouts = "$BaseFolder\Libs\GetLayouts.dll"
$AcadAddin_GetSheetset = "$BaseFolder\Libs\GetSheetset.dll"
$templet = "$BaseFolder\Templates\数据模板.xlsx"
$csvpath = "$BaseFolder\temp\tempcsv.csv"

$tempsheetsetcsv = "{0}\{1}SheetSet.csv" -f $env:TEMP,$([System.IO.Path]::GetFileNameWithoutExtension($coverfile))


########读取图册信息########
Write-Host -ForegroundColor Yellow "开始读取图册信息..."
$start = Get-Date
$script_sheetset=@"
filedia
0
secureload
0
cmdecho
0
netload
"$AcadAddin_GetSheetset"
GetSheetset
filedia
1
secureload
1
cmdecho
1
"@
$templog = "{0}\templogfile_sheetset.txt" -f $env:TEMP
$logfile = "{0}\log\获取图册信息脚本执行记录{1}.log" -f $PSScriptRoot,$(Get-Date -Format "yyyy-MM-dd HH-mm-ss")
Set-Content $TempScrPath_Sheetset -Value $script_sheetset -Encoding Default -Force
$process = Start-Process -FilePath $acccore -ArgumentList "/i ""$coverfile"" /s $TempScrPath_Sheetset /l zh-CN /isolate"  -WindowStyle Hidden -RedirectStandardOutput $templog -PassThru
$process.WaitForExit()
$log = TransCode -filepath $templog
Add-Content -Path $logfile -Value $log -Encoding UTF8
$end = Get-Date
$totaltime = "{0}分{1}秒" -f $($end - $start).Minutes,$($end - $start).Seconds
Add-Content -Path $logfile -Value "脚本执行用时$totaltime"
$sheetset = Import-Csv -Path $tempsheetsetcsv -Encoding Default
Add-Member -InputObject $sheetset -MemberType NoteProperty -Name "workpath" -Value $workfolder -Force
Add-Member -InputObject $sheetset -MemberType NoteProperty -Name "PDFfile" -Value "" -Force
$pdf = Get-ChildItem -Path $workfolder *.pdf
if($pdf.count -eq 1)
{
 $sheetset.PDFfile = $pdf.FullName
}
Write-Host -ForegroundColor Yellow "读取图册信息完成！用时$totaltime"

##############################################

###############读取图纸信息###################
Write-Host -ForegroundColor Yellow "开始读取图纸信息..."
$start = Get-Date
$maplist = @()
$map = [ordered]@{
         图纸序号 = "";
         图纸编号 = "";
         编码 = "";
         管理号 = "";
         图纸名 = "";
         布局名 = "";
         DWG文件名 = "";
         DWG文件路径 = ""
         #拆分文件路径 =""
}

$DwgFiles= Get-ChildItem -Path $workfolder *.dwg
$script =@"
filedia
0
secureload
0
cmdecho
0
netload
"$AcadAddin_GetLayouts"
GetLayouts
filedia
1
secureload
1
cmdecho
1
"@

$logfile = AcadScript -script $script -scriptpath $TempScrPath_Sheet -scriptname "读取图纸信息" -accore $acccore -path $workfolder -session $configs.session
$end = Get-Date
$totaltime = "{0}分{1}秒" -f $($end - $start).Minutes,$($end - $start).Seconds
Add-Content -Path $logfile -Value "脚本执行用时$totaltime"
Write-Host -ForegroundColor Yellow "读取图纸信息完成！用时$totaltime"

##################################################################

############创建数据表并写入图纸信息和图册信息####################

foreach($item in $DwgFiles)
{
 if($item.FullName -eq $coverfile)
 {
  $serialnumber = $sheetset.serial + "0000"
  $map.图纸序号 = $serialnumber
  $map.图纸编号 = "{0}-{1}" -f $sheetset.Zhuanye,$serialnumber
  $map.编码 = "{0}{1}/{2}" -f $sheetset.Mcode,$serialnumber,$sheetset.Version
  $map.管理号 = "{0}{1}{2}" -f $sheetset.Ccode,$serialnumber,$sheetset.Version
  $map.图纸名 = "封面"
  $map.布局名 = "封面的布局名无所谓啦"
  $map.DWG文件名 = [System.IO.Path]::GetFileNameWithoutExtension($coverfile)
  $map.DWG文件路径 = $coverfile
  $maplist +=[PSCustomObject]$map
 }
 else
 {
 $path = "{0}\{1}temphandle.csv" -f $env:TEMP,$item.BaseName
 $layouts = Import-Csv -Path $path -Encoding Default
 foreach($layout in $layouts)
 {
  #$splitfile = "{0}\{1}\{2}.dwg" -f $TempSplitfiles,$item.BaseName,$layout.LayoutName
  $serialnumber = $sheetset.serial + $layout.Serial.PadLeft(4,"0")
  $map.图纸序号 = $serialnumber
  $map.图纸编号 = "{0}-{1}" -f $sheetset.Zhuanye,$serialnumber
  $map.编码 = "{0}{1}/{2}" -f $sheetset.Mcode,$serialnumber,$sheetset.Version
  $map.管理号 = "{0}{1}{2}" -f $sheetset.Ccode,$serialnumber,$sheetset.Version
  $map.图纸名 = $layout.SheetName
  $map.布局名 = $layout.LayoutName
  $map.DWG文件名 = $item.BaseName
  $map.DWG文件路径 = $item.FullName
  #$map.拆分文件路径 = $splitfile
  $maplist +=[PSCustomObject]$map
 }
 }
 
}
$maplist = $maplist | Sort-Object -Property 图纸序号
$excelpath = "$BaseFolder\{0}-归档数据表.xlsx" -f $sheetset.SheetSetName
Out-Excel -sheetsetsource $sheetset -sheetsource $maplist -Template $templet -sheetname "图纸信息表" -sheetsetname "图册信息表" -tempcsvpath $csvpath -outpath $excelpath

#######################################################

$allend = Get-Date
$totaltime = "{0}分{1}秒" -f $($allend - $allstart).Minutes,$($allend - $allstart).Seconds
Write-Host -ForegroundColor Yellow "归档数据表生成成功！用时$totaltime"
Invoke-Item -Path $excelpath
})                                                                    
      
    #Reference 
  
    #Adding items to a dropdown/combo box
      #$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})
      
    #Setting the text of a text box to the current PC name    
      #$WPFtextBox.Text = $env:COMPUTERNAME
      
    #Adding code to a button, so that when clicked, it pings a system
    # $WPFbutton.Add_Click({ Test-connection -count 1 -ComputerName $WPFtextBox.Text
    # })
    #===========================================================================
    # Shows the form
    #===========================================================================
#write-host "To show the form, run the following" -ForegroundColor Cyan
$Form.ShowDialog() | out-null
  
  
  
# SIG # Begin signature block
# MIIFlwYJKoZIhvcNAQcCoIIFiDCCBYQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUNWiWwPNmzmkMCzg0zJ9EthRD
# 7n2gggMtMIIDKTCCAhWgAwIBAgIQRRn0wX7drrlAomMxELaoYjAJBgUrDgMCHQUA
# MB8xHTAbBgNVBAMTFFNvbmljZ1Bvd2VyU2hlbGxDZXJ0MB4XDTIwMDQyNjA2NTYy
# NloXDTM5MTIzMTIzNTk1OVowHzEdMBsGA1UEAxMUU29uaWNnUG93ZXJTaGVsbENl
# cnQwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDDIhX2jriAY4ySZStR
# N/BmXVarjlyWWXpgFLMVyxCSBual0y9w4DvdCaHUQDjj6EbFUCS7Yi0pmeqT8FQu
# ngD4rxTd7N2MBlL/rq7lYgPE4OT27bTVfkHo5vFCRDySEv/YfODoaYwNEzlR7I5e
# jS5CzEv8NOD2ERQ+h0N0edtndeSnYBMgsTPC67ng5TUNnmu3zKKGCS43gt13smvk
# pHd0Ia4EI3YEwXkkl2dkqaGZiCvSHAmd2XYPPE7BJXeJdj6CCLjGRYx0gI83Tp9Q
# g7UnGLcn+zes8W/YhyXRwSCx6FfThLVLfggWWzoBaNPS6ZK5H/xyBGlG4oGT/Dpd
# pa7NAgMBAAGjaTBnMBMGA1UdJQQMMAoGCCsGAQUFBwMDMFAGA1UdAQRJMEeAEOz1
# GxmaH9BdLSHFqM0r0hWhITAfMR0wGwYDVQQDExRTb25pY2dQb3dlclNoZWxsQ2Vy
# dIIQRRn0wX7drrlAomMxELaoYjAJBgUrDgMCHQUAA4IBAQA+d7ZYvFjKOUm2nQmp
# HO4JpZlEDMFxEejGOWJih/HbSDl+W8fQ2Qc/jt0mdOYyw6/8Dbrv8Ovb0+ilX6UO
# dhC5brhkCIZeGElg1SQeEYR9oj3Yq7ZJk+r+9cyNK9OPc9llm+/aMDv4j4H+rDIP
# LJ4WEIT628NmG8XxU7dBm53neCshqAzD07iChOKwYJ43BWr9VdmfSeWZdXBkZUNl
# 2Cn9QJ6pEcLTlT+H/b7nkRkkh+lp1ecfm7ZQDyD3bpYNvPa6wkh4uvIQMnqKM1a6
# O6LKwCSoiEtVPgYMD52a/eFG7mrYOChJoAtUlU4Hg5a/5W0UKaHOLoB/qHzuAdVh
# UUgDMYIB1DCCAdACAQEwMzAfMR0wGwYDVQQDExRTb25pY2dQb3dlclNoZWxsQ2Vy
# dAIQRRn0wX7drrlAomMxELaoYjAJBgUrDgMCGgUAoHgwGAYKKwYBBAGCNwIBDDEK
# MAigAoAAoQKAADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUhjf+MUH4/45pyfvq
# wKy4l/OEeukwDQYJKoZIhvcNAQEBBQAEggEAHlW+mN3OUGrJT73rY+Mm/UXYzEK0
# dm0b0YNcLU+nwJLSQ3T+2/tl0G/zMOjSNoH8ABxh4bRVKmMIpCPJuuEECV2El4zL
# F4D1Aoyvq6vdWb7YL9xLwClKRfxL/aNsXl0EyBPC+j4/mWkEeOWNZ+308ZwR8oQt
# QBo/3iGLizrf3Wob+J38bcGl4CPybMWmAuIUZlf3FFls9z0dU2BFGK6gOvqufGQe
# AynFEXmqhiJ31LvgTn80I9MmL7l/wQVaaGd7Xw3QEMiFxjexdyhpxGMAbEGcIZOM
# KbG8i5hzeeRSEJMX6LRxllynSCkcGahsxoJi7bweT7bBi/4f4Jc7QnYi6w==
# SIG # End signature block
