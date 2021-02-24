
$inputXML = @"
<Window x:Name="MainWindow" x:Class="PowerShellGUI.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:PowerShellGUI"
        mc:Ignorable="d"
        WindowStartupLocation="CenterScreen"
        
        Title="地铁归档助手 Ver0.5" Height="370" Width="690" ResizeMode="NoResize">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="66*"/>
            <ColumnDefinition Width="437*"/>
            <ColumnDefinition Width="113*"/>
        </Grid.ColumnDefinitions>
        <TextBlock Height="24" Margin="29,19,0,0" TextWrapping="Wrap" VerticalAlignment="Top" FontSize="15" Text="载入归档图册数据文件：" Grid.ColumnSpan="2" HorizontalAlignment="Left" Width="217"/>
        <Button x:Name="Confirm" Content="一键归档" HorizontalAlignment="Left" Height="38" Margin="13,267,0,0" VerticalAlignment="Top" Width="80" RenderTransformOrigin="0.533,1.422" Grid.Column="1"/>
        <Button x:Name="Cancel" Content="取消" HorizontalAlignment="Right" Height="38" Margin="0,267,69,0" VerticalAlignment="Top" Width="76" Grid.Column="1" Grid.ColumnSpan="2"/>
        <TextBox x:Name="TextBox_File" IsReadOnly="True" HorizontalAlignment="Left" Height="38" Margin="56,61,0,0" TextWrapping="Wrap" Text=" " VerticalAlignment="Top" Width="347" VerticalContentAlignment="Center" Grid.ColumnSpan="2" IsReadOnlyCaretVisible="True"/>
        <Button x:Name="More_File" Content="浏览文件..." HorizontalAlignment="Left" Margin="375,61,0,0" VerticalAlignment="Top" Width="87" RenderTransformOrigin="0.342,-3.105" Height="38" Grid.Column="1"/>
        <TextBlock HorizontalAlignment="Left" Height="24" Margin="29,121,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="217" FontSize="15" Grid.ColumnSpan="2"><Run Text="归档图册全名"/><LineBreak/><Run/></TextBlock>
        <TextBox x:Name="TextBox_Name" IsReadOnly="True" HorizontalAlignment="Left" Height="38" Margin="56,163,0,0" TextWrapping="Wrap" Text=" " VerticalAlignment="Top" Width="598" VerticalContentAlignment="Center" Grid.ColumnSpan="3" IsReadOnlyCaretVisible="True"/>
        <CheckBox x:Name="withCCode" Content="院归档文件增加管理号前缀" HorizontalAlignment="Left" Margin="56,232,0,0" VerticalAlignment="Top" Grid.ColumnSpan="2"/>
        <CheckBox x:Name="Plot" Content="生成出图文件" Grid.Column="1" HorizontalAlignment="Left" Margin="238,232,0,0" VerticalAlignment="Top"/>
        <CheckBox x:Name="Pdf" Content="打印PDF" Grid.Column="1" HorizontalAlignment="Left" Margin="163,232,0,0" VerticalAlignment="Top"/>
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
  
Function Get-FormVariables{
if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*
}
  
#Get-FormVariables
  
#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================
                                                                     
      
    #Reference 
  
    #Adding items to a dropdown/combo box
      #$vmpicklistView.items.Add([pscustomobject]@{'VMName'=($_).Name;Status=$_.Status;Other="Yes"})
      
    #Setting the text of a text box to the current PC name    
      #$WPFtextBox.Text = $env:COMPUTERNAME
      
    #Adding code to a button, so that when clicked, it pings a system
    # $WPFbutton.Add_Click({ Test-connection -count 1 -ComputerName $WPFtextBox.Text
    # })
Function Show-OpenFileDialog
{
  param
  (
    [string]$StartFolder = $BaseFolder,
    [string]$Title = '选择归档图册数据文件',
    [string]$Filter = '模板文件|*.xlsx'
  )

  $dialog = New-Object -TypeName Microsoft.Win32.OpenFileDialog

  $dialog.Title = $Title
  $dialog.InitialDirectory = $StartFolder
  $dialog.Filter = $Filter

  $resultat = $dialog.ShowDialog()
  if ($resultat -eq $true)
  {
    if($(IsExcelCorrect -excelpath $dialog.FileName -sheetname "图册信息表") -and $(IsExcelCorrect -excelpath $dialog.FileName -sheetname "图纸信息表"))
    {
     $WPFTextBox_File.Text = $dialog.FileName
    }
    else
    {
     Write-Host "Excel文件不正确，请检查！"
     return 
    }
  }
}

$Form.add_Loaded({

$WPFConfirm.IsEnabled = $false

})

$WPFCancel.add_Click({
$Form.Close()
exit
})

$WPFMore_File.add_Click({
Show-OpenFileDialog 
})

$WPFTextBox_File.add_TextChanged({
$WPFConfirm.IsEnabled = $true
$ExcelFile = $WPFTextBox_File.Text
$sheetset = ReadSheetSet -ExcelPath $ExcelFile -SheetName "图册信息表"
$WPFTextBox_Name.Text = $sheetset.FullName
})

$WPFConfirm.add_Click({
$allstart = Get-Date

$ExcelFile = $WPFTextBox_File.Text
Write-Host -ForegroundColor Yellow "1、开始读取图册信息..."
$sheetset = ReadSheetSet -ExcelPath $ExcelFile -SheetName "图册信息表"

$projectpath=$sheetset.图纸保存路径
$projectname=$sheetset.FullName
$pdfpath = $sheetset.PDF图纸路径_合并的文件

$SavePath = "{0}\文件归档" -f $projectpath
$SavePath_Metro ="{0}\{1}" -f $SavePath,$projectname
$SavePath_Metro_PDF ="{0}\PDF文件" -f $SavePath_Metro
$SavePath_Metro_CAD ="{0}\CAD文件" -f $SavePath_Metro
$SavePath_Base = "{0}\院归档文件" -f $SavePath

if($(CheckFolder -path $SavePath) -eq $false )
{
 Write-host -ForegroundColor Red "归档文件夹已存在，请检查后再归档！"
 return
}
Write-Host "读取图册信息完成!"

#读取图纸列表信息
Write-Host -ForegroundColor Yellow "2、开始读取图纸列表信息并校验数据..."

$sheetlist = ReadSheet -ExcelPath $ExcelFile -SheetName "图纸信息表" -colcount 8
#按图纸序号重新排序
$sheetlist = $sheetlist | Sort-Object -Property 图纸序号
#如果没有勾选打印PDF，开始校验提供的PDF文件
if($WPFPdf.IsChecked -eq $false) 
{
 if($pdfpath -eq $null)
 {
  Write-Host -ForegroundColor Red "PDF文件路径信息不存在，请检查归档数据表！"
  return
 }
 if($(Test-Path -Path $pdfpath) -eq $false)
 {
  Write-Host -ForegroundColor Red "PDF文件路径信息错误，请检查归档数据表！"
  return
 }
 $pdfpagenumber = GetPdfPageNumber -PdfPath $pdfpath
 if($pdfpagenumber -ne $sheetlist.count)
 {
  $message = "PDF文件页数与CAD图纸数不符，PDF有{0}页，CAD图纸有{1}张。" -f $pdfpagenumber,$sheetlist.count
  Write-Host -ForegroundColor Red $message
  return
 }
 }

#生成文件夹们
New-Item -ItemType "directory" -Path $SavePath,$SavePath_Metro,$SavePath_Metro_PDF,$SavePath_Metro_CAD,$SavePath_Base -Force

Write-Host "读取图纸列表信息完成，数据校验通过！"
##########下面的代码段拆分图纸布局并打印PDF（分支）###############
$splitpath = "{0}\TempSplitfiles" -f $projectpath     #临时保存分拆后的文件
###打印PDF的分支###
if($WPFPdf.IsChecked -eq $true) 
{
Write-Host -ForegroundColor Yellow "3、开始整理CAD图纸文件并打印PDF..."
$filelist = @()
$start = Get-Date
#获取DWG文件列表
foreach($sheet in $sheetlist)               
{
 $filelist += $sheet.DWG文件名
}

$filelist = $filelist | sort -Unique
$splitpath = "{0}\TempSplitfiles" -f $projectpath     #临时保存分拆后的文件
New-Item -ItemType "directory" -Path $splitpath -Force

for($i=0;$i -lt $filelist.Count;$i++)
{
 $group = $sheetlist | Where-Object{$_.DWG文件名 -eq $filelist[$i]}
 if($group.Count -eq $null)                #只有一张图（一个布局）不用分拆
 {
  $filepath = "{0}\{1}" -f $splitpath,$group.DWG文件名
  $scriptpath = "{0}\{1}.scr" -f $filepath,$group.布局名
  New-Item -ItemType "directory" -Path $filepath -Force
  $newfile = "{0}\{1}.dwg" -f $filepath,$group.布局名            #临时文件
  Copy-Item $group.DWG文件路径 -Destination $newfile
  $newPDFfile = "{0}\{1} {2}.pdf" -f $SavePath_Metro_PDF,$($group.编码.Replace("/","-")),$group.图纸名      #地铁PDF归档
  $newMfile = "{0}\{1} {2}.dwg" -f $SavePath_Metro_CAD,$($group.编码.Replace("/","-")),$group.图纸名      #地铁CAD归档
  $newCfile = ""
  if($WPFwithCCode.IsChecked -eq $true)       #院CAD归档决定是否加管理号前缀
  {
  $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$group.管理号,$group.图纸名      #加了管理号前缀
  }
  else
  {
   $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$group.图纸编号,$group.图纸名      #没有管理号前缀
  }
  $script =  @"
filedia
0
-plot
;是否需要详细打印配置？[是(Y)/否(N)] <否>:
N
;输入布局名或 [?] <>:

;输入页面设置名 <>:
$($configs.PageSetup)
;输入输出设备的名称或 [?] <>: 
$($configs.PDFPrinter)
;输入文件名<>:
"$newPDFfile"
;是否保存对页面设置的修改 [是(Y)/否(N)]? <N>
N
;是否继续打印？[是(Y)/否(N)] <Y>:
Y
saveas

"$newMfile"
saveas

"$newCfile"
filedia
1
"@
 $script | Out-File -FilePath $scriptpath -Encoding default -Force
 }
 else                                 #有多张图（多个布局），需要分拆
 {
  $grouppath = "{0}\{1}" -f $splitpath,$group[0].DWG文件名
  New-Item -ItemType "directory" -Path $grouppath -Force
    foreach($file in $group)
  {
   $newPDFfile = "{0}\{1} {2}.pdf" -f $SavePath_Metro_PDF,$($file.编码.Replace("/","-")),$file.图纸名      #地铁PDF归档
   $newMfile = "{0}\{1} {2}.dwg" -f $SavePath_Metro_CAD,$($file.编码.Replace("/","-")),$file.图纸名      #地铁CAD归档
   $newCfile = ""
   if($WPFwithCCode.IsChecked -eq $true)       #院CAD归档决定是否加管理号前缀
  {
  $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$file.管理号,$file.图纸名      #加了管理号前缀
  }
  else
  {
   $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$file.图纸编号,$file.图纸名      #没有管理号前缀
   }
   $newfile = "{0}\{1}.dwg" -f $grouppath,$file.布局名                  #拆分临时文件
   Copy-Item $file.DWG文件路径 -Destination $newfile
   $dellist = $group | Where-Object{$_.布局名 -ne $file.布局名}
   $scriptpath = "{0}\{1}.scr" -f $grouppath,$file.布局名
   #脚本，删布局的时候如果删的是当前布局会有一定几率生成个布局1（BUG？），先把保留的布局设为当前布局就好了
   $script =  @"
filedia
0
-layout
S
"$($file.布局名)"
"@
foreach($item in $dellist)
{
 $dellayout = $item.布局名
 $scradd = @"

-layout
D
"$dellayout"
"@
$script += $scradd
}
$scrend = @"

-plot
;是否需要详细打印配置？[是(Y)/否(N)] <否>:
N
;输入布局名或 [?] <>:

;输入页面设置名 <>:
$($configs.PageSetup)
;输入输出设备的名称或 [?] <>: 
$($configs.PDFPrinter)
;输入文件名<>:
"$newPDFfile"
;是否保存对页面设置的修改 [是(Y)/否(N)]? <N>
N
;是否继续打印？[是(Y)/否(N)] <Y>:
Y
saveas

"$newMfile"
saveas

"$newCfile"
filedia
1
"@

$script += $scrend
$script | Out-File -FilePath $scriptpath -Encoding default -Force
  }
 }
}

#执行脚本
$logfile = AcadScriptM -scriptname "CAD图纸拆分布局、打印PDF和整理文件" -accore $configs.accoreconsolepath -path $splitpath -session 10
$end = Get-Date
$totaltime = "{0}分{1}秒" -f $($end - $start).Minutes,$($end - $start).Seconds
Add-Content -Path $logfile -Value "脚本执行用时$totaltime"
Write-Host "完成CAD图纸文件整理和PDF文件打印!"
#生成整合的PDF文件
Write-Host -ForegroundColor Yellow "4、开始生成打印用整合的PDF文件..."
$pdffiles = Get-ChildItem -Path $SavePath_Metro_PDF *.pdf
$pdflist = New-Object System.Collections.ArrayList
foreach($file in $pdffiles)
{
 [void]$pdflist.Add($file.FullName)
}
$outpdf = "{0}\整合PDF.pdf" -f $SavePath
MergePdf -outfile $outpdf -inputfiles $pdflist
Write-Host "完成PDF文件整合！文件路径为"$outpdf""
}
###不打印PDF的分支###
else
{
Write-Host -ForegroundColor Yellow "3、开始整理CAD图纸文件..."
$filelist = @()
$start = Get-Date
#获取DWG文件列表
foreach($sheet in $sheetlist)               
{
 $filelist += $sheet.DWG文件名
}

$filelist = $filelist | sort -Unique
$splitpath = "{0}\TempSplitfiles" -f $projectpath     #临时保存分拆后的文件
New-Item -ItemType "directory" -Path $splitpath -Force

for($i=0;$i -lt $filelist.Count;$i++)
{
 $group = $sheetlist | Where-Object{$_.DWG文件名 -eq $filelist[$i]}
 if($group.Count -eq $null)                #只有一张图（一个布局）不用分拆
 {
  $newMfile = "{0}\{1} {2}.dwg" -f $SavePath_Metro_CAD,$($group.编码.Replace("/","-")),$group.图纸名      #地铁CAD归档
  $newCfile = ""
  if($WPFwithCCode.IsChecked -eq $true)       #院CAD归档决定是否加管理号前缀
  {
  $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$group.管理号,$group.图纸名      #加了管理号前缀
  }
  else
  {
   $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$group.图纸编号,$group.图纸名      #没有管理号前缀
  }
  Copy-Item $group.DWG文件路径 -Destination $newMfile
  Copy-Item $group.DWG文件路径 -Destination $newCfile
 }
 else                                 #有多张图（多个布局），需要分拆
 {
  $grouppath = "{0}\{1}" -f $splitpath,$group[0].DWG文件名
  New-Item -ItemType "directory" -Path $grouppath -Force
  foreach($file in $group)
  {
   $newMfile = "{0}\{1} {2}.dwg" -f $SavePath_Metro_CAD,$($file.编码.Replace("/","-")),$file.图纸名      #地铁CAD归档
   $newCfile = ""
   if($WPFwithCCode.IsChecked -eq $true)       #院CAD归档决定是否加管理号前缀
  {
  $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$file.管理号,$file.图纸名      #加了管理号前缀
  }
  else
  {
   $newCfile = "{0}\{1} {2}.dwg" -f $SavePath_Base,$file.图纸编号,$file.图纸名      #没有管理号前缀
   }
   $newfile = "{0}\{1}.dwg" -f $grouppath,$file.布局名                  #拆分临时文件
   Copy-Item $file.DWG文件路径 -Destination $newfile
   $dellist = $group | Where-Object{$_.布局名 -ne $file.布局名}
   $scriptpath = "{0}\{1}.scr" -f $grouppath,$file.布局名
   $script =  @"
filedia
0
-layout
S
"$($file.布局名)"
"@
foreach($item in $dellist)
{
 $dellayout = $item.布局名
 $scradd = @"

-layout
D
"$dellayout"
"@
$script += $scradd
}
$scrend = @"

saveas

"$newMfile"
saveas

"$newCfile"
filedia
1
"@

$script += $scrend
$script | Out-File -FilePath $scriptpath -Encoding default -Force
  }
 }
}

#执行脚本
$logfile = AcadScriptM -scriptname "CAD文件拆分布局和整理" -accore $configs.accoreconsolepath -path $splitpath -session 10
$end = Get-Date
$totaltime = "{0}分{1}秒" -f $($end - $start).Minutes,$($end - $start).Seconds
Add-Content -Path $logfile -Value "脚本执行用时$totaltime"
Write-Host "完成CAD文件的整理！"
#开始拆分PDF并整理
Write-Host -ForegroundColor Yellow "4、开始拆分PDF文件并整理..."

 SplitPdfFile -PdfPath $pdfpath -OutPutFolder "$BaseFolder\temp"
 #获取图名列表
 $filelist = Get-ChildItem -Path $SavePath_Metro_CAD *.dwg
 $filenames = New-Object System.Collections.ArrayList
 foreach($file in $filelist)
 {
  [void]$filenames.Add($file.BaseName)
 }
 #重命名拆分的PDF
 $pdfpagenumber = GetPdfPageNumber -PdfPath $pdfpath
 for($i=0;$i -lt $pdfpagenumber;$i++)
 {
  $j=$i+1
  $OriginName = "$BaseFolder\temp\splittedPDF$i.pdf" 
  $Name_Metro ="{0}\{1}.pdf" -f $SavePath_Metro_PDF,$filenames[$i]
  Move-Item $OriginName -Destination $Name_Metro -Force
  Write-Progress -Activity "拆分PDF文件并整理" -Status "拆分PDF文件并整理中" -CurrentOperation "$j/$pdfpagenumber" -PercentComplete ($j/$pdfpagenumber*100)  
  }
Write-Host "拆分PDF文件并整理完成！"
}


#########################

####实现出图文件功能###

#生成光盘标签
Write-Host -ForegroundColor Yellow "5、开始生成光盘标签..."
$WordFile = "{0}\Templates\disklable.docx" -f $BaseFolder
$LableSave ="{0}\光盘标签.docx"  -f $SavePath
MakeDiskLable -_ExcelFile $ExcelFile -_WordFile $WordFile -_SavePath $LableSave
Write-Host "光盘标签生成完毕！"


#生成卷内目录
Write-Host -ForegroundColor Yellow "6、开始生成卷内目录..."
 $list =@{ProjectName=$sheetset.FullName ; Owner=$sheetset.工点单位 ; Date=$sheetset.出图日期.ToString().Remove(4,1)+"01"}
 MakeSheetList -_Sheet $sheetlist -_SheetSet $sheetset -_List $list -_TempDoc "$BaseFolder\Templates\sheetlist.docx" -_SavePath "$SavePath_Metro\卷内目录.docx" -_projectnamefontsize $configs.ProjectFontsize -_mcodefontsize $configs.MFontsize -_ownerfontsize $configs.OwnerFontsize -_namefontsize $configs.NameFontsize
 Write-Host "卷内目录生成完毕！"

#生成其他出图文件
if($WPFPlot.IsChecked -eq $true)
{
 $cover = $sheetlist | Where-Object{$_.图纸名 -eq "封面"}
 $Temp_PlotA = "{0}\Templates\出图工程联系单模板A.docx" -f $BaseFolder
 $Temp_Plot0 = "{0}\Templates\出图工程联系单模板0.docx" -f $BaseFolder
 $Temp_Sub = "{0}\Templates\分包出图文件审批联络单模板.docx" -f $BaseFolder
 $Temp_Ctrl = "{0}\Templates\受控章审批单模板.docx" -f $BaseFolder
 $SavePath_Plot = "{0}\出图文件" -f $SavePath
 New-Item -ItemType "directory" -Path $SavePath_Plot
 $Year = $(Get-Date -Format %yy).toString()
 $Out_Plot = "{0}\{1}-0000 关于提交《{2}》的联系单.docx" -f $SavePath_Plot,$Year,$sheetset.FullName
 $Out_Sub = "{0}\{1}-0000 {2}-分包出图文件审批联络单.docx" -f $SavePath_Plot,$Year,$sheetset.图册全称
 $Out_Ctrl = "{0}\受控章审批单.docx" -f $SavePath_Plot
 if($sheetset.是否送审版 -eq "否")
 {
  $list_sub =@{
               ProJectName = $sheetset.工程名称
               FullName = $sheetset.FullName
               MCode = $cover.编码
               Company = $sheetset.工点单位
               
  }
  $list_ctrl =@{
                ProJectName = $sheetset.工程名称
                FullName = $sheetset.FullName
                MCode = $cover.编码
                Name = $sheetset.责任人
  }
  $list_Plot =@{
                 Version = $sheetset.是否送审版
                 ShortName = $sheetset.工程简称
                 FullName = $sheetset.FullName
                 Name = $sheetset.责任人
                 Company = $sheetset.工点单位
                 MCode = $cover.编码
  }
  SubpackageAproval -template $Temp_Sub -list $list_sub -output $Out_Sub
  ControlledAproval -template $Temp_Ctrl -list $list_ctrl -output $Out_Ctrl
  PlotRequest -templateA $Temp_PlotA -template0 $Temp_Plot0 -list $list_Plot -output $Out_Plot
  for($i=1;$i -lt 4; $i++)
  {
   Write-Progress -Activity "生成出图文件" -Status "生成出图文件中" -CurrentOperation "$i/3" -PercentComplete ($i/3*100) 
  }
 }
 elseif($sheetset.是否送审版 -eq "是")
 {
 $list_Plot =@{
                 Version = $sheetset.是否送审版
                 ShortName = $sheetset.工程简称
                 FullName = $sheetset.FullName
                 Name = $sheetset.责任人
                 Company = $sheetset.工点单位
                 MCode = $cover.编码
 }
 PlotRequest -templateA $Temp_PlotA -template0 $Temp_Plot0 -list $list_Plot -output $Out_Plot
 Write-Progress -Activity "生成出图文件" -Status "生成出图文件中" -CurrentOperation "1/1" -PercentComplete (1/1*100) -Id 5
}
else
{
 Write-Host "获取送审版信息失败！"
 return
}
}

#####清理临时文件######
if($(CheckFolder -path $splitpath) -eq $false)
{
 Remove-Item -Path $splitpath -Recurse -Force
}
Remove-Item -Path "$BaseFolder\temp\*.*" -Force


#######################


$allend = Get-Date
$totaltime = "{0}分{1}秒" -f $($allend - $allstart).Minutes,$($allend - $allstart).Seconds
Write-Host -ForegroundColor Yellow "完成归档文件整理！用时用时$totaltime."
Invoke-Item $SavePath
})










    #===========================================================================
    # Shows the form
    #===========================================================================
#write-host "To show the form, run the following" -ForegroundColor Cyan
$Form.ShowDialog() | out-null
  
  
  
# SIG # Begin signature block
# MIIFlwYJKoZIhvcNAQcCoIIFiDCCBYQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUh+JswwWjrbZbKpsjtMFLrDGV
# AeOgggMtMIIDKTCCAhWgAwIBAgIQRRn0wX7drrlAomMxELaoYjAJBgUrDgMCHQUA
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
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU/wb7tIuxJ9Gk7oGL
# jzZ8y6EXzBgwDQYJKoZIhvcNAQEBBQAEggEARdHk0fQyDjBizzZa9Bggx30OAeSE
# CqI+56mQ1wlD7Pf9CAnUSxIJ3tVL+7z8eRVyehBK/tCHPpQvu3zD2TIKNnUyEC05
# vGsuBLBs8gWT0v5GNYSJEVr3A4TxehIVBopi6+2DEX4zoKlft/v2e2eCu9Y0zOtH
# TrkrVzEN5myMDbkXiScKayVULCQQlftxcQJ+yJtpRcsDbL5KW7z3R5B48xXeGbbR
# x7QewGTBFQI52fvm++k7j7VI1l4kocrWb6TesAWGW8cjDFzJTqPoRsLAcoE+lceg
# K4X0bauFgBq5mWJrOtdK3+654kFI4+owR7q8wCggTHhQq6SQFaqCHdrauA==
# SIG # End signature block
