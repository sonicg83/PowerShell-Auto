##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
####加载依赖项####
#加载Docx（Word相关）
Add-Type -Path "$PSScriptRoot\Libs\DocX\Xceed.Document.NET.dll"
Add-Type -Path "$PSScriptRoot\Libs\DocX\Xceed.Words.NET.dll"
#加载EPPlus（Excel相关）
Add-Type -Path "$PSScriptRoot\Libs\EPPlus\EPPlus.dll"
#加载iText（PDF相关）
Add-Type -Path "$PSScriptRoot\Libs\iText\PDFSplitter.dll"
Add-Type -Path "$PSScriptRoot\Libs\iText\itext.kernel.dll"

####全局变量####





##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function FileToGroup
{
 param
 (
  [string]$folder,                 #文件夹路径
  [string]$extension = "*",        #文件后缀名，默认为所有文件
  [int32]$filenumber,              #每组文件数量
  [System.Collections.ArrayList]$TotalList         #保存分组信息的数组
 )
$path = New-Object System.IO.DirectoryInfo($folder)
  if($path -eq $false)
  {
   Write-Host "文件夹不存在，退出脚本..."
   exit
  }
  $files = Get-ChildItem -Path $folder -File "*.$extension"
  if($files -eq $null)
   {
     if($extension -eq "*")
     {
      Write-Host "该路径下未发现任何文件，退出脚本..."
     }
     else
     {
      Write-Host "该路径下未发现任何后缀为'$extension'的文件，退出脚本..."
     }
     exit
   }

$FileCount = $files.count    
$GroupCount = [math]::Ceiling($files.count/$filenumber)          #根据文件数量计算分组数
for($i=0;$i -lt $GroupCount;$i++)
{
 $templist = New-Object System.Collections.ArrayList
 #最后一组的情况
 if($i -eq $($GroupCount-1))
 {
  for($j=0;$j -lt $($files.count - $filenumber*$i);$j++)
  {
   [void]$templist.Add($files[$($filenumber*$i+$j)])
  }
  [void]$TotalList.Add($templist)
 }
 #其他组的情况
 else
 {
  for($j=0;$j -lt $filenumber;$j++)
  {
   [void]$templist.Add($files[$($filenumber*$i+$j)])
  }
  [void]$TotalList.Add($templist)
 }
}  
return $FileCount
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function FileToGroupM
{
 param
 (
  [string]$folder,                 #文件夹路径
  [string]$extension = "*",        #文件后缀名，默认为所有文件
  [int32]$filenumber,              #每组文件数量
  [System.Collections.ArrayList]$TotalList         #保存分组信息的数组
 )
$path = New-Object System.IO.DirectoryInfo($folder)
  if($path -eq $false)
  {
   Write-Host "文件夹不存在，退出脚本..."
   exit
  }
  $files = Get-ChildItem -Path $folder -File "*.$extension" -Recurse
  if($files -eq $null)
   {
     if($extension -eq "*")
     {
      Write-Host "该路径下未发现任何文件，退出脚本..."
     }
     else
     {
      Write-Host "该路径下未发现任何后缀为'$extension'的文件，退出脚本..."
     }
     exit
   }

$FileCount = $files.count    
$GroupCount = [math]::Ceiling($files.count/$filenumber)          #根据文件数量计算分组数
for($i=0;$i -lt $GroupCount;$i++)
{
 $templist = New-Object System.Collections.ArrayList
 #最后一组的情况
 if($i -eq $($GroupCount-1))
 {
  for($j=0;$j -lt $($files.count - $filenumber*$i);$j++)
  {
   [void]$templist.Add($files[$($filenumber*$i+$j)])
  }
  [void]$TotalList.Add($templist)
 }
 #其他组的情况
 else
 {
  for($j=0;$j -lt $filenumber;$j++)
  {
   [void]$templist.Add($files[$($filenumber*$i+$j)])
  }
  [void]$TotalList.Add($templist)
 }
}  
return $FileCount
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function ReadIni
{
 param
 (
  [string]$inipath
 )
 $result = @{}
 $payload = Get-Content -Path $inipath -Encoding Default |
 Where-Object {$_ -like '*=*'} |
 ForEach-Object {
 $infos = $_ -split '='
 $key = $infos[0].trim()
 $value = $($infos[1].trim()).replace('"',"")
 $result.$key = $value
 }
 return [PSCustomObject]$result
 }

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function Transcode
{
 param
 (
  [string]$filepath
 )
  $content = Get-Content -Path $filepath -Encoding Unicode | where {$_ -ne ""} #去除空白行
  return $content.Trim()
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function AcadScript
{
 param
 (
  [string]$script,
  [string]$scriptname,
  [string]$scriptpath="$env:TEMP\tempscr.scr",
  [string]$accore,
  [string]$path,
  [int32]$session = 10
 )
 #验证进程数
 $patten = "^\+?[1-9][0-9]*$"
 if($($session -match $patten) -eq $False)
 {
  Write-Host "同时运行脚本数必须为非零的正整数，请检查config.ini，退出脚本..."
  exit
 }
 #文件分组
 $TotalList = New-Object System.Collections.ArrayList
 $FilesCount = FileToGroup -folder $path -extension "dwg" -filenumber $session -TotalList $TotalList

 #分组多进程运行脚本
 Set-Content $scriptpath -Value $script -Encoding Default -Force 
 $GroupCount = $TotalList.Count
 $logfile = "{0}\log\{1}脚本执行记录{2}.log" -f $PSScriptRoot,$scriptname,$(Get-Date -Format "yyyy-MM-dd HH-mm-ss")
 $logheader = @"
>> $logfile <<
共处理 $FilesCount 个文件。
**************************************************************************************************
"@
Set-Content -Path $logfile -Value $logheader -Encoding UTF8
 $Counter = 1
 for($i=0;$i -lt $GroupCount;$i++)
 {
  $SessionCount = $TotalList[$i].Count
  $ProcessList = New-Object System.Collections.ArrayList
  $TemplogList = New-Object System.Collections.ArrayList
  for($j=0;$j -lt $SessionCount;$j++)
  {
   $templog = "{0}\templogfile{1}_{2}.txt" -f $env:TEMP,$i,$j
   $process = Start-Process -FilePath $accore -ArgumentList "/i ""$($TotalList[$i][$j].FullName)"" /s $scriptpath /l zh-CN /isolate"  -WindowStyle Hidden -RedirectStandardOutput $templog -PassThru
   [void]$ProcessList.Add($process)
   [void]$TemplogList.Add($templog)
  }
  for($j=0;$j -lt $ProcessList.Count;$j++)
{
 if($ProcessList[$j].HasExited -eq $False)
 {
  $ProcessList[$j].WaitForExit()
 }
 $addcontent = Transcode -filepath $TemplogList[$j]
 Add-Content -Path $logfile -Value ">>$($TotalList[$i][$j].FullName)<<" -Encoding UTF8
 Add-Content -Path $logfile -Value $addcontent -Encoding UTF8
 if($Counter -eq $FilesCount)
 {
  Add-Content -Path $logfile -Value "**************************************************************************************************"
 }
 else
 {
  Add-Content -Path $logfile -Value "--------------------------------------------------------------------------------------------------"
 }
 
 Write-Progress -Activity $scriptname -Status "当前文件：$($TotalList[$i][$j].BaseName)" -CurrentOperation "完成进度：$Counter/$($FilesCount)" -PercentComplete (($Counter/$FilesCount)*100)
 $Counter++
}

 }

Remove-Item -Path $scriptpath -Force
Remove-Item -Path $path\*.bak -Force
Remove-Item -Path $PSScriptRoot\*.log -Force
return $logfile
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function AcadScriptM
{
 param
 (
  [string]$scriptname,
  [string]$accore,
  [string]$path,           #DWG文件路径
  [int32]$session = 10
 )
 #验证进程数
 $patten = "^\+?[1-9][0-9]*$"
 if($($session -match $patten) -eq $False)
 {
  Write-Host "同时运行脚本数必须为非零的正整数，请检查config.ini，退出脚本..."
  exit
 }
 #文件分组
 $TotalList = New-Object System.Collections.ArrayList
 $FilesCount = FileToGroupM -folder $path -extension "dwg" -filenumber $session -TotalList $TotalList

 #分组多进程运行脚本 
 $GroupCount = $TotalList.Count
 $logfile = "{0}\log\{1}脚本执行记录{2}.log" -f $PSScriptRoot,$scriptname,$(Get-Date -Format "yyyy-MM-dd HH-mm-ss")
 $logheader = @"
>> $logfile <<
共处理 $FilesCount 个文件。
**************************************************************************************************
"@
Set-Content -Path $logfile -Value $logheader -Encoding UTF8
 $Counter = 1
 for($i=0;$i -lt $GroupCount;$i++)
 {
  $SessionCount = $TotalList[$i].Count
  $ProcessList = New-Object System.Collections.ArrayList
  $TemplogList = New-Object System.Collections.ArrayList
  for($j=0;$j -lt $SessionCount;$j++)
  {
   $templog = "{0}\templogfile{1}_{2}.txt" -f $env:TEMP,$i,$j
   $script = "{0}\{1}.scr" -f $TotalList[$i][$j].DirectoryName,$TotalList[$i][$j].BaseName
   $process = Start-Process -FilePath $accore -ArgumentList "/i ""$($TotalList[$i][$j].FullName)"" /s ""$script"" /l zh-CN /isolate"  -WindowStyle Hidden -RedirectStandardOutput $templog -PassThru
   [void]$ProcessList.Add($process)
   [void]$TemplogList.Add($templog)
  }
  for($j=0;$j -lt $ProcessList.Count;$j++)
{
 Write-Progress -Activity $scriptname -Status "当前文件：$($TotalList[$i][$j].BaseName)" -CurrentOperation "完成进度：$Counter/$($FilesCount)" -PercentComplete (($Counter/$FilesCount)*100)
 if($ProcessList[$j].HasExited -eq $False)
 {
  $ProcessList[$j].WaitForExit()
 }
 $addcontent = Transcode -filepath $TemplogList[$j]
 Add-Content -Path $logfile -Value ">>$($TotalList[$i][$j].FullName)<<" -Encoding UTF8
 Add-Content -Path $logfile -Value $addcontent -Encoding UTF8
 if($Counter -eq $FilesCount)
 {
  Add-Content -Path $logfile -Value "**************************************************************************************************"
 }
 else
 {
  Add-Content -Path $logfile -Value "--------------------------------------------------------------------------------------------------"
 }
 
 #Write-Progress -Activity $scriptname -Status "当前文件：$($TotalList[$i][$j].BaseName)" -CurrentOperation "完成进度：$Counter/$($FilesCount)" -PercentComplete (($Counter/$FilesCount)*100)
 $Counter++
}

 }

#Remove-Item -Path $scriptpath -Force
Remove-Item -Path $path\*.bak -Force -Recurse
Remove-Item -Path $PSScriptRoot\*.log -Force
return $logfile
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function Out-Excel
{
 param
 (
  $sheetsource,
  $sheetsetsource,
  [string]$Template,
  [string]$sheetname = "图纸信息表",
  [string]$sheetsetname = "图册信息表",
  [string]$tempcsvpath,
  [string]$outpath
 )
 $sheetsource | Export-Csv -Path $tempcsvpath -NoTypeInformation -Encoding Default
 $csvfile = New-Object -TypeName System.IO.FileInfo($tempcsvpath)
 (Get-Content $csvfile) -replace '"' | Set-Content $csvfile
 $Excelfile = New-Object -TypeName System.IO.FileInfo($Template)
 $ExcelList = [OfficeOpenXml.ExcelPackage]::new($Excelfile)
 $listsheet = $ExcelList.Workbook.Worksheets["$sheetname"]
 $format = New-Object -TypeName OfficeOpenXml.ExcelTextFormat
 $format.Encoding = [System.Text.Encoding]::Default
 $format.DataTypes = ([OfficeOpenXml.eDataTypes]::String),([OfficeOpenXml.eDataTypes]::String),([OfficeOpenXml.eDataTypes]::String),([OfficeOpenXml.eDataTypes]::String),([OfficeOpenXml.eDataTypes]::String),([OfficeOpenXml.eDataTypes]::String),([OfficeOpenXml.eDataTypes]::String),([OfficeOpenXml.eDataTypes]::String)
 [void]$listsheet.Cells["A1"].LoadFromText($csvfile,$format)
 $range = $listsheet.Dimension.Address
 $listsheet.Cells[$range].AutoFitColumns()

 $listsheetset = $ExcelList.Workbook.Worksheets["$sheetsetname"]
 $listsheetset.Cells[2,2].Value = $sheetsetsource.Project
 $listsheetset.Cells[4,2].Value = $sheetsetsource.SheetSetName
 $listsheetset.Cells[5,2].Value = $sheetsetsource.Unit
 $listsheetset.Cells[6,2].Value = $sheetsetsource.Date
 $listsheetset.Cells[8,2].Value = $sheetsetsource.Zhuanye
 $listsheetset.Cells[11,2].Value = $sheetsetsource.Mcode
 $listsheetset.Cells[12,2].Value = $sheetsetsource.Version
 $listsheetset.Cells[13,2].Value = $sheetsetsource.Ccode
 $listsheetset.Cells[15,2].Value = $sheetsetsource.workpath
 $listsheetset.Cells[16,2].Value = $sheetsetsource.PDFfile
 
 <#
 for($i=2;$i -le $($source.count+1);$i++)
 {
  $listsheet.Cells[$i,2].Formula = '=图册信息表!$B$8&"-"&图纸信息表!A{0}' -f $i
  $listsheet.Cells[$i,3].Formula = '=图册信息表!$B$11&"-"&图纸信息表!A{0}&"-"&图册信息表!$B$12' -f $i
  $listsheet.Cells[$i,4].Formula = '=图册信息表!$B$13&"-"&图册信息表!$B$8&"-"&图纸信息表!A{0}' -f $i
 }
 #>
 $listfile = New-Object -TypeName System.IO.FileInfo($outpath)
 $ExcelList.SaveAs($listfile)
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function ReadSheetSet
{
 param
 (
  [string]$ExcelPath,
  [string]$SheetName
 )
 $file = New-Object -TypeName System.IO.FileInfo($ExcelPath)
 $ExcelFile = New-Object -TypeName OfficeOpenXml.ExcelPackage($file)
 $Sheetset = $ExcelFile.Workbook.Worksheets["图册信息表"]
 $range = $Sheetset.Dimension.Address
 $endrow = $Sheetset.Dimension.End.Row
 $rowHash = @{ }
 foreach ($cell in $Sheetset.Cells[$range]) 
 {
   if ($cell.Value -ne $null ) 
   { 
    $rowHash[$cell.Start.row] = 1 
   }
  }
 $rows = (1..$endrow).Where({$rowHash[$_] -eq 1})
 $result =[ordered]@{}
 for($i=1;$i -le $rows[-1] ; $i++)
 {
  $key = $Sheetset.Cells[$i,1].Value
  $value = $Sheetset.Cells[$i,2].Value
  $result.$key = $value
 }
 $ExcelFile.Dispose()
 return [PSCustomObject]$result
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function ReadSheet
{
 param
 (
  [string]$ExcelPath,
  [string]$SheetName,
  [int32]$colcount
 )
 $file = New-Object -TypeName System.IO.FileInfo($ExcelPath)
 $ExcelFile = New-Object -TypeName OfficeOpenXml.ExcelPackage($file)
 $Sheet = $ExcelFile.Workbook.Worksheets[$sheetname]
 $range = $Sheet.Dimension.Address
 $endrow = $Sheet.Dimension.End.Row
 $rowHash = @{ }
 foreach ($cell in $Sheet.Cells[$range]) 
 {
   if ($cell.Value -ne $null ) 
   { 
    $rowHash[$cell.Start.row] = 1 
   }
  }
 $rows = (2..$endrow).Where({$rowHash[$_] -eq 1})
 $result = @()
 $tempHash = [ordered]@{}
 for($i=2;$i -le $rows[-1] ; $i++)
 {
  for($j =1 ; $j -le $colcount; $j++)
  {
   $key = $Sheet.Cells[1,$j].Value
   $value = $Sheet.Cells[$i,$j].Value
   $tempHash.$key = $value
  }
  if($tempHash.Values -ne $null )
  {
  $result += [PSCustomObject]$tempHash
  }
 }
 return $result
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function MakeDiskLable
{
param
(
 [string]$_ExcelFile,
 [string]$_WordFile,
 [string]$_SavePath
)
$_sheetset = ReadSheetSet -ExcelPath $_ExcelFile -SheetName "图册信息表"
$_lablename = "{0}({1})" -f $_sheetset.工程名称, $_sheetset.专业名称
$_company = $_sheetset.工点单位
$_duty = $_sheetset.责任人
$_date = $_sheetset.出图日期
$_content = $_sheetset.FullName
$DiskLable = @"
档    号：      流水号：       NO：
案卷题名：$_lablename
编制单位：$_company
责 任 人：$_duty
形成时间：$_date
1、卷内目录
2、$_content
"@
$WordFile = [Xceed.Words.NET.DocX]::Load($_WordFile)
$Table = $WordFile.Tables
for($i=0;$i -lt 2;$i++)
{
 for($j=0;$j -lt 3;$j++)
 {
  [void]$Table[0].Rows[$i].Cells[$j].Paragraphs[0].Append($DiskLable)
 }
}

$WordFile.SaveAs($_SavePath)
}


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function MakeSheetList
{
param
(
 $_Sheet,
 $_SheetSet,
 [System.Collections.Hashtable]$_List = @{ProjectName = "地铁" ; Owner = "市政院" ; Date ="20200101"},
 [String]$_TempDoc = "$PSScriptRoot\Templates\sheetlist.docx",
 [String]$_SavePath = "$PSScriptRoot\Templates\卷内目录.docx",
 [double]$_projectnamefontsize = 8,            #图册全名字体大小
 [double]$_mcodefontsize = 8,                  #地铁代码字体大小
 [double]$_ownerfontsize = 10.5,                #工点单位名称字体大小
 [double]$_namefontsize = 10.5                 #图名字体大小
)

$_SheetCount = $_Sheet.Count

if($_SheetCount -lt 15)
{
 $_ListFile = [Xceed.Words.NET.DocX]::Load($_TempDoc)
 [void]$_ListFile.Paragraphs[2].ReplaceText("N","1")
 [void]$_ListFile.Paragraphs[2].ReplaceText("P","1")
 $_ListTable = $_ListFile.Tables
 [void]$_ListTable[0].Rows[1].Cells[3].Paragraphs[0].Append($_List.ProjectName)
 [void]$_ListTable[0].Rows[1].Cells[3].Paragraphs[0].FontSize($_projectnamefontsize)                     #图册全名字体大小设置
 for($i=1;$i -le $_SheetCount;$i++)
 {
  $sheetname = $_Sheet[$i-1].图纸名
  if($_SheetSet.图名前缀 -ne "")
  {
   $sheetname = "{0} {1}" -f $_SheetSet.图名前缀,$_Sheet[$i-1].图纸名
  } 
  [void]$_ListTable[0].Rows[$i+1].Cells[0].Paragraphs[0].Append($i)
  if($_Sheet[$i-1].图纸序号 -cne $null)                     #扉页没有编号
  {
   if($_Sheet[$i-1].编码.Contains("-"))                   #卷内目录要把"-"改成"/"
   {
   [void]$_ListTable[0].Rows[$i+1].Cells[1].Paragraphs[0].Append($_Sheet[$i-1].编码.Replace("-","/"))
   }
   else
   {
    [void]$_ListTable[0].Rows[$i+1].Cells[1].Paragraphs[0].Append($_Sheet[$i-1].编码)
   }
  }
  [void]$_ListTable[0].Rows[$i+1].Cells[1].Paragraphs[0].FontSize($_mcodefontsize)                   #地铁代码字体大小设置
  [void]$_ListTable[0].Rows[$i+1].Cells[2].Paragraphs[0].Append($_List.Owner)
  [void]$_ListTable[0].Rows[$i+1].Cells[2].Paragraphs[0].FontSize($_ownerfontsize)                   #工点单位名称字体大小设置
  [void]$_ListTable[0].Rows[$i+1].Cells[3].Paragraphs[0].Append($sheetname)
  [void]$_ListTable[0].Rows[$i+1].Cells[3].Paragraphs[0].FontSize($_namefontsize)                   #图名字体大小设置
  [void]$_ListTable[0].Rows[$i+1].Cells[4].Paragraphs[0].Append($_List.Date)
  if($i -eq $_SheetCount)
  {
   [void]$_ListTable[0].Rows[$i+1].Cells[5].Paragraphs[0].Append("$i/$i")
  }
  else
  {
   [void]$_ListTable[0].Rows[$i+1].Cells[5].Paragraphs[0].Append($i)
  }
 }

 $_ListFile.SaveAs($_SavePath)
 $_ListFile.Dispose()
 Write-Progress -Activity "生成卷内目录" -Status "生成卷内目录中" -CurrentOperation "1/1" -PercentComplete (1/1*100)
}
else
{
 $_Pages = [math]::Ceiling(($_SheetCount-14)/15)+1             #向上取整

 $_FirstPage = [Xceed.Words.NET.DocX]::Load($_TempDoc)
 [void]$_FirstPage.Paragraphs[2].ReplaceText("N","1")
 [void]$_FirstPage.Paragraphs[2].ReplaceText("P","$_Pages")

 $_FirstPageTable = $_FirstPage.Tables
 [void]$_FirstPageTable[0].Rows[1].Cells[3].Paragraphs[0].Append($_List.ProjectName)
 [void]$_FirstPageTable[0].Rows[1].Cells[3].Paragraphs[0].FontSize($_projectnamefontsize)          #图册全名字体大小设置
 for($i=1;$i -le 14;$i++)
 {
  $sheetname = $_Sheet[$i-1].图纸名
  if($_SheetSet.图名前缀 -ne "")
  {
   $sheetname = "{0} {1}" -f $_SheetSet.图名前缀,$_Sheet[$i-1].图纸名
  } 
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[0].Paragraphs[0].Append($i)
  if($_Sheet[$i-1].图纸序号 -cne $null)
  {
    if($_Sheet[$i-1].编码.Contains("-"))                   #卷内目录要把"-"改成"/"
   {
   [void]$_FirstPageTable[0].Rows[$i+1].Cells[1].Paragraphs[0].Append($_Sheet[$i-1].编码.Replace("-","/"))
   }
   else
   {
    [void]$_FirstPageTable[0].Rows[$i+1].Cells[1].Paragraphs[0].Append($_Sheet[$i-1].编码)
   }
  }
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[1].Paragraphs[0].FontSize($_mcodefontsize)           #地铁代码字体大小设置
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[2].Paragraphs[0].Append($_List.Owner)
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[2].Paragraphs[0].FontSize($_ownerfontsize)            #工点单位名称字体大小设置
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[3].Paragraphs[0].Append($sheetname)
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[3].Paragraphs[0].FontSize($_namefontsize)           #图名字体大小设置
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[4].Paragraphs[0].Append($_List.Date)
  [void]$_FirstPageTable[0].Rows[$i+1].Cells[5].Paragraphs[0].Append($i)
 }
 $_FirstPage.SaveAs($_SavePath)
 $_SerialNumber = 14

 for($_Page =2;$_Page -le $_Pages;$_Page++)
 {
 
  $_AfterPage = [Xceed.Words.NET.DocX]::Load($_TempDoc)
  [void]$_AfterPage.Paragraphs[2].ReplaceText("N","$_Page")
  [void]$_AfterPage.Paragraphs[2].ReplaceText("P","$_Pages")

  $_AfterPageTable = $_AfterPage.Tables
  for($i=1;$i -le 15;$i++)
 {
  $_SerialNumber++
  $sheetname = $_Sheet[$_SerialNumber-1].图纸名
  if($_SheetSet.图名前缀 -ne "")
  {
   $sheetname = "{0} {1}" -f $_SheetSet.图名前缀,$_Sheet[$_SerialNumber-1].图纸名
  } 
  [void]$_AfterPageTable[0].Rows[$i].Cells[0].Paragraphs[0].Append($_SerialNumber)
  if($_Sheet[$_SerialNumber-1].图纸序号 -cne $null)
  {
    if($_Sheet[$_SerialNumber-1].编码.Contains("-"))                   #卷内目录要把"-"改成"/"
   {
   [void]$_AfterPageTable[0].Rows[$i].Cells[1].Paragraphs[0].Append($_Sheet[$_SerialNumber-1].编码.Replace("-","/"))
   }
   else
   {
    [void]$_AfterPageTable[0].Rows[$i].Cells[1].Paragraphs[0].Append($_Sheet[$_SerialNumber-1].编码)
   }
  }
  [void]$_AfterPageTable[0].Rows[$i].Cells[1].Paragraphs[0].FontSize($_mcodefontsize)                  #地铁代码字体大小设置
  [void]$_AfterPageTable[0].Rows[$i].Cells[2].Paragraphs[0].Append($_List.Owner)
  [void]$_AfterPageTable[0].Rows[$i].Cells[2].Paragraphs[0].FontSize($_ownerfontsize)                    #工点单位名称字体大小设置
  [void]$_AfterPageTable[0].Rows[$i].Cells[3].Paragraphs[0].Append($sheetname)
  [void]$_AfterPageTable[0].Rows[$i].Cells[3].Paragraphs[0].FontSize($_namefontsize)                     #图名字体大小设置
  [void]$_AfterPageTable[0].Rows[$i].Cells[4].Paragraphs[0].Append($_List.Date)
  if($_SerialNumber -eq $_SheetCount)
  {
   [void]$_AfterPageTable[0].Rows[$i].Cells[5].Paragraphs[0].Append("$_SerialNumber/$_SerialNumber")
   break
  }
  else
  {
   [void]$_AfterPageTable[0].Rows[$i].Cells[5].Paragraphs[0].Append($_SerialNumber)
  }
  
 } 
  $_FirstPage.InsertDocument($_AfterPage,$true)
  $_AfterPage.Dispose()
  Write-Progress -Activity "生成卷内目录" -Status "生成卷内目录中" -CurrentOperation "$_Page/$_Pages" -PercentComplete ($_Page/$_Pages*100) 
}

 $_FirstPage.Save()
 $_FirstPage.Dispose()
}
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function GetPdfPageNumber
{
 param
 (
  [string]$PdfPath
 )
 try
 {
  $pdfreader = New-Object iText.Kernel.Pdf.PdfReader($pdfpath)
  $pdfdoc = New-Object iText.Kernel.Pdf.PdfDocument($pdfreader)   
 }
 catch
 {
  $ErrorMessage = $_.Exception.Message
  Write-Warning "Error has occured: $ErrorMessage"
 }
 $NumberOfPages = $pdfdoc.GetNumberOfPages()
 try
 {
  $pdfdoc.Close()
 }
 catch
 {
  $ErrorMessage = $_.Exception.Message
  Write-Warning "Closing document $FilePath failed with error: $ErrorMessage"
 }
 return $NumberOfPages
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function SplitPdfFile
{
 param
 (
  [string]$PdfPath,
  [string]$OutPutFolder,
  [string]$OutPutName = "splittedPDF",
  [int]$SplitPageCount = 1
 )
 try 
 {
  $pdfreader = New-Object iText.Kernel.Pdf.PdfReader($PdfPath)
  $pdfdoc = New-Object iText.Kernel.Pdf.PdfDocument($pdfreader)
  $Splitter = New-Object SplitPdf.CustomSplitter($pdfdoc, $OutPutFolder, "$OutPutName")
  $List = $Splitter.SplitByPageCount($SplitPageCount)
  foreach ($_ in $List) 
   {
     $_.Close()
   }
 }  
 catch 
 {
   $ErrorMessage = $_.Exception.Message
   Write-Warning "Error has occured: $ErrorMessage"
 }
 try 
 {
  $pdfdoc.Close()
 } 
 catch 
 {
  $ErrorMessage = $_.Exception.Message
  Write-Warning "Closing document $FilePath failed with error: $ErrorMessage"
 }
}


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function IsExcelCorrect
{
 param
 (
  [string]$excelpath,
  [string]$sheetname
 )
 
 $file = New-Object -TypeName System.IO.FileInfo($ExcelPath)
 $ExcelFile = New-Object -TypeName OfficeOpenXml.ExcelPackage($file)
 $sheet = $ExcelFile.Workbook.Worksheets[$sheetname]
 if($sheet -eq $null)
 {
  $ExcelFile.Dispose()
  return $false
 }
 else
 {
  $ExcelFile.Dispose()
  return $true
 }
}


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function ControlledAproval
{
 param
 (
  [string]$template = "",
  [System.Collections.Hashtable]$list = @{ProjectName = "" ; FullName = "" ;MCode = "" ; Name = ""},
  [string]$output = ""
 )
 $doc = [Xceed.Words.NET.DocX]::Load($template)
 $doc.ReplaceText("ProjectName",$list.ProjectName)
 $doc.ReplaceText("FullName",$list.FullName)
 $doc.ReplaceText("Name",$list.Name)
 $doc.ReplaceText("MCode",$list.MCode)
 $doc.ReplaceText("Date",$(Get-Date -Format D).toString())
 $doc.SaveAs($output)
 $doc.Dispose()
}


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function SubpackageAproval
{
 param
 (
  [string]$template = "",
  [System.Collections.Hashtable]$list = @{ProjectName = "" ; FullName = "" ;MCode = "" ; Company = ""},
  [string]$output = ""
 )
 $doc = [Xceed.Words.NET.DocX]::Load($template)
 $doc.ReplaceText("ProjectName",$list.ProjectName) 
 $doc.ReplaceText("FullName",$list.FullName)
 $doc.ReplaceText("Company",$list.Company)
 $doc.ReplaceText("MCode",$list.MCode)
 $doc.ReplaceText("Year",$(Get-Date -Format %yy).toString())
 $doc.SaveAs($output)
 $doc.Dispose()
}


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function PlotRequest
{
 param
 (
  [string]$templateA = "",
  [string]$template0 = "",
  [System.Collections.Hashtable]$list = @{Version = "" ; ShortName = "" ; FullName = "" ;Name = "" ; Company = "" ; MCode = ""},
  [string]$output = ""
 )
 if($list.Version -eq "是")
 {
  $doc = [Xceed.Words.NET.DocX]::Load($template0)
 }
 elseif($list.Version -eq "否")
 {
  $doc = [Xceed.Words.NET.DocX]::Load($templateA)
 }
 else
 {
  Write-Host "获取送审版信息失败！"
  return
 }
 $doc.ReplaceText("[ShortName]",$list.ShortName) 
 $doc.ReplaceText("[FullName]",$list.FullName)
 $doc.ReplaceText("[Company]",$list.Company)
 $doc.ReplaceText("[Name]",$list.Name)
 $doc.ReplaceText("[MCode]",$list.MCode)
 $doc.ReplaceText("[Date]",$(Get-Date -Format D).toString())
 $doc.ReplaceText("[Year]",$(Get-Date -Format %yy).toString())
 $doc.SaveAs($output)
 $doc.Dispose()
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function CheckFolder
{
 param
 (
  [string]$path
 )
 $pathinfo = New-Object System.IO.DirectoryInfo($path)
 if($pathinfo.Exists -eq $false)
 {
  return $true
 }
 else
 {
  return $false
 }
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################
Function MergePdf
{
 param
 (
  [string]$outfile,
  [System.Collections.ArrayList]$inputfiles
 )
 $pdfwriter = [iText.Kernel.Pdf.PdfWriter]::new($outfile)
 $pdf = [iText.Kernel.Pdf.PdfDocument]::new($pdfwriter)
 $merger = [iText.Kernel.Utils.PdfMerger]::new($pdf)
 $pages = $inputfiles.Count
 for($i=0;$i -lt $inputfiles.Count;$i++)
 {
  $pdfreader = [iText.Kernel.Pdf.PdfReader]::new($inputfiles[$i])
  $soursepdf = [iText.Kernel.Pdf.PdfDocument]::new($pdfreader)
  [void]$merger.Merge($soursepdf,1,$soursepdf.GetNumberOfPages()).SetCloseSourceDocuments($true)
  $soursepdf.Close()
  Write-Progress -Activity "合并PDF" -Status "合并PDF文件中" -CurrentOperation "$($i+1)/$pages" -PercentComplete ((($i+1)/$pages)*100) 
 }
 $merger.Close()
 $pdf.Close()
}

##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################


##################################################################
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
#                                                                #
##################################################################


# SIG # Begin signature block
# MIIFlwYJKoZIhvcNAQcCoIIFiDCCBYQCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUcqXSuxoXtO+CO/ZRvDmfUYZu
# UXCgggMtMIIDKTCCAhWgAwIBAgIQRRn0wX7drrlAomMxELaoYjAJBgUrDgMCHQUA
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
# AgELMQ4wDAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQUdaNsZV44atU2tFcX
# iPdfjF8vG/cwDQYJKoZIhvcNAQEBBQAEggEAeZvx021gzqxfy5EgBRyT8kcprDJt
# SvuRzNXHK24HBN9H4hlhMptke2j35zT70djvuhQW0n+Ffa9NZOjfsbJL8mk8H9bZ
# ANTW9BWy055RPEvFRzitvLmC/+uEcHznIfc09dEDSsi49Ko3gpDOFZZWphFBQBpX
# BiSBB6vK/yApta3jZ9rc/ERkGgm58+w+PN1Z/IfXtHyiMyWGThZTLw2zvFDAVEui
# 4sTy17HhurK7fsgXq0lGPyIuj/AEe1o0bd7jPDyDaf4GFD4tC5JzSlr7eZyrqyz/
# gqDMPiye4X8xC7XFYsDXnJ9yMuccgh7b6plt+bAHGjf9lzVGsbxrq4Sm4Q==
# SIG # End signature block
