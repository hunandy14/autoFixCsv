# 載入Get-Encoding函式
Invoke-RestMethod 'raw.githubusercontent.com/hunandy14/Get-Encoding/master/Get-Encoding.ps1'|Invoke-Expression


# 輸出至檔案
function WriteContent {
    [CmdletBinding(DefaultParameterSetName = "D")]
    param (
        [Parameter(Mandatory, Position = 0, ParameterSetName = "")]
        [string] $Path,
        [Parameter(Position = 1, ParameterSetName = "C")]
        [int] $Encoding,
        [Parameter(Position = 1, ParameterSetName = "D")]
        [switch] $DefaultEncoding,

        [Parameter(ParameterSetName = "")]
        [switch] $NoNewline,
        [Parameter(ParameterSetName = "")]
        [switch] $Append,
        [Parameter(ValueFromPipeline, ParameterSetName = "")]
        [System.Object] $InputObject
    )
    BEGIN {
        # 獲取編碼
        if ($DefaultEncoding) { # 使用當前系統編碼
            # $Enc = [Text.Encoding]::Default
            $Enc = PowerShell.exe -C "& {return [Text.Encoding]::Default}"
        } elseif ((!$Encoding) ) { # 完全不指定預設
            $Enc = New-Object System.Text.UTF8Encoding $False
            # $Enc = [Text.Encoding]::Default
        } elseif ($Encoding -eq 65001) { # 指定UTF8
            $Enc = New-Object System.Text.UTF8Encoding $False
        } else { # 使用者指定
            $Enc = [Text.Encoding]::GetEncoding($Encoding)
        }

        # 建立檔案
        if (!$Append) { 
            (New-Item $Path -ItemType:File -Force) | Out-Null
        } $Path = [IO.Path]::GetFullPath($Path)
        
    } process{
        [IO.File]::AppendAllText($Path, "$_`n", $Enc);
    }
    END { }
} # WriteContent

# 自動修復CSV檔案格式
function autoFixCsv {
    [CmdletBinding(DefaultParameterSetName = "A")]
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $Path,
        [Parameter(Position = 1, ParameterSetName = "A")]
        [string] $Destination,
        [Parameter(Position = 1, ParameterSetName = "B")]
        [switch] $OutObject,
        [Parameter(Position = 1, ParameterSetName = "C")]
        [switch] $Overwrite,
        [Parameter(Position = 2, ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $TrimValue,
        [switch] $OutNull
    )
    # 檢查
    if (!(Test-Path -PathType:Leaf $Path)) { throw "Input file does not exist"; return}
    $File = Get-Item $Path
    if (!$Destination) { $Destination = $File.BaseName + "_fix" + $File.Extension }
    if ($OutObject) { $OutNull = $true } # 輸出物件的時候不要輸出信息
    if ($Destination -eq $Path) {
        Write-Host "Warring:: The source path is the same as the destination path. If you want to overwrite, Please use `"-Overwrite`"." -ForegroundColor:Yellow; return
    }
    if ($Overwrite) { $Destination = $Path } # 覆蓋原檔
    
    # 處理編碼
    if ($Encoding) { # 自訂編碼
        $Enc = $Encoding
    } else {
        if ($UTF8) { # 不帶BOM的UTF8
            $Enc = New-Object System.Text.UTF8Encoding $False
        } else { # 系統語言
            if (!$__SysEnc__) { $Script:__SysEnc__ = [Text.Encoding]::GetEncoding((powershell -nop "([Text.Encoding]::Default).WebName")) }
            $Enc = $__SysEnc__
        }
    } ($Enc.EncodingName) -match '\((.*?)\)'|Out-Null
    $EncName = $matches[1]
    # 讀取檔案
    try {
        $Contact = [IO.File]::ReadAllLines($Path, $Enc)
    } catch { Write-Error ($Error[$Error.Count-1]); return }
    
    # 輸出訊息
    if (!$OutNull) {
        Write-Host "From [" -NoNewline
        Write-Host $EncName -NoNewline -ForegroundColor:Yellow
        Write-Host "]:: $Path"
        Write-Host "  └─>[$EncName]:: $Destination"
        Write-Host "Convert start... " -NoNewline
    }
    
    # 計時開始
    $StWh = New-Object System.Diagnostics.Stopwatch; $StWh.Start()
    
    # 轉換至物件
    $Csv = $Contact|ConvertFrom-Csv
    
    # 消除多餘空白
    if ($TrimValue) {
        foreach ($Item in $CSV) {
            foreach ($_ in $Item.PSObject.Properties) {
                if ($_.Value) { $_.Value = ($_.Value).trim() } else { $_.Value=$null }
            }
        }
    }
    
    # 輸出物件
    if ($OutObject) {
        return $Csv
    # 輸出Csv檔案
    } else {
        $Contact = $Csv|ConvertTo-Csv -NoTypeInformation
        [IO.File]::WriteAllLines($Destination, $Contact, $Enc)
        # 輸出提示訊息
        if (!$OutNull) {
            $StWh.Stop()
            $Time = "{0:hh\:mm\:ss\.fff}" -f [timespan]::FromMilliseconds($StWh.ElapsedMilliseconds)
            Write-Host "Finish [" -NoNewline; Write-Host $Time -NoNewline -ForegroundColor:DarkCyan; Write-Host "]"
        }
    }
} # autoFixCsv 'sample1.csv'
# autoFixCsv 'sample1.csv'
# autoFixCsv 'sample1.csv' -TrimValue
# autoFixCsv 'sample1.csv' -OutObject -TrimValue -UTF8
# (autoFixCsv 'sample1.csv' -OutObject)|Export-Csv 'sample1_fix.csv'
# autoFixCsv 'AddItem.csv' -Encoding:(Get-Encoding 932)
# autoFixCsv 'sample1.csv' -Overwrite -UTF8

# 循環 CSV Item 物件
function ForEachCsvItem {
    [CmdletBinding(DefaultParameterSetName = "A")]
    param (
        # 循環項目的ForEach區塊
        [Parameter(Position = 0, ParameterSetName = "A", Mandatory)]
        [Parameter(Position = 1, ParameterSetName = "B", Mandatory)]
        [scriptblock] $ForEachBlock,
        # PS表格 轉換為自訂 哈希表
        [Parameter(Position = 0, ParameterSetName = "B")]
        [scriptblock] $ConvertObject={
            [Object] $obj = @{}
            foreach ($it in ($_.PSObject.Properties)) {
                $obj += @{$it.Name = $it.Value}
            } return $obj
        },
        # 輸入的物件
        [Parameter(ParameterSetName = "", ValueFromPipeline)]
        [Object] $_
    ) BEGIN { } PROCESS {
    foreach ($_ in $_) {
        $_ = &$ConvertObject($_)
        &$ForEachBlock($_)
    } } END { }
}

# 使用預設轉換函式
# (autoFixCsv 'sample2.csv' -OutObject)|ForEachCsvItem{ $_.'個人ＩＤ' }

# 自訂轉換函式
# $csv = (autoFixCsv 'sample2.csv' -OutObject)
# $ConvertObject={
#     $obj = @{}
#     $title_idx=0
#     $field_idx=1
#     foreach ($it in ($_.PSObject.Properties)) {
#         if($Title[$title_idx]){
#             $Name = $Title[$title_idx]
#             $title_idx=$title_idx+1
#         } else {
#             $Name = "field_$($field_idx)"
#             $field_idx = $field_idx+1
#         } $obj += @{$Name = $it.Value}
#     } return $obj
# }
# $Title=@('ID', 'Title')
# $csv|ForEachCsvItem -ConvertObject ([ScriptBlock]::Create({$Title=$Title}.ToString() + $ConvertObject)) {
#     Write-Host "$($_.ID) | $($_.Title) | $($_.field_1) | $($_.field_2)"
# }
