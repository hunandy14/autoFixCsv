# 當前終端預設編碼文字 (Import-Csv等函式預設值)
function DefEnc {
    param(
        [switch] $FullName
    )
    if (!$FullName) {
        if ($PSVersionTable.PSVersion.Major -ge 7) { $Result = "UTF8" } else { $Result = "Default" }
    } else {
        (([System.Text.Encoding]::Default).EncodingName) -match '\((.*?)\)'|Out-Null
        $Result = $matches[1]
    } return $Result
} # DefEnc -FullName

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
}

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
        [switch] $TrimValue,
        [switch] $OutNull
    )
    # 檢查
    $EncName = DefEnc -FullName
    if (!(Test-Path -PathType:Leaf $Path)) { throw "Input file does not exist" }
    $File = Get-Item $Path
    if (!$Destination) { $Destination = $File.BaseName + "_fix" + $File.Extension }
    if ($OutObject) { $OutNull = $true }
    
    
    # 輸出訊息
    if (!$OutNull) {
        Write-Host "From [$EncName]::" -NoNewline
        Write-Host $Path -NoNewline -ForegroundColor:White
        Write-Host " convert to ..."
    }
    
    # 計時開始
    $Date = (Get-Date); $StWh = New-Object System.Diagnostics.Stopwatch; $StWh.Start()
    # 轉換CSV檔案
    $CSV = (Import-Csv $Path -Encoding:(DefEnc))
    if ($TrimValue) { # 消除多餘空白
        foreach ($Item in $CSV) {
            foreach ($_ in $Item.PSObject.Properties) {
                if ($_.Value) { $_.Value = ($_.Value).trim() } else { $_.Value=$null }
            }
        }
    }
    # 計時停止
    $StWh.Stop()
    
    # 輸出物件
    if ($OutObject) {
        return $CSV
    # 輸出Csv檔案
    } else {
        $CSV|Export-Csv $Destination -Encoding:(DefEnc) -NoTypeInformation
        # 輸出提示訊息
        if (!$OutNull) {
            Write-Host "   └─[$EncName]::" -NoNewline
            Write-Host $Destination -ForegroundColor:Yellow
            $Time = "{0:hh\:mm\:ss\.fff}" -f [timespan]::FromMilliseconds($StWh.ElapsedMilliseconds)
            Write-Host "[$Date] 開始執行, 耗時 [" -NoNewline; Write-Host $Time -NoNewline -ForegroundColor:DarkCyan; Write-Host "] 執行結束"
        }
    }
} # autoFixCsv 'sample1.csv'
# autoFixCsv 'sample1.csv'
# autoFixCsv 'sample1.csv' -TrimValue
# autoFixCsv 'sample1.csv' -OutObject
# (autoFixCsv 'sample1.csv' -OutObject)|Export-Csv 'sample1_fix.csv' -NoTypeInformation


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
        $_ = $ConvertObject.Invoke($_)
        $ForEachBlock.Invoke($_)
    } } END { }
}

# 使用預設轉換函式
# (autoFixCsv 'sample2.csv' -OutObject)|ForEachCsvItem{ $_.'個人ＩＤ' }

# 自訂轉換函式
# $csv = (autoFixCsv 'sample2.csv' -OutObject)
# $ConvertObject={
#     [Object] $obj = @{}
#     $i=0; foreach ($it in ($_.PSObject.Properties)) {
#         if ($i -eq 0) { } elseif ($i -eq 1) {
#             $obj += @{"Title" = $it.Value}
#         } else {
#             $obj += @{"field_$($i-1)" = $it.Value}
#         } $i=$i+1
#     } return $obj
# }
# $csv|ForEachCsvItem -ConvertObject:$ConvertObject{
#     $_.Title
# }
