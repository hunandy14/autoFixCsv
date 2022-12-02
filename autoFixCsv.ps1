# 載入Get-Encoding函式
Invoke-RestMethod 'raw.githubusercontent.com/hunandy14/Get-Encoding/master/Get-Encoding.ps1'|Invoke-Expression

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
        
        [Parameter(ParameterSetName = "")]
        [object] $Sort,
        [Parameter(ParameterSetName = "")]
        [object] $Unique,
        [switch] $Count, # 統計有多少重複的 (Unique啟用時才能統計)
        [Parameter(ParameterSetName = "")]
        [object] $Select,
        
        [Parameter(ParameterSetName = "")]
        [object] $WhereField,
        [Parameter(ParameterSetName = "")]
        [object] $WhereValue,
        
        [Parameter(ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $UTF8BOM,
        
        [switch] $TrimValue,
        [switch] $AddIndex,
        [switch] $OutNull,
        
        [Parameter(ParameterSetName = "")]
        [scriptblock] $ScriptBlock
    )
    # 檢查
    [IO.Directory]::SetCurrentDirectory(((Get-Location -PSProvider FileSystem).ProviderPath))
    $Path = [System.IO.Path]::GetFullPath($Path)
    if (!(Test-Path -PathType:Leaf $Path)) { Write-Error "Input file `"$Path`" does not exist" -ErrorAction:Stop }
    if (!$Destination) {
        $File = Get-Item $Path
        $Destination = ($File.BaseName + "_fix" + $File.Extension)
    }
    if ($OutObject) { $OutNull = $true } # 輸出物件的時候不要輸出信息
    if ($Destination -eq $Path) { Write-Host "Warring:: The source path is the same as the destination path. If you want to overwrite, Please use `"-Overwrite`"." -ForegroundColor:Yellow; return }
    if ($Overwrite) { $Destination = $Path } # 覆蓋原檔
    
    # 處理編碼
    if ($Encoding) { # 自訂編碼
        $Enc = $Encoding
    } else { # 預選項編碼
        if ($UTF8) {
            $Enc = New-Object System.Text.UTF8Encoding $False
        } elseif ($UTF8BOM) {
            $Enc = New-Object System.Text.UTF8Encoding $True
        } else { # 系統語言
            if (!$__SysEnc__) { $Script:__SysEnc__ = [Text.Encoding]::GetEncoding((powershell -nop "([Text.Encoding]::Default).WebName")) }
            $Enc = $__SysEnc__
        }
    } ($Enc.EncodingName) -match '\((.*?)\)'|Out-Null
    $EncName = $matches[1]
    
    # 讀取檔案
    try { # 阻止編碼錯誤時繼續執行代碼
        $Content = [IO.File]::ReadAllLines($Path, $Enc)
    } catch { Write-Error $PSItem -ErrorAction -ErrorAction:Stop }
    
    # 輸出訊息
    if (!$OutNull) {
        Write-Host "From [" -NoNewline
        Write-Host $EncName -NoNewline -ForegroundColor:Yellow
        Write-Host "]:: $Path"
        Write-Host "  └──[$EncName]:: " -NoNewline
        Write-Host $Destination -ForegroundColor:White
        Write-Host "Convert start... " -NoNewline
    }
    
    # 計時開始
    $StWh = New-Object System.Diagnostics.Stopwatch; $StWh.Start()
    
    # 轉換至物件
    try {
        $Csv = $Content|ConvertFrom-Csv
    } catch { Write-Error $PSItem -ErrorAction -ErrorAction:Stop }
    
    
    
    # 排序
    if ($Sort) { $Csv = $Csv|Sort-Object -Property $Sort }
    
    # 消除相同
    if ($Unique -or ($Unique -eq "")) {
        # 方法1 (能刪除重複，但在超大數據下似乎不能保留當前順序的第一個)
        # $CsvUq = $Csv|Sort-Object -Property $Unique -Unique
        # $Csv = ([Linq.Enumerable]::Intersect([object[]]$Csv, [object[]]$CsvUq))
        # 方法2
        $hashTable = @{}; $Array = @(); $idx=0
        $Csv|ForEach-Object{
            if($Unique -eq "") { $item = $_ } else { $item = $_|Select-Object -Property $Unique }
            $str  = ($item|ConvertTo-Csv -NoTypeInformation)[1]
            try { $hashTable.Add($str, $idx);$flag=$True } catch { $flag=$False }
            if ($flag) {
                if ($Count) { # 統計總共有多少重複的
                    $item = $_|Select-Object @{Name='Count';Expression={1}},*
                } else {
                    $item = $_
                } $Array += $item; $idx++
            } else { # 統計總共有多少重複的
                if ($Count) { $Array[$hashTable.$str].Count++ }
            }
        }; $Csv=$Array; $Array=$null
    }
    
    # 取出特定項目
    if ($Select) { 
        if ($Count) {
            $Csv = $Csv|Select-Object -Property Count,$Select
        } else {
            $Csv = $Csv|Select-Object -Property $Select
        }
    }
    
    # 消除多餘空白
    if ($TrimValue) {
        foreach ($Item in $Csv) {
            foreach ($_ in $Item.PSObject.Properties) {
                if ($_.Value) { $_.Value = ($_.Value).trim() } else { $_.Value=$null }
            }
        }
    }
    
    # 追加流水番號
    if ($AddIndex) {
        for ($i = 0; $i -lt $Csv.Count; $i++) {
            $Csv[$i] = $Csv[$i]|Select-Object @{Name='Index';Expression={($i+1)}},*
        }
    }
    
    # 取出特定數值的項目
    if ($WhereField) {
        $Array = @() # 寫在這裡是為了卡住如果ItemValue是NULL至少把輸出變成空白
        if ($null -ne $WhereValue ) {
            # 輸入如果不是陣列則將他轉微陣列
            if ($WhereField -isnot [array]) { $WhereField = @($WhereField) }
            if ($WhereValue -isnot [array]) { $WhereValue = @($WhereValue) }
            # PSObject轉Array語句
            # $ConvertToArray=@()
            # for ($i = 0; $i -lt $WhereField.Count; $i++) {
            #     $ConvertToArray += "`$Item.(`$WhereField[$i])"
            # } $ConvertToArray = "@($($ConvertToArray -join ', '))"
            
            # 將輸入的陣列轉成同樣的 CsvObject
            $Item2 = $Csv[0]|Select-Object $WhereField
            for ($i = 0; $i -lt $WhereField.Count; $i++) {
                $FidleName = $WhereField[$i]
                $InputValue = $WhereValue[$i]
                $Item2.$FidleName = $InputValue
            } $InItemStr   = ($Item2|ConvertTo-Csv -NoTypeInformation)[1]
            
            # 找出相同的項目加入新陣列中
            for ($i = 0; $i -lt $Csv.Count; $i++) {
                $Item = $Csv[$i]|Select-Object $WhereField # 取出特定字段
                $CsvItemStr  = ($Item|ConvertTo-Csv -NoTypeInformation)[1]
                if($InItemStr -eq $CsvItemStr){ $Array += $Csv[$i] }
                # $ItemArr = $ConvertToArray|Invoke-Expression
                # if ($ItemArr) {
                #     $IsEqual = !(Compare-Object $ItemArr $WhereValue -SyncWindow 0)
                #     if($IsEqual){ $Array += $Csv[$i] }
                # }
            }
        } else {
            Write-Host "Error:: -WhereValue is Null" -ForegroundColor:Yellow; return
        } $Csv=$Array; $Array=$null
    }
    
    # 自訂功能
    if ($ScriptBlock) { & $ScriptBlock }
    
    
    
    # 輸出物件
    if ($OutObject) {
        return $Csv
        
    # 輸出Csv檔案
    } else {
        $Content = $Csv|ConvertTo-Csv -NoTypeInformation
        if(!$Content){ $Content = "" }
        $Destination = [System.IO.Path]::GetFullPath($Destination)
        if ($Destination -and !(Test-Path $Destination)) { New-Item $Destination -Force|Out-Null }
        [IO.File]::WriteAllLines($Destination, $Content, $Enc)
        # 輸出提示訊息
        if (!$OutNull) {
            $StWh.Stop()
            $Time = "{0:hh\:mm\:ss\.fff}" -f [timespan]::FromMilliseconds($StWh.ElapsedMilliseconds)
            Write-Host "Finish [" -NoNewline; Write-Host $Time -NoNewline; Write-Host "]"
            if(!$Content){ Write-Host "Warring:: Csv out content is empty" -ForegroundColor:Yellow }
        }
    }
} # autoFixCsv 'sample1.csv'
# autoFixCsv 'sample1.csv'
# autoFixCsv 'sample1.csv' 'sample1_fix.csv'
# autoFixCsv 'sample1.csv' 'Z:\sample1_fix.csv'
# autoFixCsv 'sample1.csv' -TrimValue -UTF8
# autoFixCsv 'sample1.csv' -OutObject -TrimValue -UTF8
# (autoFixCsv 'sample1.csv' -OutObject)|Export-Csv 'sample1_fix.csv'
# autoFixCsv 'AddItem.csv' -Encoding:(Get-Encoding 932)
# autoFixCsv 'sample1.csv' -Overwrite -UTF8
# autoFixCsv 'sample1.csv' -Overwrite -UTF8BOM
# autoFixCsv 'sort.csv' -Unique A -UTF8
# autoFixCsv 'sort.csv' -Sort ID
# autoFixCsv 'sort.csv' -Sort ID,A,B
# autoFixCsv 'sort.csv' -Sort ID,A,B -Unique A
# autoFixCsv 'sort.csv' -Sort A,B -Unique ID
# autoFixCsv 'sort.csv' -Unique C,D
# autoFixCsv 'sort.csv' -Select A,B
# autoFixCsv 'sort.csv' -Unique E -UTF8
# autoFixCsv 'sample3.csv' -UTF8
# autoFixCsv 'sort.csv' -Unique "" -UTF8
# autoFixCsv 'sort.csv' -Unique "" -Count -UTF8BOM
# autoFixCsv 'sort.csv' -Unique "" -Count -UTF8BOM -AddIndex
# autoFixCsv 'sort.csv' -Unique "A" -Select "A" -Count -UTF8BOM

# autoFixCsv 'sort.csv' -WhereField A,B -WhereValue B,1 -UTF8
# autoFixCsv 'sample2.csv' -WhereField 会社略称 -WhereValue ＨＩＳＹＳ－ＥＳ -UTF8
# autoFixCsv 'sample2.csv' -WhereField 会社略称 -WhereValue "" -UTF8
# autoFixCsv 'sample2.csv' -WhereField 会社略称,役員並び順 -WhereValue ＨＩＳＹＳ,99 -UTF8
# autoFixCsv 'sort.csv' -WhereField ID,B -WhereValue 10,1 -UTF8
# autoFixCsv 'sort.csv' -UTF8
# autoFixCsv 'sample1.csv' -UTF8BOM

# 測試自訂功能
# autoFixCsv 'sort.csv' -Unique "A" -Select "A" -UTF8BOM -ScriptBlock{
#     for ($i = 0; $i -lt $Csv.Count; $i++) {
#         $Csv[$i] = $Csv[$i]|Select-Object @{Name='Index';Expression={($i+1)}},*
#     }
# }

# 例外測試
# autoFixCsv 'XXXXXXX.csv'
# autoFixCsv 'sample2.csv'
# autoFixCsv 'sort.csv' -Unique G
# try { autoFixCsv 'XXXXXXX.csv' } catch { Write-Output "Catch:: " ($Error[$Error.Count-1]) }




# 循環 CSV Item 物件 (並由陣列轉換為哈希表)
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
        [Object] $InputObject
    ) BEGIN { } PROCESS {
    foreach ($_ in $InputObject) {
        $obj = &$ConvertObject($_)
        &$ForEachBlock($obj)
    } } END { }
}

# 使用預設轉換函式
# (autoFixCsv 'sample2.csv' -OutObject -UTF8)|ForEachCsvItem{ $_.'個人ＩＤ' }

# 自訂轉換函式
# $csv = (autoFixCsv 'sample2.csv' -OutObject -UTF8)
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




# 檢測CSV
function CheckCsv {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $Path,
        [Parameter(ParameterSetName = "")]
        [string] $Title,
        [Parameter(ParameterSetName = "")]
        [object] $TypeIsInt,
        [string] $ItemCount,
        
        [Parameter(ParameterSetName = "")]
        [Text.Encoding] $Encoding,
        [switch] $UTF8,
        [switch] $UTF8BOM
    )
    # 處理編碼
    if ($Encoding) { # 自訂編碼
        $Enc = $Encoding
    } else { # 預選項編碼
        if ($UTF8) {
            $Enc = New-Object System.Text.UTF8Encoding $False
        } elseif ($UTF8BOM) {
            $Enc = New-Object System.Text.UTF8Encoding $True
        } else { # 系統語言
            if (!$__SysEnc__) { $Script:__SysEnc__ = [Text.Encoding]::GetEncoding((powershell -nop "([Text.Encoding]::Default).WebName")) }
            $Enc = $__SysEnc__
        }
    }
    
    # 檢查
    $Path = [IO.Path]::GetFullPath([IO.Path]::Combine((Get-Location -PSProvider FileSystem).ProviderPath, $Path))
    
    # 檔案是否存在
    if (!(Test-Path -PathType:Leaf $Path)) { Write-Output "Input file `"$Path`" does not exist"; return}
    
    # 讀取檔案
    try { # 阻止編碼錯誤時繼續執行代碼
        $Content = [IO.File]::ReadAllLines($Path, $Enc)
    } catch { Write-Error $PSItem -ErrorAction -ErrorAction:Stop }
    # 轉換至物件
    try {
        $Csv = $Content|ConvertFrom-Csv
    } catch { Write-Error $PSItem -ErrorAction -ErrorAction:Stop }
    
    # 檔案是否為空檔
    If ((Get-Item $Path).length -eq 0kb) { Write-Output "Input file `"$Path`" is zero byte"; return}
    
    # 校驗字段
    if ($Title) {
        if ($Content[0] -ne $Title) { Write-Output "Title check Fail"; return}
    }
    
    # 校驗CSV
    if (!$Csv) { Write-Output "Content check Fail"; return}
    
    # 校驗資料數目是否有少
    if ($ItemCount) {
        $idx=0; $ErrorCount=0
        ($Content)|ForEach-Object{
            $line = $_
            $flag=$true; $c=0
            for ($i = 0; $i -lt $line.Length; $i++) {
                $char=$line[$i]
                if ($flag) {
                    if ($char -eq "`"") {
                        $flag=$false
                    } elseif ($char -eq ',') {
                        $c++
                    }
                } else {
                    if ($char -eq "`"") {
                        $flag=$true
                    }
                }
            }
            if (($ItemCount-1) -ne $c) { Write-Output "In line [$idx] item quantity has wrong"; $ErrorCount++ }
            $idx++
        }
        if ($ErrorCount -ne 0) { return }
    }
    
    # 驗證型態
    if ($TypeIsInt) {
        $ErrorCount=0
        for ($j = 0; $j -lt $Csv.Count; $j++) {
            $Item = $Csv[$j]|Select-Object $TypeIsInt
            $Item = ($Item.PSObject.Properties.Value)
            for ($i = 0; $i -lt $Item.Count; $i++) {
                $Value = $Item[$i]
                if ($Value -notmatch "^[0-9]*$") {
                    Write-Output "In line [$j], item [$($TypeIsInt[$i])] has the wrong type"; $ErrorCount++
                }
            }
        }
    }
}
# 路徑不存在
# CheckCsv "ck\0.csv"
# 空檔
# CheckCsv "ck\1.0byte.csv"
# 字段
# CheckCsv "ck\2.onlyitem.csv" -Title "A,B,C,D"
# 項目至少一項
# CheckCsv "ck\3.onlytitle.csv" -Title "A,B,C,D"
# 檢查資料中有沒有少項目的(逗號缺少)
# CheckCsv "ck\4.coma.csv" -Title "A,B,C,D" -ItemCount 4
# 檢查資料型態
# CheckCsv "ck\5.type.csv" -Title "A,B,C,D" -ItemCount 4 -TypeIsInt B,C
