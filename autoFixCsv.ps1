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


# 自動修復CSV檔案格式
function autoFixCsv {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $Path,
        [Parameter(Position = 1, ParameterSetName = "")]
        [string] $Destination,
        [switch] $TrimValue
    )
    # 檢查
    if (!(Test-Path -PathType:Leaf $Path)) { throw "Input file does not exist" }
    $File = Get-Item $Path
    if (!$Destination) { $Destination = $File.BaseName + "_fix" + $File.Extension }

    # 轉換CSV檔案
    $CSV = (Import-Csv $Path -Encoding:(DefEnc))
    if ($TrimValue) { # 消除多餘空白
        foreach ($Item in $CSV) {
            ($Item.PSObject.Properties)|ForEach-Object{ $_.Value = ($_.Value).trim() }
        }
    } $CSV|Export-Csv $Destination -Encoding:(DefEnc) -NoTypeInformation

    # 輸出訊息
    $EncName = DefEnc -FullName
    Write-Host "From [$EncName]::" -NoNewline
    Write-Host $Path -NoNewline -ForegroundColor:White
    Write-Host " convert to"
    Write-Host "   └─[$EncName]::" -NoNewline
    Write-Host $Destination -ForegroundColor:Yellow
} # autoFixCsv 'sample1.csv'
# autoFixCsv 'sample1.csv' -TrimValue
# autoFixCsv 'sample1.csv' 'sample1_fix.csv'
