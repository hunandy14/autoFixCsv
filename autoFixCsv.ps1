function autoFixCsv {
    param (
        [Parameter(Position = 0, ParameterSetName = "", Mandatory)]
        [string] $Path,
        [Parameter(Position = 1, ParameterSetName = "")]
        [string] $Destination
    )
    if (!$Destination) { $Destination = $Path }
    (Import-Csv $Path) | Export-Csv $Destination -NoTypeInformation
} 
# autoFixCsv 'sample2.csv' 'out.csv'
# autoFixCsv 'out.csv'
