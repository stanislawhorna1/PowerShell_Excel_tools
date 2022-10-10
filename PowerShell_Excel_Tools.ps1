$filename = Read-Host "Enter filename: "
try {
    $csv = Import-Csv ./Output/$filename
}
catch {
    {Impossible to read the file}
    exit 1
}
$Headers=(Get-Member -InputObject $csv[0] -MemberType NoteProperty).Name
for ($i = 0; $i -lt $Headers.Count; $i++) {
    $Column = $Headers[$i]
    Write-Host "$i. $Column"
}
$ind = Read-Host "Enter index of column which you would like to select: "
$Column = $Headers[$ind]

$csv | Sort-Object -Property $Column - 


Remove-Variable filename 
Remove-Variable csv
Remove-Variable Headers 
Remove-Variable Column
Remove-Variable ind