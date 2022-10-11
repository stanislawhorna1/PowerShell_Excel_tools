$filename = Read-Host "Enter filename: "
try {
    $csv = Import-Csv ./Output/$filename
}
catch {
    { Impossible to read the file }
    exit 1
}


$title = 'Sorting'
$question = 'Would you like to sort input data?'
$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Sort data"))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Do not sort data"))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit program"))
$decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
if ($decision -eq 3) {
    Remove-Variable filename
    exit 0
}

# /// SORTING ///
if ($decision -eq 0) {
    $Headers = (Get-Member -InputObject $csv[0] -MemberType NoteProperty).Name
    Write-Host ""
    Write-Host "Columns available in source file"
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $Column = $Headers[$i]
        Write-Host "$i. $Column"
    }
    $ind = Read-Host "Select the column you want to use for sorting:"
    $column_sort = $Headers[$ind]
    Write-Host ""
    if ($column_sort.ToLower() -like "*time*" -or $column_sort.ToLower() -like "*date*") {
        $csv = ($csv | Sort-Object { Get-Date $_.$column_sort } -Descending)
        #$csv | Format-Table
    }
    elseif ($column_sort.ToLower() -eq "id" -or $column_sort.ToLower() -like "*num*") {
        $csv = ($csv | Sort-Object { [int]$_.$column_sort })
        #$csv | Format-Table
    }
    else {
        $csv = ($csv | Sort-Object -Property $column_sort)
        # $csv | Format-Table
    }
    Remove-Variable Headers 
    Remove-Variable Column
    Remove-Variable ind
}

$csv | Format-table
Write-Host ""
$title = 'Remove duplicates'
$question = 'Would you like to remove dupicates?'
$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Remove duplicates"))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Do not remove duplicates"))
$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit program"))
$decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
if ($decision -eq 3) {
    $output_filename = Read-Host "Enter output file name: "
    Export-Csv -InputObject $csv -Path ./output/$output_filename
    Remove-Variable filename
    Remove-Variable output_filename
    exit 0
}
if ($decision -eq 0) {
    $Headers = (Get-Member -InputObject $csv[0] -MemberType NoteProperty).Name
    Write-Host ""
    Write-Host "Columns available in source file"
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $Column = $Headers[$i]
        Write-Host "$i. $Column"
    }
    $ind = Read-Host "Select the column you want to use for deleting duplicated values:"
    $column_remove_duplicates = $Headers[$ind]
    Write-Host ""
    Write-Host "$Column"
    $csv = ($csv | Sort-Object $column_remove_duplicates -Unique | Sort-Object $column_sort)
    

    Remove-Variable Headers 
    Remove-Variable Column
    Remove-Variable ind
}
# $output_filename = Read-Host "Enter output file name: "
# Export-Csv -InputObject $csv -Path ./output/$output_filename
$csv | Format-Table
# Variables Removal #
Remove-Variable filename
#Remove-Variable output_filename 
