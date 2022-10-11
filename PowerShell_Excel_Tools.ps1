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
    $columnsort = $Headers[$ind]
    Write-Host ""
    if ($columnsort.ToLower() -like "*time*" -or $columnsort.ToLower() -like "*date*") {
        $csv = ($csv | Sort-Object { Get-Date $_.$columnsort } -Descending)
        #$csv | Format-Table
    }
    elseif ($columnsort.ToLower() -eq "id" -or $columnsort.ToLower() -like "*num*") {
        $csv = ($csv | Sort-Object { [int]$_.$columnsort })
        #$csv | Format-Table
    }
    else {
        $csv = ($csv | Sort-Object -Property $columnsort)
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
    $Column = $Headers[$ind]
    Write-Host ""
    Write-Host "$Column"
    $csv = ($csv | Sort-Object $Column -Unique | Sort-Object $columnsort)
    

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
