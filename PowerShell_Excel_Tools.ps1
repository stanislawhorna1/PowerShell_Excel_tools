$filename = Read-Host "Enter filename: "

for (; ; ) {
    $title = 'Delimiter'
    $question = 'Select file type to import:'
    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Nexthink', "Export from Nexthink uses ; as a delimiter"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Coma Separated Valuse', "Standard csv file uses , as a delimiter"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Special', "Define custom delimiter sign"))
    $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
    if ($decision -eq 0) {
        $delimiter = ";"
    }
    elseif ($decision -eq 1) {
        $delimiter = ","
    }
    else {
        $delimiter = Read-Host "Enter custom delimiter "
    }
    try {
        $csv = Import-Csv ./Output/$filename -Delimiter $delimiter
    }
    catch {
        { Impossible to read the file }
        exit 1
    }
    if (($csv | Get-Member | Where-Object -Property MemberType -eq NoteProperty).count -le 1) {
        $title = 'Delimiter'
        $question = 'Only one column is imported, would you like to change delimiter?'
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Change delimiter"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Continue with current selection"))
        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
        if ($decision -eq 1) {
            break
        }
    }else{
        break
    }
}

$counter = 0
for (; ; ) {
    $title = 'Function'
    $question = 'Select function:'
    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Sorting', "Sort data"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Removing Duplicates', "Remove duplicated values based on selected column"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Display Table', "Show some table entries"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit program"))
    $decision = $Host.UI.PromptForChoice($title, $question, $choices, ($choices.Count - 1))
    
    # /// SORTING ///
    if ($decision -eq 0) {
        $counter = 1
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
    # /// REMOVING DUPLICATES ///
    if ($decision -eq 1) {
        $counter = 1
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

    # /// DISPLAYING TABLE ///
    if ($decision -eq 2) {
        $title = 'Display'
        $question = 'Do you want to display'
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&First 10 entries', "Display first 10 table rows"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Last 10 entries', "Display last 10 table rows"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit program"))
        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 2)
        if ($decision -eq 0) {
            $csv | Select-Object -First 10 | Format-table
        }
        if ($decision -eq 1) {
            $csv | Select-Object -Last 10 | Format-table
        }
    }

    # /// EXIT ///
    if ($decision -eq ($choices.Count - 1)) {
        if ($counter -ge 1) {
            $title = 'SAVE'
            $question = 'Do you want to save modified table?'
            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Save modified Table in a new csv file with ; as a delimiter"))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Discard changes"))
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
            if ($decision -eq 0) {
                $output_filename = Read-Host "Enter output file name: "
                Export-Csv -InputObject $csv -Path ./output/$output_filename
                Remove-Variable output_filename
            }
        }
        Remove-Variable filename
        exit 0
    }
}


