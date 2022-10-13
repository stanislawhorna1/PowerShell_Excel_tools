Clear-Host
Write-Host ""
$filename = Read-Host "Enter filename: "

for (; ; ) {
    $title = 'Delimiter'
    $question = 'Select file type to import:'
    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&NXQL export', "Export from Nexthink uses ; as a delimiter"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Standard csv file', "Standard csv file uses , as a delimiter"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Custom', "Define custom delimiter sign"))
    $decision = $Host.UI.PromptForChoice("", $question, $choices, 0)
    
    if ($decision -eq 0) {
        $delimiter = "`t"
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
    $cols = (($csv | Get-Member | Where-Object -Property MemberType -eq NoteProperty).count)
    if ($cols -le 1) {
        $title = 'Delimiter'
        $question = 'Only one column is imported, would you like to change delimiter?'
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Change delimiter"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Continue with current selection"))
        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
        if ($decision -eq 1) {
            break
        }
    }
    else {
        break
    }
}

Clear-Host
Write-Host ""
Write-host " $cols columns imported"
$rows = ($csv.Length)
Write-host " $rows rows imported"
$counter = 0
for (; ; ) {
    $title = 'Function'
    $question = 'Select function:'
    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&View Table', "Show some table entries"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Sorting', "Sort data"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Duplicates Removal', "Remove duplicated values based on selected column"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Replace', "Replace character or word in all table entries in selected column"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Filter', "Filter based on custom conditions applied for particular column"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit program"))
    $menu_decision = $Host.UI.PromptForChoice($title, $question, $choices, ($choices.Count - 1))
    Clear-Host
    # /// DISPLAYING TABLE ///
    if ($menu_decision -eq 0) {
        $rows = $csv.Length
        Write-host ""
        Write-host "Input table has $length lines"
        if ($rows -gt 10) {
            $title = 'Display'
            $question = 'Do you want to display'
            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&First 10 entries', "Display first 10 table rows"))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Last 10 entries', "Display last 10 table rows"))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit program"))
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 2)
        }
        else {
            $decision = 0
        }
        if ($decision -eq 0) {
            $csv | Select-Object -First 10 | Format-table
        }
        if ($decision -eq 1) {
            $csv | Select-Object -Last 10 | Format-table
        }
    }
    # /// SORTING ///
    if ($menu_decision -eq 1) {
        Write-Host "SORTING"
        $counter = 1
        $Headers = (Get-Member -InputObject $csv[0] -MemberType NoteProperty).Name
        Write-Host ""
        Write-Host "Columns available in source file"
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $Column = $Headers[$i]
            Write-Host "$i. $Column"
        }
        $ind = Read-Host "Select the column you want to use for sorting"
        $column_sort = $Headers[$ind]
        Write-Host ""
        $title = 'Sorting Type'
        $question = 'Do you want to sort column as a text or date?'
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Text', "Sort as a string (* signs are allowed)"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Number', "Sort as a number"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Date', "Sort as date"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit function"))
        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
        if ($decision -eq 0) {
            $csv = ($csv | Sort-Object -Property $column_sort)
        }
        elseif ($decision -eq 1) {
            $csv = ($csv | Sort-Object { [int]$_.$column_sort })
        }
        elseif ($decision -eq 2) {



            
            for ($i = 0; $i -lt $csv.Count; $i++) {
                $csv[$i].$column_sort = (Get-Date -Day ($csv[$i].$column_sort.Split("T")[0]).Split(".")[0] -Month ($csv[$i].$column_sort.Split("T")[0]).Split(".")[1] -Year ($csv[$i].$column_sort.Split("T")[0]).Split(".")[2] -Hour ($csv[$i].$column_sort.Split("T")[1]).Split(":")[0] -Minute ($csv[$i].$column_sort.Split("T")[1]).Split(":")[1] -Second ($csv[$i].$column_sort.Split("T")[1]).Split(":")[2])
            }
            $csv = ($csv | Sort-Object -Property $column_sort -Descending)
        }
        Remove-Variable Headers 
        Remove-Variable Column
        Remove-Variable ind
    }
    # /// REMOVING DUPLICATES ///
    if ($menu_decision -eq 2) {
        $counter = 1
        $Headers = (Get-Member -InputObject $csv[0] -MemberType NoteProperty).Name
        Write-Host ""
        Write-Host "Columns available in source file"
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $Column = $Headers[$i]
            Write-Host "$i. $Column"
        }
        $ind = Read-Host "Select the column you want to use for deleting duplicated values"
        $column_remove_duplicates = $Headers[$ind]
        Write-Host ""
        $csv = ($csv | Sort-Object $column_remove_duplicates -Unique | Sort-Object $column_sort)
        Remove-Variable Headers 
        Remove-Variable Column
        Remove-Variable ind
    }
    # /// REPLACE ///
    if ($menu_decision -eq 3) {
        Write-Host "REPLACING"
        $counter = 1
        $Headers = (Get-Member -InputObject $csv[0] -MemberType NoteProperty).Name
        Write-Host ""
        Write-Host "Columns available in source file"
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $Column = $Headers[$i]
            Write-Host "$i. $Column"
        }
        $ind = Read-Host "Select the column you want to edit"
        $column_replace = $Headers[$ind]
        write-host ""
        $old_str = Read-Host "Find what "
        $new_str = Read-Host "Replace with "
        for ($i = 0; $i -lt $csv.Count; $i++) {
            $csv[$i].$column_replace = $csv[$i].$column_replace.Replace($old_str, $new_str)
        }
    }
    # /// FILTERING ///
    if ($menu_decision -eq 4) {
        Write-Host "FILTERING"
        $counter = 1
        $Headers = (Get-Member -InputObject $csv[0] -MemberType NoteProperty).Name
        Write-Host ""
        Write-Host "Columns available in source file"
        for ($i = 0; $i -lt $Headers.Count; $i++) {
            $Column = $Headers[$i]
            Write-Host "$i. $Column"
        }
        $ind = Read-Host "Select the column you want to use for filtering"
        $column_filter = $Headers[$ind]
        $title = 'Filtering Type'
        $question = 'Do you want to filter column as a text or date?'
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Text', "Sort as a string (* signs are allowed)"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Date', "Sort as date (After and before operators available)"))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Quit', "Quit program"))
        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 2)
        if ($decision -eq 0) {
            Write-Host "Remember to add * in a proper positions"
            $condition = Read-Host "Enter the condition you would like to apply"
            $csv = ($csv | Where-Object $column_filter -Like $condition)
        }
        elseif ($decision -eq 1) {
            $year = Read-Host "Enter year which you would like to use as filter"
            $title = 'Filtering Operator'
            $question = 'Do you want to list all entries before or after selected date?'
            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Before', "All entries before selected year will be selected"))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&After', "All entries after selected year will be selected"))
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 1)

            for ($i = 0; $i -lt $csv.Count; $i++) {
                $csv[$i].$column_filter = (Get-Date -Day ($csv[$i].$column_filter.Split("T")[0]).Split(".")[0] -Month ($csv[$i].$column_filter.Split("T")[0]).Split(".")[1] -Year ($csv[$i].$column_filter.Split("T")[0]).Split(".")[2] -Hour ($csv[$i].$column_filter.Split("T")[1]).Split(":")[0] -Minute ($csv[$i].$column_filter.Split("T")[1]).Split(":")[1] -Second ($csv[$i].$column_filter.Split("T")[1]).Split(":")[2])
            }
            
            if ($decision -eq 0) {
                $title = 'Filtering Operator'
                $question = 'Entered date should be included in selected entries?'
                $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Year previously provided will be included in selection"))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Year previously provided will be excluded from selection"))
                $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
                if ($decision -eq 0) {
                    $csv = ($csv | where-object { (Get-Date $_.$column_filter -Format yyyy) -le $year })
                }
                else {
                    $csv = ($csv | where-object { (Get-Date $_.$column_filter -Format yyyy) -lt $year })
                }
            }
            else {
                $title = 'Filtering Operator'
                $question = 'Entered date should be included in selected entries?'
                $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Year previously provided will be included in selection"))
                $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Year previously provided will be excluded from selection"))
                $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
                if ($decision -eq 0) {
                    $csv = ($csv | where-object { (Get-Date $_.$column_filter -Format yyyy) -ge $year })
                }
                else {
                    $csv = ($csv | where-object { (Get-Date $_.$column_filter -Format yyyy) -gt $year })
                }
            }
            
        }
    }
    # /// EXIT ///
    if ($menu_decision -eq ($choices.Count - 1)) {
        if ($counter -ge 1) {
            $title = 'SAVE'
            $question = 'Do you want to save modified table?'
            $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Save modified Table in a new csv file with ; as a delimiter"))
            $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "Discard changes"))
            $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
            if ($decision -eq 0) {
                for (; ; ) {
                    $output_filename = Read-Host "Enter output file name: "
                    try {
                        $csv | Export-Csv -Path ./output/$output_filename -Delimiter $delimiter
                        Remove-Variable output_filename
                        exit 0
                    }
                    catch {
                        { Provide correct output file name }
                    }
                }
            }
        }
        Remove-Variable filename
        exit 0
    }
}


