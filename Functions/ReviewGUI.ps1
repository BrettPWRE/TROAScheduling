Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Global Variable List to contain all items and their list names
$global:ItemReviewList = @()


Function RefreshButtonClick(){
    $RefreshButton.Enabled = $false
    $Table = @()

    if($ModifiedTypeRadio.Checked){
        $ModDate = [datetime]::ParseExact($DateInput.Text,'M/d/yyyy',$null)
        $ModDate_str = $DateInput.Text
        foreach($entry in $ItemReviewList){
            $RecentlyModified = 0
            foreach($item in $entry['Items']){
                if([TimeZoneInfo]::ConvertTimeBySystemTimeZoneId($item['Modified'],'Mountain Standard Time') -ge $ModDate -and $item['SubmittedStatus'] -ne 'Integrated'){$RecentlyModified += 1}
            } 
            $hash = @{'List Name' = $entry['List']; 'Recently Modified' = $RecentlyModified}
            $Table += $hash        
        }

        $form.TopMost = $false
        $Table | ForEach {[PSCustomObject]$_} | Select-Object -Property 'List Name','Recently Modified'  | Out-GridView -Title "Recently modified draft items, on or after $ModDate_str" -Wait
        $form.TopMost = $true
        
        

    }
    elseif($CountTypeRadio.Checked){
        $IntDate = [datetime]::ParseExact($DateInput.Text,'M/d/yyyy',$null)
        $IntDate_str = $DateInput.Text
        foreach($entry in $ItemReviewList){
            $numDraft = 0
            $numIntegrated = 0
            

            foreach($item in $entry['Items']){
                if($item['SubmittedStatus'] -eq "Integrated" -and $item['SubmittedDate'].ToString('M/d/yyyy') -eq $IntDate_str){$numIntegrated += 1}
                elseif($item['SubmittedStatus'] -ne "Integrated" -and $item['SubmittedStatus'] -ne $null){$numDraft += 1}
                
            }
            $hash = @{'List Name' = $entry['List']; 'Draft' = $numDraft; 'Integrated' = $numIntegrated; 'Change' = $numDraft - $numIntegrated}
            $Table += $hash 
        }
        $form.TopMost = $false
        $Table | ForEach {[PSCustomObject]$_} | Select-Object -Property 'List Name','Draft','Integrated','Change'  | Out-GridView -Title "Draft and Integrated Items, Integrated on $IntDate_str" -Wait
        $form.TopMost = $true

    }
    elseif($UnsubmittedTypeRadio.Checked){
        foreach($entry in $ItemReviewList){
            $numUnsubmitted = 0
            foreach($item in $entry['Items']){
                if($item['SubmittedStatus'] -ne "Integrated" -and $item['SubmittedStatus'] -ne $null -and $item['Party_x0020_Status'] -ne "Submitted"){$numUnsubmitted += 1}
            }
            $hash = @{'List Name' = $entry['List']; 'Unsubmitted Draft Items' = $numUnsubmitted}
            $Table += $hash 

        }
        $form.TopMost = $false
        $Table | ForEach {[PSCustomObject]$_} | Select-Object -Property 'List Name','Unsubmitted Draft Items'  | Out-GridView -Title "Items in Draft View Without Submitted Party Status" -Wait
        $form.TopMost = $true
    }
    elseif($ChangeModTypeRadio.Checked){
        $ListName = $ListInput.Text
        $NewModDate = [datetime]::ParseExact($DateInput.Text,'M/d/yyyy',$null)

        if($TestModeRadio.Checked){
            $Ctn = $PWRE_Ctn
            $form.topmost = $false
            $IDsToEdit = Select-ItemsToEdit -Mode "Test" -ListName $ListName -Action "Edit Modified"
            $form.topmost = $true
        }elseif($USWMModeRadio.Checked){
            $Ctn = $USWM_Ctn
            $form.topmost = $false
            $IDsToEdit = Select-ItemsToEdit -Mode "USWM" -ListName $ListName -Action "Edit Modified"
            $form.topmost = $true
        }

        foreach($id in $IDsToEdit){
            $EditedItem = Set-PnPListItem -List $ListName -Identity $id -Connection $Ctn -Values @{'Modified' = $NewModDate}
        }
        Write-Host "----- Done Editing Modified Date(s) -----"
    }
    CheckRefreshButton
}

Function ReloadButtonClick(){
    $RefreshButton.Enabled = $false

    #Load items from each list, depending on form selections
    if($TestModeRadio.Checked){
        ."$PSScriptRoot\PWRELogin.ps1"
        $Ctn = $PWRE_Ctn
    }elseif($USWMModeRadio.Checked){
        $Ctn = $USWM_Ctn
    }
    
    $ScheduleLists = Build-ListOfScheduleLists -Connection $Ctn
    $numScheduleLists = $ScheduleLists.Length
    $global:ItemReviewList = @()

    $listctr = 0
    foreach($list in $ScheduleLists){
        $listProgress = [math]::Round($listctr/$numScheduleLists,3)*100
        $list_name = $list.Title
        Write-Progress -Activity "Reading List: $list_name" -Status "$listctr of $numScheduleLists Completed" -PercentComplete $listProgress
        $global:ItemReviewList += @{'List' = $list_name;'Items' = (Get-PnPListItem -List $list_name -Connection $Ctn)}
        $listctr += 1
    }
    Write-Progress -Activity "Reading List: $list.Title" -Status "$listctr of $numScheduleLists Completed" -Completed
    Write-Host "----- Ready to Review Lists -----"



    CheckRefreshButton
}
Function UnsubmittedClick(){
    $DateInput.Enabled = $false
    $DateLabel.Enabled = $false
    $ListLabel.Enabled = $false
    $ListInput.Enabled = $false
    $RefreshButton.Text = 'Show Results'
    CheckRefreshButton
}

Function CountClick(){
    $DateInput.Enabled = $true
    $DateLabel.Enabled = $true
    $ListLabel.Enabled = $false
    $ListInput.Enabled = $false
    $DateLabel.Text = 'Enter the Submitted Date below (m/d/yyyy)'
    $RefreshButton.Text = 'Show Results'
    CheckRefreshButton
}

Function ModifiedClick(){
    $DateInput.Enabled = $true
    $DateLabel.Enabled = $true
    $ListLabel.Enabled = $false
    $ListInput.Enabled = $false
    $DateLabel.Text = 'Enter the Recently Modified Date below (m/d/yyyy)'
    $RefreshButton.Text = 'Show Results'
    CheckRefreshButton
}

Function ChangeModifiedClick(){
    $DateInput.Enabled = $true
    $DateLabel.Enabled = $true
    $ListLabel.Enabled = $true
    $ListInput.Enabled = $true
    $DateLabel.Text = 'Specify a new Modified Date below (m/d/yyyy)'
    $RefreshButton.Text = 'Select Items'
    CheckRefreshButton
}

Function CheckRefreshButton(){
    $ItemsLoaded = ($ItemReviewList.Length -gt 0)
    $DateHasInput = ($DateInput.Text -ne "")
    $ListHasInput = ($ListInput.Text -ne "")

    # specify conditions for each review type to allow refresh button to be clicked
    if($ModifiedTypeRadio.Checked -and $ItemsLoaded -and $DateHasInput){
        $RefreshButton.Enabled = $true
    }elseif($CountTypeRadio.Checked -and $ItemsLoaded -and $DateHasInput){
        $RefreshButton.Enabled = $true
    }elseif($UnsubmittedTypeRadio.Checked -and $ItemsLoaded){
        $RefreshButton.Enabled = $true
    }elseif($ChangeModTypeRadio.Checked -and $DateHasInput -and $ListHasInput){
        $RefreshButton.Enabled = $true
    }else{
        $RefreshButton.Enabled = $false
    }

    if($ChangeModTypeRadio.Checked -and $ListInput.Text -ne "" -and $DateInput.Text -ne ""){
        $RefreshButton.Enabled = $true
    }elseif($ItemReviewList.Length -le 0){
        $RefreshButton.Enabled = $false
    }elseif($UnsubmittedTypeRadio.Checked){
        $RefreshButton.Enabled = $true
    }elseif($DateInput.Text -eq ""){
        $RefreshButton.Enabled = $false
    }else{
        $RefreshButton.Enabled = $true
    }
    
}
Function Get-ReviewParameters(){
    
    # Build form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'TROA Schedule List Review'
    $form.Size = '480,400'
    $form.StartPosition = 'CenterScreen'
    
    # Add Button to refresh output
    $RefreshButton = New-Object System.Windows.Forms.Button
    $RefreshButton.Location = '25,325'
    $RefreshButton.Size = '125,23'
    $RefreshButton.Text = 'Show Results'
    $RefreshButton.Enabled = $false
    $RefreshButton.DialogResult = [System.Windows.Forms.DialogResult]::None
    $form.Controls.Add($RefreshButton)

    #Add Cancel Button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '150,325'
    $CancelButton.Size = '125,23'
    $CancelButton.Text = 'Exit Review'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    #Add button to reload data
    $ReloadButton = New-Object System.Windows.Forms.Button
    $ReloadButton.Location = '320,60'
    $ReloadButton.Size = '100,23'
    $ReloadButton.Text = 'Load Data'
    $ReloadButton.DialogResult = [System.Windows.Forms.DialogResult]::None
    $form.Controls.Add($ReloadButton)

    # Add group of of radio buttons to select Review Mode (Test site or Watermaster site)
    $ModeGroup = New-Object System.Windows.Forms.GroupBox
    $ModeGroup.Location = '25,15'
    $ModeGroup.size = '400,75'
    $ModeGroup.text = "Select the Sharepoint site to review schedule lists on:"
    
    $TestModeRadio = New-Object System.Windows.Forms.RadioButton
    $TestModeRadio.Location = '20,20'
    $TestModeRadio.size = '350,20'
    $TestModeRadio.Checked = $false
    $TestModeRadio.Text = "PWRE - Test"
 
    $USWMModeRadio = New-Object System.Windows.Forms.RadioButton
    $USWMModeRadio.Location = '20,45'
    $USWMModeRadio.size = '350,20'
    $USWMModeRadio.Checked = $true
    $USWMModeRadio.Text = "USWM - Official"

    $ModeGroup.Controls.AddRange(@($TestModeRadio,$USWMModeRadio))
    $form.Controls.Add($ModeGroup)

    # Add group of radio buttons to select Review type
    $TypeGroup = New-Object System.Windows.Forms.GroupBox
    $TypeGroup.Location = '25,100'
    $TypeGroup.size = '400,120'
    $TypeGroup.text = "Select the type of review to perform:"
    
    $ModifiedTypeRadio = New-Object System.Windows.Forms.RadioButton
    $ModifiedTypeRadio.Location = '20,20'
    $ModifiedTypeRadio.size = '350,20'
    $ModifiedTypeRadio.Checked = $true 
    $ModifiedTypeRadio.Text = "Count Recently Modified Items"
 
    $CountTypeRadio = New-Object System.Windows.Forms.RadioButton
    $CountTypeRadio.Location = '20,45'
    $CountTypeRadio.size = '350,20'
    $CountTypeRadio.Checked = $false
    $CountTypeRadio.Text = "Count Integrated and Draft Items"

    $UnsubmittedTypeRadio = New-Object System.Windows.Forms.RadioButton
    $UnsubmittedTypeRadio.Location = '20,70'
    $UnsubmittedTypeRadio.size = '350,20'
    $UnsubmittedTypeRadio.Checked = $false
    $UnsubmittedTypeRadio.Text = "Check for Unsubmitted Draft Items"

    $ChangeModTypeRadio = New-Object System.Windows.Forms.RadioButton
    $ChangeModTypeRadio.Location = '20,95'
    $ChangeModTypeRadio.size = '350,20'
    $ChangeModTypeRadio.Checked = $false
    $ChangeModTypeRadio.Text = "Edit Modified Date of specific Item(s)"

    $TypeGroup.Controls.AddRange(@($ModifiedTypeRadio,$CountTypeRadio,$UnsubmittedTypeRadio,$ChangeModTypeRadio))
    $form.Controls.Add($TypeGroup)

    # Add input box for date
    $DateLabel = New-Object System.Windows.Forms.Label
    $DateLabel.Location = '25,225'
    $DateLabel.Size = '350,20'
    $DateLabel.Text = 'Enter the Recently Modified Date below (m/d/yyyy)'
    $DateLabel.Enabled = $true
    $form.Controls.Add($DateLabel)

    $DateInput = New-Object System.Windows.Forms.TextBox
    $DateInput.Location = '25,245'
    $DateInput.Size = '400,20'
    $DateInput.Enabled = $true
    $form.Controls.Add($DateInput)

    # Add input box for list
    $ListLabel = New-Object System.Windows.Forms.Label
    $ListLabel.Location = '25,270'
    $ListLabel.Size = '350,20'
    $ListLabel.Text = 'Enter the name of a single List below'
    $ListLabel.Enabled = $false
    $form.Controls.Add($ListLabel)

    $ListInput = New-Object System.Windows.Forms.TextBox
    $ListInput.Location = '25,290'
    $ListInput.Size = '400,20'
    $ListInput.Enabled = $false
    $form.Controls.Add($ListInput)
    
    #Button Events
    $RefreshButton.Add_Click({RefreshButtonClick})
    $ReloadButton.Add_Click({ReloadButtonClick})
    $ModifiedTypeRadio.Add_Click({ModifiedClick})
    $CountTypeRadio.Add_Click({CountClick})
    $UnsubmittedTypeRadio.Add_Click({UnsubmittedClick})
    $ChangeModTypeRadio.Add_Click({ChangeModifiedClick})
    $DateInput.Add_TextChanged({CheckRefreshButton})
    $ListInput.Add_TextChanged({CheckRefreshButton})
    
    # Show form, get response (OK or Cancel)
    $form.TopMost = $true
    $Result = $form.ShowDialog()

}

