Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function CheckScopeStatus(){
    if($SingleScopeRadio.Checked){
        $SingleListLabel.Enabled = $true
        $SingleListInput.Enabled = $true
    }else{
        $SingleListLabel.Enabled = $false
        $SingleListInput.Enabled = $false
        $SingleListInput.Text = $null
    }
    CheckOKStatus
}

Function CheckOKStatus(){
    #Disable the OK button if Single List scope is selected and the Single List input box is empty
    if($SingleScopeRadio.Checked-and $SingleListInput.Text -eq ''){
        $OKButton.Enabled = $false
    }else{
        $OKButton.Enabled = $true
    }

}

Function Get-CopyTestSiteParameters(){
    
    # Build form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Copy TROA Scheduling Lists to Test Site'
    $form.Size = '480,225'
    $form.StartPosition = 'CenterScreen'
    
    # Add OK Button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '25,150'
    $OKButton.Size = '75,23'
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    #Add Cancel Button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '100,150'
    $CancelButton.Size = '75,23'
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)


    # Add group of radio buttons to select Scope (All or Single)
    $ScopeGroup = New-Object System.Windows.Forms.GroupBox
    $ScopeGroup.Location = '25,15'
    $ScopeGroup.size = '400,75'
    $ScopeGroup.text = "Select the scope of of lists to copy:"
    
    $AllScopeRadio = New-Object System.Windows.Forms.RadioButton
    $AllScopeRadio.Location = '20,20'
    $AllScopeRadio.size = '350,20'
    $AllScopeRadio.Checked = $true 
    $AllScopeRadio.Text = "All Lists"

    $SingleScopeRadio = New-Object System.Windows.Forms.RadioButton
    $SingleScopeRadio.Location = '20,45'
    $SingleScopeRadio.size = '350,20'
    $SingleScopeRadio.Checked = $false
    $SingleScopeRadio.Text = "Single List"

    $ScopeGroup.Controls.AddRange(@($AllScopeRadio,$SingleScopeRadio))
    $form.Controls.Add($ScopeGroup)

    # Add input box for single list name
    $SingleListLabel = New-Object System.Windows.Forms.Label
    $SingleListLabel.Location = '25,100'
    $SingleListLabel.Size = '350,20'
    $SingleListLabel.Text = 'Enter the name of a single list below:'
    $SingleListLabel.Enabled = $false
    $form.Controls.Add($SingleListLabel)

    $SingleListInput = New-Object System.Windows.Forms.TextBox
    $SingleListInput.Location = '25,120'
    $SingleListInput.Size = '400,20'
    $SingleListInput.Enabled = $false
    $form.Controls.Add($SingleListInput)
    

    #Radio Button events
    $AllScopeRadio.Add_Click({CheckScopeStatus})
    $SingleScopeRadio.Add_Click({CheckScopeStatus})

    #Single List name input box event
    $SingleListInput.Add_TextChanged({CheckOKStatus})

    # Show form, get response (OK or Cancel)
    $form.TopMost = $true
    $Result = $form.ShowDialog()

    if($Result -eq 'OK'){
        # Return variables in a hashtable if user presses ok
        if($AllScopeRadio.Checked)
        {
            $Scope = "All"
            $SingleList = $null
        }
        elseif($SingleScopeRadio.Checked)
        {
            $Scope = "Single"
            $SingleList = $SingleListInput.Text
        }


        
    }else{ # If user presses x or cancel
        $Scope = $null
        $SingleList = $null
    }
    $Response = @{
        'Scope' = $Scope;
        'SingleList' = $SingleList
    }
    Return $Response
}