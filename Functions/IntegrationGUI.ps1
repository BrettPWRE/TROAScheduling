Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function CheckScopeStatus(){
    if($SingleScopeRadio.Checked -or $SpecificScopeRadio.Checked){
        $SingleListLabel.Enabled = $true
        $SingleListInput.Enabled = $true
        $PartyListLabel.Enabled = $false
        $PartyListBox.Enabled=$false
        $PartyListBox.ClearSelected()
    }elseif($PartyScopeRadio.Checked){
        $SingleListLabel.Enabled = $false
        $SingleListInput.Enabled = $false
        $PartyListLabel.Enabled = $true
        $PartyListBox.Enabled=$true
        $SingleListInput.Text = $null
    }else{
        $SingleListLabel.Enabled = $false
        $SingleListInput.Enabled = $false
        $PartyListLabel.Enabled = $false
        $PartyListBox.Enabled=$false
        $PartyListBox.ClearSelected()
        $SingleListInput.Text = $null
    }
    CheckOKStatus
}

Function CheckOKStatus(){
    #Disable the OK button if Single List scope is selected and the Single List input box is empty, or
    # if the Party Scope is selected and no party has been selected
    if($PartyScopeRadio.Checked -and $PartyListBox.SelectedIndex -eq -1){
        $OKButton.Enabled = $false
    }elseif(($SingleScopeRadio.Checked -or $SpecificScopeRadio.Checked) -and $SingleListInput.Text -eq ''){
        $OKButton.Enabled = $false
    }else{
        $OKButton.Enabled = $true
    }

}

Function Get-IntegrationParameters(){
    
    # Build form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'TROA Schedule Integration Setup'
    $form.Size = '480,495'
    $form.StartPosition = 'CenterScreen'
    
    # Add OK Button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '25,420'
    $OKButton.Size = '75,23'
    $OKButton.Text = 'OK'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    #Add Cancel Button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '100,420'
    $CancelButton.Size = '75,23'
    $CancelButton.Text = 'Cancel'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)

    # Add group of of radio buttons to select Integration Mode (Test site or Watermaster site)
    $ModeGroup = New-Object System.Windows.Forms.GroupBox
    $ModeGroup.Location = '25,15'
    $ModeGroup.size = '400,75'
    $ModeGroup.text = "Select the Sharepoint site to integrate schedule items on:"
    
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

    # Add group of radio buttons to select Integration Scope (All, Party or Single)
    $ScopeGroup = New-Object System.Windows.Forms.GroupBox
    $ScopeGroup.Location = '25,100'
    $ScopeGroup.size = '400,125'
    $ScopeGroup.text = "Select the scope of integration to perform:"
    
    $AllScopeRadio = New-Object System.Windows.Forms.RadioButton
    $AllScopeRadio.Location = '20,20'
    $AllScopeRadio.size = '350,20'
    $AllScopeRadio.Checked = $true 
    $AllScopeRadio.Text = "All Lists"
 
    $PartyScopeRadio = New-Object System.Windows.Forms.RadioButton
    $PartyScopeRadio.Location = '20,45'
    $PartyScopeRadio.size = '350,20'
    $PartyScopeRadio.Checked = $false
    $PartyScopeRadio.Enabled = $false
    $PartyScopeRadio.Text = "By Party"

    $SingleScopeRadio = New-Object System.Windows.Forms.RadioButton
    $SingleScopeRadio.Location = '20,70'
    $SingleScopeRadio.size = '350,20'
    $SingleScopeRadio.Checked = $false
    $SingleScopeRadio.Text = "Single List"

    $SpecificScopeRadio = New-Object System.Windows.Forms.RadioButton
    $SpecificScopeRadio.Location = '20,95'
    $SpecificScopeRadio.size = '350,20'
    $SpecificScopeRadio.Checked = $false
    $SpecificScopeRadio.Text = "Specific item(s)"

    $ScopeGroup.Controls.AddRange(@($AllScopeRadio,$PartyScopeRadio,$SingleScopeRadio,$SpecificScopeRadio))
    $form.Controls.Add($ScopeGroup)

    # Add input box for single list name
    $SingleListLabel = New-Object System.Windows.Forms.Label
    $SingleListLabel.Location = '25,230'
    $SingleListLabel.Size = '350,20'
    $SingleListLabel.Text = 'Enter the name of a single list below:'
    $SingleListLabel.Enabled = $false
    $form.Controls.Add($SingleListLabel)

    $SingleListInput = New-Object System.Windows.Forms.TextBox
    $SingleListInput.Location = '25,250'
    $SingleListInput.Size = '400,20'
    $SingleListInput.Enabled = $false
    $form.Controls.Add($SingleListInput)
    
    #Add selection box for single party
    $PartyListLabel = New-Object System.Windows.Forms.Label
    $PartyListLabel.Location = '25,280'
    $PartyListLabel.Size = '300,20'
    $PartyListLabel.Text = 'Select a Party below:'
    $PartyListLabel.Enabled = $false
    $form.Controls.Add($PartyListLabel)

    $PartyListBox = New-Object System.Windows.Forms.Listbox
    $PartyListBox.Location = '25,300'
    $PartyListBox.Size = '400,20'
    $PartyListBox.SelectionMode = 'One'
    $PartyListBox.Enabled = $false

    $Parties = ('Administrator','California','Fernley','Pyramid Lake Paiute Tribe','Reno, Sparks, Washoe County','TMWA','US','Water Master')
    foreach($Party in $Parties){[void] $PartyListBox.Items.Add($Party)}

    $PartyListBox.Height = 120
    $form.Controls.Add($PartyListBox)

    #Radio Button events
    $AllScopeRadio.Add_Click({CheckScopeStatus})
    $PartyScopeRadio.Add_Click({CheckScopeStatus})
    $SingleScopeRadio.Add_Click({CheckScopeStatus})
    $SpecificScopeRadio.Add_Click({CheckScopeStatus})

    #Single List name input box event
    $SingleListInput.Add_TextChanged({CheckOKStatus})

    #Party ListBox change event
    $PartyListBox.Add_SelectedIndexChanged({CheckOKStatus})
    
    # Show form, get response (OK or Cancel)
    $form.TopMost = $true
    $Result = $form.ShowDialog()

    if($Result -eq 'OK'){
        # Return variables in a hashtable if user presses ok
        if($TestModeRadio.Checked){$Mode = "Test"}
        elseif($USWMModeRadio.Checked){$Mode = "USWM"}

        if($AllScopeRadio.Checked)
        {
            $Scope = "All"
            $Party = $null
            $SingleList = $null
            $SelectedItems = $null
        }
        elseif($PartyScopeRadio.Checked)
        {
            $Scope = "Party"
            $Party = $PartyListBox.SelectedItem
            $SingleList = $null
            $SelectedItems = $null
        }
        elseif($SingleScopeRadio.Checked)
        {
            $Scope = "Single"
            $Party = $null
            $SingleList = $SingleListInput.Text
            $SelectedItems = $null
        }
        elseif($SpecificScopeRadio.Checked)
        {
            $TempItems = Select-ItemsToEdit -Mode $Mode -List $SingleListInput.Text -Action "Integrate"
            If($TempItems -ne $null){
                $Scope = "Specific"
                $Party = $null
                $SingleList = $SingleListInput.Text
                $SelectedItems = $TempItems
            }else{
                $Mode = "Cancel"
                $Scope = $null
                $Party = $null
                $SingleList = $null
                $SelectedItems = $null
            }
        }


        
    }else{ # If user presses x or cancel
        $Mode = "Cancel"
        $Scope = $null
        $Party = $null
        $SingleList = $null
        $SelectedItems = $null
    }
    $Response = @{
        'Mode' = $Mode;
        'Scope' = $Scope;
        'Party' = $Party;
        'SingleList' = $SingleList
        'SelectedItems' = $SelectedItems
    }
    Return $Response
}
