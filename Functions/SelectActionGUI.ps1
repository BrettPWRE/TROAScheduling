Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Function Select-TROASchedulingAction(){
    
    # Build form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'TROA Scheduling Actions'
    $form.Size = '480,200'
    $form.StartPosition = 'CenterScreen'
    
    # Add Next Button
    $OKButton = New-Object System.Windows.Forms.Button
    $OKButton.Location = '25,125'
    $OKButton.Size = '75,23'
    $OKButton.Text = 'Next'
    $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.AcceptButton = $OKButton
    $form.Controls.Add($OKButton)

    #Add Cancel Button
    $CancelButton = New-Object System.Windows.Forms.Button
    $CancelButton.Location = '100,125'
    $CancelButton.Size = '75,23'
    $CancelButton.Text = 'Exit'
    $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.CancelButton = $CancelButton
    $form.Controls.Add($CancelButton)


    # Add group of radio buttons to select action
    $ActionGroup = New-Object System.Windows.Forms.GroupBox
    $ActionGroup.Location = '25,15'
    $ActionGroup.size = '400,105'
    $ActionGroup.text = "Select the TROA Scheduling Action to perform:"
    
    $IntegrationRadio = New-Object System.Windows.Forms.RadioButton
    $IntegrationRadio.Location = '20,20'
    $IntegrationRadio.size = '350,20'
    $IntegrationRadio.Checked = $false
    $IntegrationRadio.Text = "Schedule Integration"

    $ReviewRadio = New-Object System.Windows.Forms.RadioButton
    $ReviewRadio.Location = '20,45'
    $ReviewRadio.size = '350,20'
    $ReviewRadio.Checked = $true
    $ReviewRadio.Text = "Schedule List Review"

    $CopyRadio = New-Object System.Windows.Forms.RadioButton
    $CopyRadio.Location = '20,70'
    $CopyRadio.size = '350,20'
    $CopyRadio.Checked = $false
    $CopyRadio.Text = "Copy Schedule List(s) to Test Site"

    $ActionGroup.Controls.AddRange(@($IntegrationRadio,$ReviewRadio,$CopyRadio))
    $form.Controls.Add($ActionGroup)



    # Show form, get response (OK or Cancel)
    $form.TopMost = $true
    $Result = $form.ShowDialog()

    if($Result -eq 'OK'){
        # Return variables in a hashtable if user presses ok
        if($IntegrationRadio.Checked)
        {
            $Action = "Integration"
        }
        elseif($ReviewRadio.Checked)
        {
            $Action = "Review"
        }elseif($CopyRadio.Checked)
        {
            $Action = "Copy"
        }



        
    }else{ # If user presses x or cancel
        $Action = $null
    }

    Return $Action
}
