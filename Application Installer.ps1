# © Copyright 2022, Jacob Searcy, All rights reserved.

function Install-MISCapp {

    [CmdletBinding()]

    Param(
        [String] $ComputerName,
        [String] $AppName
    )

    Begin{
        If ($AppName -eq "Office365") {
            $source_files = "\\misc\entmgmt\Applications\SCCM\Microsoft\Microsoft 365 Apps for Enterprise"
        }
        If ($AppName -eq "Office2010") {
            $source_files = "\\wnis01\ftp$\-=DepartmentApps=-\Microsoft Office Professional Plus 2010"
        }
        If ($AppName -eq "Dymo") {
            $source_files = "\\misc\entmgmt\Applications\SCCM\DYMO"
        }
        If ($AppName -eq "ScreenSaver") {
            $source_files = "\\wnis01\ftp$\-=DepartmentApps=-\ScreenSaver - NoTimeout"
        }
    }

    Process{
        If(Test-Connection $ComputerName -Count 1 -quiet){
            icacls "\\$ComputerName\c$\temp" /grant Everyone:F
            New-Item -Path "\\$ComputerName\c$\temp" -Name "$AppName" -ItemType "directory" -ErrorAction SilentlyContinue
            icacls "\\$ComputerName\c$\temp\$AppName" /grant Everyone:F
            Write-Host "Copying $AppName Install Files to $ComputerName . . ."
            Robocopy $source_files "\\$ComputerName\c$\temp\$AppName" *.* /E /S
            Write-Host "Copy complete"
            Write-Host "Launching $AppName Install . . ."
            psexec -s -d -i 1 -w c:\temp\$AppName \\$ComputerName c:\temp\$AppName\Deploy-Application.exe
            Write-Host "$AppName Install started on $ComputerName!" -ForegroundColor Green
            Write-Host $null
            Write-Host "Installing $AppName . . ."
            Write-Host $null
            $check = pslist \\$ComputerName Deploy-Application
            $check
            $check[5,6,7,8]
            $null
        }else{
            Write-Host $null
            Write-Host "$ComputerName is Offline" -ForegroundColor Red
            Write-Host $null
            $offline = echo $offline $ComputerName
        }
        Write-Host "Checking $AppName Install Status . . ."
        Write-Host $null
        $header = $check[7]
        $count = $comps.Count
        $continue = $true
        While ($continue -eq $true) {
            If (Test-Connection $ComputerName -Count 1 -quiet){
                $check = pslist \\$ComputerName Deploy-Application
                If ($check -match "was not found") {
                    Write-Host $null
                    Write-Host "$AppName Install finished on $ComputerName!" -ForegroundColor Green
                    $continue = $false
                }
            }
        }
        Remove-Item -LiteralPath "\\$ComputerName\c$\temp\$AppName" -Force -Recurse
        If ($AppName -eq "Microsoft Office365") {
            $source = "\\wnis01\ftp$\Shortcuts"
            $destination = "\\$ComputerName\c$\Users\Public\Desktop"
            Write-Host $null
            Write-Host "Copying 365 Shortcuts to Public Desktop . . ."
            Write-Host $null
            copy "$source\Word.lnk" $destination
            copy "$source\Excel.lnk" $destination
            copy "$source\OneNote.lnk" $destination
            copy "$source\Outlook.lnk" $destination
            copy "$source\PowerPoint.lnk" $destination
            Write-Host "$ComputerName has Shortcuts!" -ForegroundColor Green
            Write-Host $null
        }
        Write-Host $null
        Write-Host "Process Complete. Press ENTER to Exit."
        pause
    }

    End{}
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName PresentationFramework
$apps_list = @()
$SCCMList = (Get-WmiObject -Class CCM_Application -Namespace "root\ccm\clientSDK").Name
<#If ($SCCMList -notmatch '\w') {
    Write-Host $null
    Write-Host "SCCM did not return any results. This tool will only be able to install Office 365, Office 2010, or Dymo" -ForegroundColor Red
    Write-Host $null
    pause
}#>
$OtherApps = "Dymo","Microsoft Office365","Microsoft Office365 Web Apps","Microsoft Office 2010"
$apps_list = $SCCMList + $OtherApps

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$search_form = New-Object System.Windows.Forms.Form
$search_form.Text = "Applications Installer"
$search_form.AutoSize = $true
$search_form.StartPosition = 'CenterScreen'
$search_form.TopMost = $true

$image = [System.Drawing.Image]::FromFile("$PSScriptRoot\Files\search_icon.jpeg")

#ComputerName List
$Computer_label = New-Object System.Windows.Forms.Label
$Computer_label.Text = "Enter Computer Name:"
$Computer_label.Location = New-Object System.Drawing.Point(5,8)
$Computer_label.Size = New-Object System.Drawing.Size(150,20)
$search_form.Controls.Add($Computer_label)

$ComputerList_label = New-Object System.Windows.Forms.Label
$ComputerList_label.Text = "Computer Names:"
$ComputerList_label.Location = New-Object System.Drawing.Point(5,70)
$ComputerList_label.Size = New-Object System.Drawing.Size(100,21)
$search_form.Controls.Add($ComputerList_label)

$Computer_bar = New-Object System.Windows.Forms.TextBox
$Computer_bar.Location = New-Object System.Drawing.Size(5,28)
$Computer_bar.Size = New-Object System.Drawing.Size(200,21)
$search_form.Controls.Add($Computer_bar)
$Computer_bar.Text = $ComputerName

$Computer_Box = New-Object System.Windows.Forms.ListBox
$Computer_Box.Location = New-Object System.Drawing.Point(5,91)
$Computer_Box.AutoSize = $true
$Computer_Box.MinimumSize = New-Object System.Drawing.Size(224,200)
$Computer_Box.MaximumSize = New-Object System.Drawing.Size(0,200)
$Computer_Box.ScrollAlwaysVisible = $true
$Computer_Box.Items.Clear()
$search_form.Controls.Add($Computer_Box)

$Computer_button = New-Object System.Windows.Forms.Button
$Computer_button.Location = New-Object System.Drawing.Point(208,26)
$Computer_button.Size = New-Object System.Drawing.Size(21,21)
$Computer_button.Text = "V"
$search_form.AcceptButton = $Computer_button

$Computer_trigger = {
    $ComputerName = $Computer_bar.Text
    If ($Computer_Box.Items -notcontains $ComputerName) {
        $Computer_Box.Items.Add($ComputerName)
    }
}
$Computer_button.Add_click($Computer_trigger)
$search_form.Controls.Add($Computer_button)

#Application Search
$search_label = New-Object System.Windows.Forms.Label
$search_label.Text = "Applications Search:"
$search_label.Location = New-Object System.Drawing.Point(245,8)
$search_label.Size = New-Object System.Drawing.Size(150,20)
$search_form.Controls.Add($search_label)

$results_label = New-Object System.Windows.Forms.Label
$results_label.Text = "Application Results:"
$results_label.Location = New-Object System.Drawing.Point(245,70)
$results_label.Size = New-Object System.Drawing.Size(150,21)
$search_form.Controls.Add($results_label)

$search_bar = New-Object System.Windows.Forms.TextBox
$search_bar.Location = New-Object System.Drawing.Size(245,28)
$search_bar.Size = New-Object System.Drawing.Size(400,21)
$search_form.Controls.Add($search_bar)
$search_bar.Text = $assigned_to_box.Text

$selectionBox = New-Object System.Windows.Forms.ListBox
$selectionBox.Location = New-Object System.Drawing.Point(245,91)
$selectionBox.AutoSize = $true
$selectionBox.MinimumSize = New-Object System.Drawing.Size(421,200)
$selectionBox.MaximumSize = New-Object System.Drawing.Size(0,200)
$selectionBox.ScrollAlwaysVisible = $true
$selectionBox.Items.Clear()
$search_form.Controls.Add($selectionBox)

$search_button = New-Object System.Windows.Forms.Button
$search_button.Location = New-Object System.Drawing.Point(648,26)
$search_button.Size = New-Object System.Drawing.Size(21,21)
$search_button.BackgroundImage = $image
$search_button.BackgroundImageLayout = 'Zoom'

$search_trigger = {
    $search = $search_bar.Text
    $selectionBox.Items.Clear()
    $selectionBox.Items.Add("Processing . . .")
    If ($SCCMList -notmatch '\w') {
        If ($Computer_Box.Items -eq $null) {
            [System.Windows.MessageBox]::Show("SCCM App List isn't populated. Add a Computer Name to the Box to try to populate the list.")
        }else{
            for ($i = 0; $i -lt $Computer_Box.Items.Count; $i++) {
                $SCCMList = (Get-WmiObject -Class CCM_Application -Namespace "root\ccm\clientSDK" -ComputerName $Computer_Box.Items[$i]).Name
                If ($SCCMList -match '\w') {
                    $i = $Computer_Box.Items.Count
                }
            }
            If ($SCCMList -notmatch '\w') {
                [System.Windows.MessageBox]::Show("SCCM App List isn't populated. You can try adding another Computer Name to see if it will populate.")
            }else{
                $apps_list = $SCCMList + $OtherApps
            }
        }
    }
    $search_result = $apps_list | Where-Object {$_ -match $search}
    $selectionBox.Items.Clear()
    If ($search_result -notmatch '\w') {
        $selectionBox.Items.Add("No Results Found")
    }else{
        foreach ($item in $search_result) {
            [void] $selectionBox.Items.Add($item)
        }
    }
}

$search_button.add_click($search_trigger)
$search_form.Controls.Add($search_button)

#Install Status
$status_label = New-Object System.Windows.Forms.Label
$status_label.Location = New-Object System.Drawing.Point(5,$($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 30))
$status_label.Size = New-Object System.Drawing.Size(300,20)
$status_label.Text = "Installation Status:"
$search_form.Controls.Add($status_label)

$status_box = New-Object System.Windows.Forms.ListBox
$status_box.Location = New-Object System.Drawing.Point(5,$($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 50))
$status_box.AutoSize = $true
$status_box.MinimumSize = New-Object System.Drawing.Size(658,260)
$status_box.MaximumSize = New-Object System.Drawing.Size(0,260)
$status_box.ScrollAlwaysVisible = $true
$status_box.Items.Clear()

$Show_Status = {
    If ($status_box.SelectedItem -match '\w') {
        If ($SCCMList -contains $selectionBox.SelectedItem) {
            $search_form.Enabled = $false
            $computercheck = $($status_box.SelectedItem).Split(" ")
            If ($computercheck[0] -ne "Installed" -and $computercheck[0] -ne "Uninstalled" -and $computercheck[0] -ne "Error" -and $computercheck[0] -ne "Already") {
                $computercheck = $computercheck[$($computercheck.count - 1)]
                $appcheck_array = $status_box.SelectedItem.Split(" ")[4..($($status_box.SelectedItem.Split(" ")).Length-3)]
                Clear-Variable -Name appcheck -ErrorAction SilentlyContinue
                for ($i = 0; $i -lt $appcheck_array.count; $i++) {
                    If ($i -eq 0) {
                        $appcheck = $($appcheck_array[$i])
                    }else{
                        $appcheck = "$appcheck $($appcheck_array[$i])"
                    }
                }
                $check_SCCM_status = Get-WmiObject -Class CCM_Application -Namespace "root\ccm\clientSDK" -ComputerName $computercheck | Where-Object {$_.Name -like "$appcheck"}
                $search_form.Enabled = $true
                $methodcheck = $status_box.SelectedItem.Split(" ")[2]
                $message = "$computercheck
$appcheck is: $($check_SCCM_status.InstallState)"
                Add-Type -AssemblyName PresentationFramework
                [System.Windows.MessageBox]::Show("$message")
                If ($check_SCCM_status.InstallState -eq "Installed") {
                    If ($methodcheck -eq "Install") {
                        $status_item = "Installed $appcheck on $computercheck"
                        $status_box.Items.RemoveAt($status_box.SelectedIndex)
                        $status_box.Items.Add($status_item)
                    }
                }else{
                    If ($methodcheck -eq "Uninstall") {
                        $status_item = "Uninstalled $appcheck on $computercheck"
                        $status_box.Items.RemoveAt($status_box.SelectedIndex)
                        $status_box.Items.Add($status_item)
                    }
                }
            }else{
                $search_form.Enabled = $true
            }
        }else{

        }
    }
}

$status_box.add_Click($Show_Status)
$search_form.Controls.Add($status_box)

$install_button = New-Object System.Windows.Forms.Button
$install_button.Location = New-Object System.Drawing.Point(513, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 5))
$install_button.Size = New-Object System.Drawing.Size(75,23)
$install_button.Text = 'Install'

$install_trigger = {
    If ($Computer_box.SelectedItem -match '\w' -and $selectionBox.SelectedItem -match '\w') {
        $selected_app = $selectionBox.SelectedItem
        $selected_computer = $Computer_box.SelectedItem
        $status_text = "Installing $selected_app on $selected_computer"
        $compare_text1 = "Successfully started Install of $selected_app on $selected_computer"
        $compare_text2 = "Installed $selected_app on $selected_computer"
        If ($status_box.Items -contains $status_text -or $status_box.Items -contains $compare_text1 -or $status_box.Items -contains $compare_text2) {
            $status_text = "Already installing $selected_app on $selected_computer"
            $status_box.Items.Add($status_text)
        }else{
            $status_box.Items.Add($status_text) 
            If ($SCCMList -contains $selected_app) {
                $search_form.Enabled = $false
                $Application = (Get-WmiObject -Class CCM_Application -Namespace "root\ccm\clientSDK" -ComputerName $selected_computer | Where-Object {$_.Name -like $selected_app})
                $SCCM_Request = Invoke-WmiMethod -Namespace "root\ccm\clientSDK" -Class CCM_Application -Name "Install" -ComputerName $selected_computer -ArgumentList @(0, $Application.Id, $Application.IsMachineTarget, $False, 'High', $Application.Revision) | select -ExpandProperty ReturnValue
                $search_form.Enabled = $true
                If ($selected_app -match "\(") {
                    $selected_app_open_top = $selected_app.Split("(")[0]
                    $selected_app_open_bottom = $selected_app.Split("(")[1..$($selected_app.Split("(").Count - 1)]
                    $selected_app_open_trim = "$($selected_app_open_top)\($($selected_app_open_bottom)"
                    $selected_app_closed_top = $selected_app_open_trim.Split(")")[0]
                    $selected_app_closed_bottom = $selected_app_open_trim.Split(")")[1..$($selected_app_open_trim.Split(")").Count - 1)]
                    $selected_app_trim = "$($selected_app_closed_top)\)$($selected_app_closed_bottom)"
                }else{
                    $selected_app_trim = $selected_app
                }
                If ($SCCM_Request -ne 0) {
                    $remove_old_status = $($status_box.Items | Where-Object {$_ -match $selected_app_trim})
                    If ($remove_old_status.Count -gt 1) {
                        foreach ($old_status in $remove_old_status) {
                            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$old_status))
                        }
                    }else{
                        $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$remove_old_status))
                    }
                    $status_box.Items.Add("Error starting Install of $($Application.Name) on $selected_computer")
                }else{
                    $remove_old_status = $($status_box.Items | Where-Object {$_ -match $selected_app_trim})
                    If ($remove_old_status.Count -gt 1) {
                        foreach ($old_status in $remove_old_status) {
                            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$old_status))
                        }
                    }else{
                        $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$remove_old_status))
                    }
                    $status_box.Items.Add("Successfully started Install of $($Application.Name) on $selected_computer")
                }
            }else{
                If ($selected_app -eq "Microsoft Office365 Web Apps") {
                    $Error.Clear()
                    Robocopy "\\wnis01\ftp$\Shortcuts\Office365" "\\$selected_computer\c$\ProgramData\Microsoft\Windows\Start Menu\Programs"
                    Robocopy "\\wnis01\ftp$\Shortcuts\Office365" "\\$selected_computer\c$\Users\Public\Desktop"
                    If ($Error) {
                        $remove_old_status = $($status_box.Items | Where-Object {$_ -match $selected_app})
                        If ($remove_old_status.Count -gt 1) {
                            foreach ($old_status in $remove_old_status) {
                                $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$old_status))
                            }
                        }else{
                            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$remove_old_status))
                        }
                        $status_box.Items.Add("Error starting Install of $selected_app on $selected_computer")
                    }else{
                        $remove_old_status = $($status_box.Items | Where-Object {$_ -match $selected_app})
                        If ($remove_old_status.Count -gt 1) {
                            foreach ($old_status in $remove_old_status) {
                                $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$old_status))
                            }
                        }else{
                            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$remove_old_status))
                        }
                        $status_box.Items.Add("Installed $selected_app on $selected_computer")
                    }
                }else{
                    If ($selected_app -eq "Microsoft Office365") {
                        $selected_app_install = "Office365"
                    }else{
                        If ($selected_app -eq "Microsoft Office 2010") {
                            $selected_app_install = "Office2010"
                        }else{
                            $selected_app_install = $selected_app
                        }
                    }
                    Start-Process powershell.exe -ArgumentList "Install-MISCapp -ComputerName $selected_computer -AppName $selected_app_install"
                }
            }
        }
    }
}

$install_button.Add_click($install_trigger)
$search_form.Controls.Add($install_button)

$uninstall_button = New-Object System.Windows.Forms.Button
$uninstall_button.Location = New-Object System.Drawing.Point(591, $($($selectionBox.Location.Y) + $($selectionBox.Size.Height) + 5))
$uninstall_button.Size = New-Object System.Drawing.Size(75,23)
$uninstall_button.Text = 'Uninstall'

$uninstall_trigger = {
    If ($Computer_box.SelectedItem -match '\w' -and $selectionBox.SelectedItem -match '\w') {
        $selected_app = $selectionBox.SelectedItem
        $selected_computer = $Computer_box.SelectedItem
        $compare_text1 = "Successfully started Uninstall of $selected_app on $selected_computer"
        $compare_text2 = "Uninstalled $selected_app on $selected_computer"
        If ($SCCMList -contains $selectionBox.SelectedItem) {
            $status_text = "Uninstalling $selected_app on $selected_computer"
            If ($status_box.Items -contains $status_text -or $status_box.Items -contains $compare_text1 -or $status_box.Items -contains $compare_text2) {
                $status_text = "Already uninstalling $selected_app on $selected_computer"
                $status_box.Items.Add($status_text)
            }else{
                $status_box.Items.Add($status_text)
                $search_form.Enabled = $false
                $Application = (Get-WmiObject -Class CCM_Application -Namespace "root\ccm\clientSDK" -ComputerName $selected_computer | Where-Object {$_.Name -like $selected_app})
                $SCCM_Request = Invoke-WmiMethod -Namespace "root\ccm\clientSDK" -Class CCM_Application -Name "Uninstall" -ComputerName $selected_computer -ArgumentList @(0, $Application.Id, $Application.IsMachineTarget, $False, 'High', $Application.Revision) | select -ExpandProperty ReturnValue
                $search_form.Enabled = $true
                If ($selected_app -match "\(") {
                    $selected_app_open_top = $selected_app.Split("(")[0]
                    $selected_app_open_bottom = $selected_app.Split("(")[1..$($selected_app.Split("(").Count - 1)]
                    $selected_app_open_trim = "$($selected_app_open_top)\($($selected_app_open_bottom)"
                    $selected_app_closed_top = $selected_app_open_trim.Split(")")[0]
                    $selected_app_closed_bottom = $selected_app_open_trim.Split(")")[1..$($selected_app_open_trim.Split(")").Count - 1)]
                    $selected_app_trim = "$($selected_app_closed_top)\)$($selected_app_closed_bottom)"
                }else{
                    $selected_app_trim = $selected_app
                }
                If ($SCCM_Request -ne 0) {
                    $remove_old_status = $($status_box.Items | Where-Object {$_ -match $selected_app_trim})
                    If ($remove_old_status.Count -gt 1) {
                        foreach ($old_status in $remove_old_status) {
                            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$old_status))
                        }
                    }else{
                        $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$remove_old_status))
                    }
                    $status_box.Items.Add("Error starting Uninstall of $($Application.Name) on $selected_computer")
                }else{
                    $remove_old_status = $($status_box.Items | Where-Object {$_ -match $selected_app_trim})
                    If ($remove_old_status.Count -gt 1) {
                        foreach ($old_status in $remove_old_status) {
                            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$old_status))
                        }
                    }else{
                        $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$remove_old_status))
                    }
                    $status_box.Items.Add("Successfully started Uninstall of $($Application.Name) on $selected_computer")
                }
            }
        }else{
            $status_text = "Can't Uninstall $selected_app Remotely"
            $status_box.Items.Add($status_text)
        }
    }
}
$uninstall_button.Add_click($uninstall_trigger)
$search_form.Controls.Add($uninstall_button)

$end_button = New-Object System.Windows.Forms.Button
$end_button.Location = New-Object System.Drawing.Point(591,$($($status_box.Location.Y) + $($status_box.Size.Height) + 3))
$end_button.Size = New-Object System.Drawing.Size(75,23)
$end_button.Text = 'Exit'
$end_button.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$search_form.Controls.Add($end_button)

$computer_accept = {
    $search_form.AcceptButton = $Computer_button
    $Computer_box.SelectedItem = $null
    $selectionBox.SelectedItem = $null
    $status_box.SelectedItem = $null
    If ($status_box.Items -match "Remotely" -or $statusBox.Items -match "Already") {
        $wrong_status = $($status_box.Items | Where-Object {$_ -match "Remotely" -or $_ -match "Already"})
        If ($wrong_status.Count -gt 1) {
            foreach ($incorrect_status in $wrong_status) {
                $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$incorrect_status))
            }
        }else{
            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$wrong_status))
        }
    }
}
$search_accept = {
    $search_form.AcceptButton = $search_button
    $selectionBox.SelectedItem = $null
    $status_box.SelectedItem = $null
    If ($status_box.Items -match "Remotely" -or $status_box.Items -match "Already") {
        $wrong_status = $($status_box.Items | Where-Object {$_ -match "Remotely" -or $_ -match "Already"})
        If ($wrong_status.Count -gt 1) {
            foreach ($incorrect_status in $wrong_status) {
                $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$incorrect_status))
            }
        }else{
            $status_box.Items.RemoveAt([array]::indexof($($status_box.Items),$wrong_status))
        }
    }
}
$form_accept = {
    $search_form.AcceptButton = $install_button
}

$Computer_bar.add_MouseDown($computer_accept)
$search_bar.add_MouseDown($search_accept)
$selectionBox.add_MouseDown($form_accept)

$search_form.ShowDialog()