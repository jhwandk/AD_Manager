# things to work on: find and replace substring of an attribute string value
# preserving find and replace value during the run-time of the program

<#
    The function on the very bottom, Browse AD function/GUI is NOT my own work (please see function for credit)
    all other functions are my work
    Author: Jinhwan Kim
#>

Set-ExecutionPolicy Bypass -Scope Process -Force

function ShowCredit
{
    Write-Host '-------------------------------------------------------------------------------'
    Write-Host "AD Attributes Manager"
    Write-Host "Author: Jinhwan Kim"
    Write-Host "Date: 7/27/23"
    Write-Host '-------------------------------------------------------------------------------'
}

# Main menu
function ShowUserMenu
{
    while ($true)
    {
        Write-Host '-------------------------------------------------------------------------------'
        Write-Host "User menu`n"
        Write-Host '1. User(s)'
        Write-Host '2. Users in specific OU'
        Write-Host 'X. Quit'
        Write-Host '-------------------------------------------------------------------------------'
        $Answer = Read-Host 'Select a scope'
        switch ($Answer)
        {
            1 { ShowTaskMenu $(GetUsername) }
            2 { ShowTaskMenu $(GetUsername -OU) } 
            X { return }
            default
            {
                Write-Host 'Invalid input - try again' -ForegroundColor Red
            }
        }
    }
}

function GetUsername
{
    param ([switch] $OU)

    # If no OU switch (single user selection)
    if (!$OU.IsPresent)
    {
        # Get username input
        Write-Host '-------------------------------------------------------------------------------'
        $Answer = Read-Host 'Enter username(s) separated by commas (e.g. John.Doe, Jane.Done)'
        [String[]]$AnswerArray = @($Answer.Split(',').Trim())

        if (!$AnswerArray)
        {
            Write-Host 'No input' -ForegroundColor Red
        }
        else
        {
            $invalidElements = @()
        
            foreach ($Element in $AnswerArray)
            {
                # If empty or whitespace-only element
                if ([string]::IsNullOrWhiteSpace($Element))
                {
                    $invalidElements += "Empty or whitespace-only element"
                    continue
                }
                
                # If user not found in AD
                if (!(Get-ADUser -Filter "samAccountName -eq '$($Element)'"))
                {
                    $invalidElements += "$Element not found"
                    continue
                }
            }
        
            if ($invalidElements)
            {
                Write-Host "$($invalidElements -join '; ')" -ForegroundColor Red
            }
            else
            {
                # Input is valid - return string array of usernames
                return $AnswerArray
            }
        }
    }
    # Else - OU switch is on
    else
    {
        # Get OU input
        Write-Host 'Select OU or domain from the window - child OUs will be included'
        $Answer = BrowseAD
        # Validate input - if no selection
        if (!$Answer)
        {
            Write-Host 'No OU selected' -ForegroundColor Red
        }
        # If root is selected
        elseif ($Answer -eq 'root')
        {
            Write-Host 'You must select an OU or domain' -ForegroundColor Red
        }
        # If empty OU
        elseif (!(Get-ADUser -SearchBase "$Answer" -Filter * | Select-Object -expand samAccountName))
        {
            Write-Host "Empty OU" -ForegroundColor Red
        }
        # Else - input is valid
        else
        {
            # Return a string array with usernames strings from OU
            return @(Get-ADUser -SearchBase "$Answer" -Filter * | Select-Object -expand samAccountName)
        }
    }
}
function ShowTaskMenu
{   
    param ([string[]] $Usernames)
    # If no usernames is passed, return to previous menu/function
    if (!$Usernames)
    {
        return
    }

    while ($true)
    {
        Write-Host '-------------------------------------------------------------------------------'
        if (!$Usernames[1])
        {
            Write-Host "Managing user: $Usernames`n"
        }
        else
        {
            $DisplayLimit = 4
            $DisplayedUsernames = $Usernames[0..$DisplayLimit] -join ', '
            # If more than 5 users are selected, return true = 1 = 2nd index to append ellipses
            $DisplayedUsernames += $('', ', ...')[($Usernames.Count -gt $DisplayLimit+1)]
            Write-Host "Managing users: $DisplayedUsernames`n"
        }
        Write-Host '1. Modify office name (PhysicalDeliveryAddressName)'
        Write-Host '2. Modify street address'
        Write-Host '3. Modify city (L)'
        Write-Host '4. Modify state (ST)'
        Write-Host '5. Modify postal code'
        Write-Host '6. Modify full address (all of the above)'
        Write-Host '7. View full address'
        Write-Host 'X. Go back'
        Write-Host '-------------------------------------------------------------------------------'
        $Answer = Read-Host 'Select a task'
        switch ($Answer)
        {
            1 { ModifyAttribute $Usernames 'PhysicalDeliveryOfficeName' }
            2 { ModifyAttribute $Usernames 'StreetAddress' }
            3 { ModifyAttribute $Usernames 'L' }
            4 { ModifyAttribute $Usernames 'ST' }
            5 { ModifyAttribute $Usernames 'PostalCode' }
            6 { ModifyAttribute $Usernames @('PhysicalDeliveryOfficeName', 'StreetAddress', 'L', 'PostalCode', 'ST') }
            7 { ShowAttribute $Usernames @('PhysicalDeliveryOfficeName', 'StreetAddress' , 'L', 'PostalCode', 'ST') }
            X { return }
            default
            {
                Write-Host 'Invalid input - try again' -ForegroundColor Red
            }
        }
    }
}

function GetAttributeName
{
    Write-Host '-------------------------------------------------------------------------------'
    $Input = Read-Host 'Enter AD attributes(s) separated by commas'
    [String[]]$Entries = @($Input.Split(',').Trim())
    if (!$Entries)
    {
        Write-Host 'No input' -ForegroundColor Red
    }
    else
    {
        $InvalidEntries = @()
    
        foreach ($Entry in $Entries)
        {
            # If empty or whitespace-only element
            if ([string]::IsNullOrWhiteSpace($Entry))
            {
                $InvalidEntries += "Empty or whitespace-only element"
                continue
            }
            try
            {
                $null = Get-ADUser -Filter * -Properties $Entry -ErrorAction Stop
            }
            catch
            {
                $InvalidEntries += "$Entry is not an LDAP attribute"
                continue
            }
        }
        if ($InvalidEntries)
        {
            Write-Host "$($InvalidEntries -join '; ')" -ForegroundColor Red
        }
        else
        {
            # Input is valid - return the entries array
            return $Entries
        }
    }
}

function ModifyAttribute
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string[]]$Usernames,
        [Parameter(Mandatory=$false, Position=1)]
        [string[]]$AttributeNames
    )

    function ConfirmModification
    {
        while ($true)
        {
            Write-Host '-------------------------------------------------------------------------------'
            foreach ($AttributeName in $AttributeNames)
            {
                $FindAttributeValue = (Get-Variable "Find$AttributeName").Value
                $ReplaceAttributeValue = (Get-Variable "Replace$AttributeName").Value
                Write-Host "Attribute: $AttributeName"
                if ($FindAttributeValue -ne '*')
                {
                    Write-Host "Find: $FindAttributeValue"
                }
                Write-Host "Replace with: $ReplaceAttributeValue`n"
            }
            Write-Host 'P. Proceed'
            $Count = 0
            foreach ($AttributeName in $AttributeNames)
            {
                Write-Host "$Count. Adjust $AttributeName"
                $Count++
            }
            Write-Host 'X. Cancel'
            Write-Host '-------------------------------------------------------------------------------'
            $Answer = Read-Host 'Select an option'
            $Index = $Answer
            if ($Answer -eq 'P')
            {
                # Create log file if it doesn't exist
                $LogFilePath = "$PSScriptRoot\log.txt"
                if (!(Test-Path $LogFilePath))
                {
                    New-Item -ItemType File -Path $LogFilePath | Out-Null
                }
                # Add date and time stamp before logging results
                $DateTimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                "=== $DateTimeStamp ===" | Out-File -FilePath $LogFilePath -Append

                # Loop through usernames
                foreach ($Username in $Usernames)
                {
                    # Get user at the index
                    $User = Get-ADUser -Identity $Username -Properties *
                    foreach ($AttributeName in $AttributeNames)
                    {
                        # Get attribute values for currently indexed user
                        $CurrentValue = $User.$AttributeName
                        $FindValue = (Get-Variable "Find$AttributeName").value
                        $ReplaceValue = (Get-Variable "Replace$AttributeName").value
                        # If current value is same as new value
                        if ($FindValue -eq '*')
                        {
                            Set-ADUser -Identity $User.samAccountName -Replace @{$AttributeName=$ReplaceValue}
                            "$Username - $AttributeName - The value was overwritten" | Tee-Object -FilePath $LogFilePath -Append
                        }
                        elseif ($CurrentValue -eq $ReplaceValue)
                        {
                            # Print message
                            Write-Output "$Username - $AttributeName - The value is already new" | Tee-Object -FilePath $LogFilePath -Append
                        }
                        # If current value is same as old value
                        elseif ($CurrentValue -eq $FindValue)
                        {
                            Set-ADUser -Identity $user.samAccountName -Replace @{$AttributeName=$ReplaceValue}
                            Write-Output "$Username - $AttributeName - Match was found and replaced" | Tee-Object -FilePath $LogFilePath -Append
                        }
                        # Otherwise, if value is neither old or new
                        else
                        {
                            # Print message
                            Write-Output "$Username - $AttributeName - No match was found" | Tee-Object -FilePath $LogFilePath -Append
                        }
                    }
                }
                Write-Host 'Result has been appended to the log file'
                Write-Host 'Wait 1 minute for the change to take effect across the board'
                Read-Host 'Press enter to continue'
            }
            elseif ($AttributeNames[$Index])
            {
                while ($true)
                {
                    Write-Host '-------------------------------------------------------------------------------'
                    Write-Host "Adjusting $($AttributeNames[$Index])"
                    Write-Host "W. Overwrite"
                    Write-Host "F. Find and replace"
                    Write-Host "X. Cancel"
                    $Answer = Read-Host
                    if ($Answer -eq 'W')
                    {
                        Set-Variable -Name "Find$($AttributeNames[$Index])" -Value '*'
                        $Value = Read-Host "Enter new value to replace it with"
                        Set-Variable -Name "Replace$($AttributeNames[$Index])" -Value $Value
                        break
                    }
                    elseif ($Answer -eq 'F')
                    {
                        $Value = Read-Host "Enter old value to find"
                        Set-Variable -Name "Find$($AttributeNames[$Index])" -Value $Value
                        $Value = Read-Host "Enter new value to replace it with"
                        Set-Variable -Name "Replace$($AttributeNames[$Index])" -Value $Value
                        break
                    }
                    elseif ($Answer -eq 'X')
                    {
                        break
                    }
                    else
                    {
                        Write-Host 'Invalid input - try again' -ForegroundColor Red
                    }
                }
            }
            elseif ($Index -eq 'X')
            {
                break
            }
            else
            {
                Write-Host 'Invalid input - try again' -ForegroundColor Red
            }
        }
    }

    # Check Attribute Names
    if (!$AttributeNames)
    {
        return
    }
    else
    {
        foreach ($AttributeName in $AttributeNames)
        {
            try
            {
                $null = Get-ADUser -Filter * -Properties $AttributeName -ErrorAction Stop
            }
            catch
            {
                Write-Host "$AttributeName is not an LDAP attribute" -ForegroundColor Red
                return
            }
        }
    }

    foreach ($AttributeName in $AttributeNames)
    {
        # If the variable exists in memory
        if (Get-Variable "Find$AttributeName" -Scope Script -ErrorAction SilentlyContinue)
        {
            Write-Host "Variables already exist for attribute `'$AttributeName`'"
        }
        # If the attribute variable doesn't exist in memory
        else
        {
            try
            {
                New-Variable -Name "Find$AttributeName" -Value ''
                New-Variable -Name "Replace$AttributeName" -Value ''

                # Create file to store default placeholder values if the file doesn't exist
                $PlaceHolderFilePath = "$PSScriptRoot\placeholders.txt"
                if (!(Test-Path $PlaceHolderFilePath))
                {
                    New-Item -ItemType File -Path $PlaceHolderFilePath | Out-Null
                    $PlaceHolderContent = "`nPhysicalDeliveryOfficeName"
                    $PlaceHolderContent += "`tHeadquarters - ABC, CA"
                    $PlaceHolderContent += "`tHeadquarters - XYZ, CA"
                    $PlaceHolderContent += "`nStreetAddress"
                    $PlaceHolderContent += "`tABC Drive, 1st Floor"
                    $PlaceHolderContent += "`tXYZ Pkwy #300"
                    $PlaceHolderContent += "`nL"
                    $PlaceHolderContent += "`tNice City"
                    $PlaceHolderContent += "`tGood City"
                    $PlaceHolderContent += "`nST"
                    $PlaceHolderContent += "`tCA"
                    $PlaceHolderContent += "`tCA"
                    $PlaceHolderContent += "`nPostalCode"
                    $PlaceHolderContent += "`t92000"
                    $PlaceHolderContent += "`t92111"
                    Set-Content -Path $PlaceHolderFilePath -Value $PlaceHolderContent
                }
                $File = Get-Content $PlaceHolderFilePath
                foreach ($Line in $File)
                {
                    $SplitLine = $Line.split("`t").Trim()
                    $LineAttributeName = $SplitLine[0]
                    $LineFindValue = $SplitLine[1]
                    $LineReplaceValue = $SplitLine[2]
                    if ($LineAttributeName -eq $AttributeName)
                    {
                        # Supress any error by using $null, instead of console
                        Set-Variable -Name "Find$AttributeName" -Value $LineFindValue -ErrorAction Stop
                        Set-Variable -Name "Replace$AttributeName" -Value $LineReplaceValue -ErrorAction Stop
                        break
                    }
                }
            }
            catch
            {
                Write-Host "Failed to create variables for attribute '$AttributeName': $_" -Verbose
            }
        }    
    }
    
    ConfirmModification

}

function ShowAttribute
{
    param
    (
        [Parameter(Mandatory=$true, Position=0)]
        [string[]] $Usernames,
        [Parameter(Mandatory=$true, Position=1)]
        [string[]] $AttributeNames
    )

    Write-Host '-------------------------------------------------------------------------------'
    $FirstUser = $true
    foreach ($Username in $Usernames)
    {
        if (!$FirstUser)
        {
            Write-Host "`n" -NoNewLine
        }
        Write-Host 'Username: ' -NoNewLine
        (Get-ADUser -Identity $Username).samAccountName
        foreach ($AttributeName in $AttributeNames)
        {
            Write-Host "$AttributeName`: " -NoNewLine
            (Get-ADUser -Identity $Username -Properties *).$AttributeName
        }
        $FirstUser = $false
    }
    Write-Host '-------------------------------------------------------------------------------'
    Read-Host 'Press enter to continue'
}

function BrowseAD()
{
    # original inspiration: https://itmicah.wordpress.com/2013/10/29/active-directory-ou-picker-in-powershell/
    # author: Rene Horn the.rhorn@gmail.com
<#
    Copyright (c) 2015, Rene Horn
    All rights reserved.
    Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:
    1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.
    2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.
    3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.
    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#>
    $dc_hash = @{}
    $selected_ou = $null

    Import-Module ActiveDirectory
    $forest = Get-ADForest
    [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") | Out-Null
    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null

    function Get-NodeInfo($sender, $dn_textbox)
    {
        $selected_node = $sender.Node
        $dn_textbox.Text = $selected_node.Name
    }

    function Add-ChildNodes($sender)
    {
        $expanded_node = $sender.Node

        if ($expanded_node.Name -eq "root") {
            return
        }

        $expanded_node.Nodes.Clear() | Out-Null

        $dc_hostname = $dc_hash[$($expanded_node.Name -replace '((OU|CN)=[^,]+,)*((DC=\w+,?)+)','$3')]
        $child_OUs = Get-ADObject -Server $dc_hostname -Filter 'ObjectClass -eq "organizationalUnit" -or ObjectClass -eq "container"' -SearchScope OneLevel -SearchBase $expanded_node.Name
        if($child_OUs -eq $null) {
            $sender.Cancel = $true
        } else {
            foreach($ou in $child_OUs) {
                $ou_node = New-Object Windows.Forms.TreeNode
                $ou_node.Text = $ou.Name
                $ou_node.Name = $ou.DistinguishedName
                $ou_node.Nodes.Add('') | Out-Null
                $expanded_node.Nodes.Add($ou_node) | Out-Null
            }
        }
    }

    function Add-ForestNodes($forest, [ref]$dc_hash)
    {
        $ad_root_node = New-Object Windows.Forms.TreeNode
        $ad_root_node.Text = $forest.RootDomain
        $ad_root_node.Name = "root"
        $ad_root_node.Expand()

        $i = 1
        foreach ($ad_domain in $forest.Domains) {
            Write-Progress -Activity "Querying AD forest for domains and hostnames..." -Status $ad_domain -PercentComplete ($i++ / $forest.Domains.Count * 100)
            $dc = Get-ADDomainController -Server $ad_domain
            $dn = $dc.DefaultPartition
            $dc_hash.Value.Add($dn, $dc.Hostname)
            $dc_node = New-Object Windows.Forms.TreeNode
            $dc_node.Name = $dn
            $dc_node.Text = $dc.Domain
            $dc_node.Nodes.Add("") | Out-Null
            $ad_root_node.Nodes.Add($dc_node) | Out-Null
        }

        return $ad_root_node
    }
    
    $main_dlg_box = New-Object System.Windows.Forms.Form
    $main_dlg_box.ClientSize = New-Object System.Drawing.Size(400,600)
    $main_dlg_box.MaximizeBox = $false
    $main_dlg_box.MinimizeBox = $false
    $main_dlg_box.FormBorderStyle = 'FixedSingle'

    # widget size and location variables
    $ctrl_width_col = $main_dlg_box.ClientSize.Width/20
    $ctrl_height_row = $main_dlg_box.ClientSize.Height/15
    $max_ctrl_width = $main_dlg_box.ClientSize.Width - $ctrl_width_col*2
    $max_ctrl_height = $main_dlg_box.ClientSize.Height - $ctrl_height_row
    $right_edge_x = $max_ctrl_width
    $left_edge_x = $ctrl_width_col
    $bottom_edge_y = $max_ctrl_height
    $top_edge_y = $ctrl_height_row

    # setup text box showing the distinguished name of the currently selected node
    $dn_text_box = New-Object System.Windows.Forms.TextBox
    # can not set the height for a single line text box, that's controlled by the font being used
    $dn_text_box.Width = (14 * $ctrl_width_col)
    $dn_text_box.Location = New-Object System.Drawing.Point($left_edge_x, ($bottom_edge_y - $dn_text_box.Height))
    $main_dlg_box.Controls.Add($dn_text_box)
    # /text box for dN

    # setup Ok button
    $ok_button = New-Object System.Windows.Forms.Button
    $ok_button.Size = New-Object System.Drawing.Size(($ctrl_width_col * 2), $dn_text_box.Height)
    $ok_button.Location = New-Object System.Drawing.Point(($right_edge_x - $ok_button.Width), ($bottom_edge_y - $ok_button.Height))
    $ok_button.Text = "Ok"
    $ok_button.DialogResult = 'OK'
    $main_dlg_box.Controls.Add($ok_button)
    # /Ok button

    # setup tree selector showing the domains
    $ad_tree_view = New-Object System.Windows.Forms.TreeView
    $ad_tree_view.Size = New-Object System.Drawing.Size($max_ctrl_width, ($max_ctrl_height - $dn_text_box.Height - $ctrl_height_row*1.5))
    $ad_tree_view.Location = New-Object System.Drawing.Point($left_edge_x, $top_edge_y)
    $ad_tree_view.Nodes.Add($(Add-ForestNodes $forest ([ref]$dc_hash))) | Out-Null
    $ad_tree_view.Add_BeforeExpand({Add-ChildNodes $_})
    $ad_tree_view.Add_AfterSelect({Get-NodeInfo $_ $dn_text_box})
    $main_dlg_box.Controls.Add($ad_tree_view)
    # /tree selector

    $main_dlg_box.ShowDialog() | Out-Null

    return  $dn_text_box.Text
}


ShowCredit
ShowUserMenu
#$Usernames = @('Jinhwan.Kim', 'Mary.Jane', 'John.Doe')
# Delete the cache file for attribute values remembered from run-time
# $Usernames = @('Jinhwan.Kim', 'Mary.Jane')
# ModifyAttribute -Usernames $Usernames -AttributeNames @('StreetAddress', 'PostalCode', 'State')
# ShowAttribute $Usernames @('PhysicalDeliveryOfficeName', 'StreetAddress' , 'L', 'PostalCode', 'State')
