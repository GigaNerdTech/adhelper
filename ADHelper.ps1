# Active Directory Create/Add/Transfer script
# Written by Joshua Woleben
# Written 8/27/19

# Import Active Directory PowerShell module
Import-Module ActiveDirectory

# Query security groups
$security_groups = Get-ADObject -Filter 'ObjectClass -eq "group"' -SearchBase "OU=Example Computers,DC=example,DC=com" 
$corporate_computer_ous = Get-ADObject -Filter 'ObjectClass -eq "OrganizationalUnit"' -SearchBase "OU=Example Computers,DC=example,DC=com" | Where-Object { $_.distinguishedName -notmatch "Filter1" -and $_.distinguishedName -notmatch "Filter2" -and $_.distinguishedName -notmatch "Filter3" -and $_.distinguishedName -notmatch "Filter4" }
$hospital_computer_ous = Get-ADObject -Filter 'ObjectClass -eq "OrganizationalUnit"' -SearchBase "OU=Example Computers,DC=example,DC=com"| Where-Object { $_.distinguishedName -notmatch "Filter1" -and $_.distinguishedName -notmatch "Filter2" -and $_.distinguishedName -notmatch "Filter3" -and $_.distinguishedName -notmatch "Filter4" }
$clinic_computer_ous = Get-ADObject -Filter 'ObjectClass -eq "OrganizationalUnit"' -SearchBase "OU=Example Computers,DC=example,DC=com"| Where-Object { $_.distinguishedName -notmatch "Filter1" -and $_.distinguishedName -notmatch "Filter2" -and $_.distinguishedName -notmatch "Filter3" -and $_.distinguishedName -notmatch "Filter4" }

# GUI Code
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Active Directory Helper" Height="850" Width="450" MinHeight="500" MinWidth="400" ResizeMode="CanResizeWithGrip">
    <StackPanel>
        <Label x:Name="AddCreateLabel" Content="Add/Create Object" FontSize="14" FontWeight="Bold"/>
        <Label x:Name="ObjectAddLabel" Content="New Object Name"/>
        <TextBox x:Name="ObjectAddTextBox" Margin="10,10,10,0"/>
        <Label x:Name="SecurityGroupLabel" Content="Select Security Groups"/>
        <ListBox x:Name="SecurityGroupSelect" MinHeight = "50" AllowDrop="True" SelectionMode="Extended" Margin="10,10,10,0"/>
        <Label x:Name="ComputerOULabel" Content="Select Computer OUs"/>
        <ListBox x:Name="ComputerOUSelect" MinHeight = "50" AllowDrop="True" SelectionMode="Extended" Margin="10,10,10,0"/>
        <Button x:Name="AddCreateButton" Content="Add/Create Object" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
        <Separator Height="45"/>
        <Label x:Name="TransferLabel" Content="Transfer Object" FontSize="14" FontWeight="Bold"/>
        <Label x:Name="SourceObjectLabel" Content="Source Object Name"/>
        <TextBox x:Name="SourceObjectTextBox" Margin="10,10,10,0"/>
        <Label x:Name="TargetObjectLabel" Content="Target Object Name"/>
        <TextBox x:Name="TargetObjectTextBox" Margin="10,10,10,0"/>
        <Button x:Name="TransferButton" Content="Transfer Object" Margin="10,10,10,0" VerticalAlignment="Top" Height="25"/>
    </StackPanel>
</Window>
'@
 
$global:Form = ""
# XAML Launcher
$reader=(New-Object System.Xml.XmlNodeReader $xaml) 
try{$global:Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; break}
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name ($_.Name) -Value $global:Form.FindName($_.Name)}


# Take control of form elements
$ObjectAddTextBox = $global:Form.FindName('ObjectAddTextBox')
$SecurityGroupSelect = $global:Form.FindName('SecurityGroupSelect')
$ComputerOUSelect = $global:Form.FindName('ComputerOUSelect')
$AddCreateButton = $global:Form.FindName('AddCreateButton')
$SourceObjectTextBox = $global:Form.FindName('SourceObjectTextBox')
$TargetObjectTextBox = $global:Form.FindName('TargetObjectTextBox')
$TransferButton = $global:Form.FindName('TransferButton')

# Populate security group list
$security_groups | ForEach-Object { $SecurityGroupSelect.Items.Add($_.Name) | out-null }

# Populate OUs
$corporate_computer_ous | ForEach-Object { $name = (Select-String -InputObject $_.distinguishedName -Pattern "OU=Computers,OU=(.*?),").Matches.Groups[1].Value; $ComputerOUSelect.Items.Add($name) | out-null }
$hospital_computer_ous | ForEach-Object { $name = (Select-String -InputObject $_.distinguishedName -Pattern "OU=(.*?),").Matches.Groups[1].Value; if ($name -match "Computers") { $name="Methodist Health System Root"} $ComputerOUSelect.Items.Add($name) | out-null }
$clinic_computer_ous | ForEach-Object { $name = (Select-String -InputObject $_.distinguishedName -Pattern "OU=Computers,OU=(.*?),").Matches.Groups[1].Value; $ComputerOUSelect.Items.Add($name) | out-null }

# Button control functions
$AddCreateButton.Add_Click({

    # Get hostname from textbox
    $computer_name = $ObjectAddTextBox.Text
    $selected_groups = @()
    # Get selected Security groups
    foreach ($group in $SecurityGroupSelect.SelectedItems) {
        $selected_groups += $group.ToString()
    }

    # Get Selected OU
    $selected_ou = $ComputerOUSelect.SelectedItem

    # Query AD for existence of computer object
    try {
        $ad_computer = Get-ADComputer -Identity $computer_name
    }
    catch {
        
        New-ADComputer -Name $computer_name -SAMAccountName $computer_name -DNSHostName "$computer_name.mhs.int" -DisplayName $computer_name -Confirm:$false
        sleep 10
    } 

    # Modify security groups and OU
    ForEach ($group in $selected_groups) {
        Write-Host "Adding to group $group..."
        Add-ADGroupMember -Identity (Get-ADGroup -Filter 'name -eq $group').distinguishedName -Members (Get-ADComputer -Identity $computer_name).distinguishedName
    }
    $combined_ous = @()
    $corporate_computer_ous | ForEach-Object { $combined_ous += $_ }
    $hospital_computer_ous | ForEach-Object { $combined_ous += $_ }
    $clinic_computer_ous | ForEach-Object { $combined_ous += $_ }

    Move-ADObject -Identity (Get-ADComputer -Identity $computer_name).distinguishedName -TargetPath ($combined_ous | Where-Object { $_.name -match $selected_ou }).distinguishedName
  
    [System.Windows.MessageBox]::Show("Completed!")

})

$TransferButton.Add_Click({

    # Get source and target computer name
    $source_computer = $SourceObjectTextBox.Text
    $target_computer = $TargetObjectTextBox.Text

    # Check for existence of source
    try {
        Get-ADComputer -Identity $source_computer
    }
    catch {
        [System.Windows.MessageBox]::Show("Error! Source doesn't exist!")
        break
    }

    # check for non-existence of target
   try {
        Get-ADComputer -Identity $target_computer
    }


    catch {
        # Create new object
        New-ADComputer -Name $target_computer -SAMAccountName $target_computer -DNSHostName "$target_computer.mhs.int" -DisplayName $target_computer -Confirm:$false

        # Get source security groups
        $source_groups = Get-ADObject -Identity $source_computer -Filter 'ObjectClass -eq "Group"'

        # Get source OU
        $source_ou = (Get-ADObject -Identity $source_computer -Filter 'ObjectClass -eq "OrganizationalUnit"').distinguishedName

        # Add target to security groups
        $source_groups | ForEach-Object { Add-ADGroupMember -Identity $_.distinguishedName -Members $target_computer -Confirm:$false }

        # Move target to correct OU
        Move-ADObject -Identity $target_computer -TargetPath $source_ou

    [System.Windows.MessageBox]::Show("Completed!")
    }
   [System.Windows.MessageBox]::Show("Error! Target already exists!")
})

# Show GUI
$global:Form.ShowDialog() | out-null