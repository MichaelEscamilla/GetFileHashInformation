<#
.SYNOPSIS
This script provides a graphical user interface (GUI) for viewing and copying properties of MSI files.

.DESCRIPTION
The script creates a WPF-based GUI that allows users to drag and drop MSI files to view their properties such as Product Name, Manufacturer, Product Version, Product Code, and Upgrade Code.
It also provides functionality to copy these properties to the clipboard and to clear the displayed information.
Additionally, the script includes options to install and uninstall a context menu item for MSI files to retrieve their properties.

.PARAMETER FilePath
Optional parameter to specify the path of the MSI file to automatically load the information for.

.NOTES
Author: Michael Escamilla
Date: 9-30-2024

Version History:
1.0.0.0     - Initial release
2.0.0.0     - 10-4-2024 - Added file hash information, and context menu items for the installation and uninstallation of a Right-Click Option in Windows Explorer. And some other UI improvements.
2024.10.4.1 - Updated the version numbering, and a sepearator in the context menu.
2024-10.13.0- Added an error message when the file is locked
#>

param (
  [Parameter(Mandatory = $false)]
  [string]$FilePath
)

#############################################
################# Variables #################
#############################################
# Script Name
$Global:ScriptName = "GetMSIInformation.ps1"
# Script Version
[System.Version]$Global:ScriptVersion = "2024.10.13.0"
# Right-Click Menu Name
$Global:RightClickMenuName = "Get MSI Information"
# Get the Security Principal
$Global:currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())

#############################################
################# Functions #################
#############################################
#region Functions
function Test-FileLock {
  param (
    [Parameter(Mandatory = $true)]
    [string]$Path
  )

  try {
    $FileStream = [System.IO.File]::Open("$($Path)", 'Open', 'Write')
    $FileStream.Close()
    $FileStream.Dispose()
    return $false
  }
  catch {
    return $true
  }
}

function Get-MsiProperties {
  param (
    [Parameter(Mandatory = $true)]
    [IO.FileInfo[]]$Path
  )
	
  # Check if the MSI file path exists
  if (-not (Test-Path $Path)) {
    throw "The file $Path does not exist."
  }
	
  # Create a new Windows Installer COM object
  $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
	
  # Open the MSI database in read-only mode
  $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path.FullName, 0))
	
  # Open a view on the Property table
  $MSIPropertyView = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, @("SELECT * FROM Property"))
	
  # Execute the view query
  $MSIPropertyView.GetType().InvokeMember("Execute", "InvokeMethod", $null, $MSIPropertyView, $null)
	
  # Fetch the first record from the result set
  $MSIRecord = $MSIPropertyView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $MSIPropertyView, $null)
	
  # Initialize an empty System Object to store properties
  [System.Object]$Properties = @{}
	
  # Loop through all records in the result set
  while ($null -ne $MSIRecord) {
    # Get the property name from the first column
    $property = $MSIRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $MSIRecord, @(1))
			
    # Get the property value from the second column
    $Value = $MSIRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $MSIRecord, @(2))
			
    # Add the property name and value to the hashtable
    $Properties[$Property] = $Value
			
    # Fetch the next record from the result set
    $MSIRecord = $MSIPropertyView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $MSIPropertyView, $null)
  }
	
  # Return the System Object of properties
  $Properties
}

function Enable-AllButtons {
  # Get all button variables
  $Buttons = Get-Variable -Name "btn_*" -ValueOnly -ErrorAction SilentlyContinue
  foreach ($Button in $Buttons) {
    # Enable Button
    $Button.IsEnabled = $true
  }
}

function Disable-AllButtons {
  # Get all button variables
  $Buttons = Get-Variable -Name "btn_*" -ValueOnly -ErrorAction SilentlyContinue
  foreach ($Button in $Buttons) {
    # Disable Button
    $Button.IsEnabled = $false
  }
}

function Clear-Textboxes {
  # Get all textbox variables
  $Textboxes = Get-Variable -Name "txt_*" -ValueOnly -ErrorAction SilentlyContinue
  foreach ($Textbox in $Textboxes) {
    # Disable Button
    $Textbox.Clear()
  }
}

# Stolen from: https://github.com/PatchMyPCTeam/CustomerTroubleshooting/blob/Release/PowerShell/Get-LocalContentHashes.ps1
Function Get-EncodedHash {
  [CmdletBinding()]
  Param(
    [Parameter(Position = 0)]
    [System.Object]$HashValue
  )

  $hashBytes = $hashValue.Hash -split '(?<=\G..)(?=.)' | ForEach-Object { [byte]::Parse($_, 'HexNumber') }
  Return [Convert]::ToBase64String($hashBytes)
}

function Get-FileHashInformation {
  param (
    [Parameter(Mandatory = $true)]
    [IO.FileInfo[]]$Path

  )

  Write-Host "Getting File Hash Information for: [$Path]"

  # Initialize the hash object
  $Hashes = @{}

  # Get File Hash - MD5
  $FileHashMD5 = Get-FileHash -Path $Path -Algorithm MD5
  $Hashes["MD5"] = $FileHashMD5

  # Get File Hash - SHA1
  $FileHashSHA1 = Get-FileHash -Path $Path -Algorithm SHA1
  $Hashes["SHA1"] = $FileHashSHA1

  # Get File Hash - SHA256
  $FileHashSHA256 = Get-FileHash -Path $Path -Algorithm SHA256
  $Hashes["SHA256"] = $FileHashSHA256

  # Get File Hash - SHA1 - Encoded
  $FileHashEncoded = Get-EncodedHash -HashValue $FileHashSHA1
  $Hashes["Digest"] = $FileHashEncoded

  # Return the hash object
  $Hashes
}

function Set-TextboxInformation {
  param (
    [Parameter(Mandatory = $true)]
    [System.Object]$MSIPropertiesInfo,
    [Parameter(Mandatory = $true)]
    [hashtable]$FileHashInfo
  )

  # Set the MSI file properties textboxes
  $txt_ProductName.Text = $MSIPropertiesInfo.ProductName
  $txt_Manufacture.Text = $MSIPropertiesInfo.Manufacturer
  $txt_ProductVersion.Text = $MSIPropertiesInfo.ProductVersion
  $txt_ProductCode.Text = $MSIPropertiesInfo.ProductCode
  $txt_UpgradeCode.Text = $MSIPropertiesInfo.UpgradeCode

  # Set the File Hash Information textboxes
  $txt_MD5.Text = $FileHashInfo.MD5.Hash
  $txt_SHA1.Text = $FileHashInfo.SHA1.Hash
  $txt_SHA256.Text = $FileHashInfo.SHA256.Hash
  $txt_Digest.Text = $FileHashInfo.Digest
}
#endregion Functions

#############################################
################# Main Script ################
#############################################

# Load Assemblies
Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Build the GUI
[xml]$XAMLformMSIProperties = @"
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Name="form1"
  Width="900"
  Height="425"
  ResizeMode="NoResize"
  Title="MSI Properties"
  FontSize="12">

  <DockPanel>
    <Menu DockPanel.Dock="Top">
      <MenuItem Header="Right Click Menu">
        <MenuItem Name="MenuItem_Install"
                  Header="Install"/>
        <MenuItem Name="MenuItem_Uninstall"
                  Header="Uninstall"/>
      </MenuItem>
      <MenuItem Header="About">
        <MenuItem Name="MenuItem_GitHub"
                  Header="GitHub - GetMSIInformation"/>
        <MenuItem Name="MenuItem_About"
                  Header="michaeltheadmin.com"/>
        <Separator/>
        <MenuItem Name="MenuItem_Version"
                  Header="Version 1.0.0"
                  IsEnabled="False"/>
      </MenuItem>
    </Menu>

    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="32"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="5"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="32"/>
        <RowDefinition Height="*"/>
      </Grid.RowDefinitions>
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="Auto"/>
        <ColumnDefinition Width="*"/>
        <ColumnDefinition Width="0.15*"/>
      </Grid.ColumnDefinitions>
      <Grid.Resources>
        <Style TargetType="Label">
          <Setter Property="Margin"
                  Value="2.5"/>
          <Setter Property="HorizontalAlignment"
                  Value="Stretch"/>
          <Setter Property="HorizontalContentAlignment"
                  Value="Right"/>
          <Setter Property="VerticalAlignment"
                  Value="Stretch"/>
          <Setter Property="VerticalContentAlignment"
                  Value="Center"/>
          <Setter Property="IsEnabled"
                  Value="True"/>
        </Style>
        <Style TargetType="TextBox">
          <Setter Property="Margin"
                  Value="2.5"/>
          <Setter Property="Width"
                  Value="Auto"/>
          <Setter Property="HorizontalAlignment"
                  Value="Stretch"/>
          <Setter Property="VerticalAlignment"
                  Value="Stretch"/>
          <Setter Property="VerticalContentAlignment"
                  Value="Center"/>
          <Setter Property="IsEnabled"
                  Value="True"/>
          <Setter Property="IsReadOnly"
                  Value="True"/>
        </Style>
        <Style TargetType="Button">
          <Setter Property="Margin"
                  Value="2.5"/>
          <Setter Property="Width"
                  Value="Auto"/>
          <Setter Property="HorizontalAlignment"
                  Value="Stretch"/>
          <Setter Property="VerticalAlignment"
                  Value="Stretch"/>
          <Setter Property="VerticalContentAlignment"
                  Value="Center"/>
          <Setter Property="IsEnabled"
                  Value="False"/>
        </Style>
        <Style TargetType="ListBoxItem">
          <Setter Property="HorizontalAlignment"
                  Value="Stretch"/>
          <Setter Property="HorizontalContentAlignment"
                  Value="Center"/>
          <Setter Property="VerticalAlignment"
                  Value="Stretch"/>
          <Setter Property="VerticalContentAlignment"
                  Value="Center"/>
          <Setter Property="Height"
                  Value="{Binding ElementName=lsbox_FilePath, Path=ActualHeight}"/>
        </Style>
      </Grid.Resources>

      <!-- Row 0 -->
      <!-- MD5 -->
      <Label
        Grid.Row="0"
        Grid.Column="0"
        Name="lbl_MD5"
        Content="MD5"/>
      <TextBox
        Grid.Row="0"
        Grid.Column="1"
        Name="txt_MD5"
        xml:space="preserve"/>
      <Button
        Grid.Row="0"
        Grid.Column="2"
        Name="btn_MD5_Copy"
        Content="Copy"/>

      <!-- Row 1 -->
      <!-- Row SHA1 -->
      <Label
        Grid.Row="1"
        Grid.Column="0"
        Name="lbl_SHA1"
        Content="SHA1"/>
      <TextBox
        Grid.Row="1"
        Grid.Column="1"
        Name="txt_SHA1"
        xml:space="preserve"/>
      <Button
        Grid.Row="1"
        Grid.Column="2"
        Name="btn_SHA1_Copy"
        Content="Copy"/>

      <!-- Row 2 -->
      <!-- Row SHA256 -->
      <Label
        Grid.Row="2"
        Grid.Column="0"
        Name="lbl_SHA256"
        Content="SHA256"/>
      <TextBox
        Grid.Row="2"
        Grid.Column="1"
        Name="txt_SHA256"
        xml:space="preserve"/>
      <Button
        Grid.Row="2"
        Grid.Column="2"
        Name="btn_SHA256_Copy"
        Content="Copy"/>

      <!-- Row 3 -->
      <!-- Digest -->
      <Label
        Grid.Row="3"
        Grid.Column="0"
        Name="lbl_Digest"
        Content="Digest"/>
      <TextBox
        Grid.Row="3"
        Grid.Column="1"
        Name="txt_Digest"
        xml:space="preserve"/>
      <Button
        Grid.Row="3"
        Grid.Column="2"
        Name="btn_Digest_Copy"
        Content="Copy"/>

      <!-- Row Gridline -->
      <!-- Row 4 -->
      <Line
        Grid.Row="4"
        Grid.Column="0"
        Grid.ColumnSpan="3"
        X1="0"
        Y1="0"
        X2="1"
        Y2="0"
        Stroke="Black"
        StrokeThickness="2"
        Stretch="Uniform"/>

      <!-- Row -->
      <Label
        Grid.Row="5"
        Grid.Column="0"
        Name="lbl_ProductName"
        Content="Product Name"/>
      <TextBox
        Grid.Row="5"
        Grid.Column="1"
        Name="txt_ProductName"/>
      <Button
        Grid.Row="5"
        Grid.Column="2"
        Name="btn_ProductName_Copy"
        Content="Copy"/>

      <!-- Row -->
      <Label
        Grid.Row="6"
        Grid.Column="0"
        Name="lbl_Manufacturer"
        Content="Manufacturer"/>
      <TextBox
        Grid.Row="6"
        Grid.Column="1"
        Name="txt_Manufacture"
        xml:space="preserve"/>
      <Button
        Grid.Row="6"
        Grid.Column="2"
        Name="btn_Manufacture_Copy"
        Content="Copy"/>

      <!-- Row -->
      <Label
        Grid.Row="7"
        Grid.Column="0"
        Name="lbl_ProductVersion"
        Content="Product Version"/>
      <TextBox
        Grid.Row="7"
        Grid.Column="1"
        Name="txt_ProductVersion"
        xml:space="preserve"/>
      <Button
        Grid.Row="7"
        Grid.Column="2"
        Name="btn_ProductVersion_Copy"
        Content="Copy"/>

      <!-- Row -->
      <Label
        Grid.Row="8"
        Grid.Column="0"
        Name="lbl_ProductCode"
        Content="Product Code"/>
      <TextBox
        Grid.Row="8"
        Grid.Column="1"
        Name="txt_ProductCode"
        xml:space="preserve"/>
      <Button
        Grid.Row="8"
        Grid.Column="2"
        Name="btn_ProductCode_Copy"
        Content="Copy"/>

      <!-- Row -->
      <Label
        Grid.Row="9"
        Grid.Column="0"
        Name="lbl_UpgradeCode"
        Content="Upgrade Code"/>
      <TextBox
        Grid.Row="9"
        Grid.Column="1"
        Name="txt_UpgradeCode"
        xml:space="preserve"/>
      <Button
        Grid.Row="9"
        Grid.Column="3"
        Name="btn_UpgradeCode_Copy"
        Content="Copy"/>

      <!-- Row -->
      <Button
        Grid.Row="10"
        Grid.Column="0"
        Name="btn_AllProperties"
        Margin="5"
        HorizontalAlignment="Stretch"
        VerticalAlignment="Stretch"
        Content="All Properties"
        Width="Auto"
        IsEnabled="False"/>
      <ListBox
        Grid.Row="10"
        Grid.Column="1"
        Name="lsbox_FilePath"
        Margin="5"
        HorizontalAlignment="Stretch"
        HorizontalContentAlignment="Center"
        VerticalAlignment="Stretch"
        VerticalContentAlignment="Center"
        AllowDrop="True"
        IsEnabled="True"
        TabIndex="0">
        <ListBox.Items>
          <ListBoxItem>
            <TextBlock Text="Drag and drop files here - *.msi"/>
          </ListBoxItem>
        </ListBox.Items>
      </ListBox>
      <Button
        Grid.Row="10"
        Grid.Column="3"
        Name="btn_FilePath_Copy"
        Content="Copy"/>

    </Grid>
  </DockPanel>
</Window>
"@

# Create a new XML node reader for reading the XAML content
$readerformMSIProperties = New-Object System.Xml.XmlNodeReader $XAMLformMSIProperties

# Load the XAML content into a WPF window object using the XAML reader
[System.Windows.Window]$formMSIProperties = [Windows.Markup.XamlReader]::Load($readerformMSIProperties)

# Create Variables for all the controls in the XAML form
$XAMLformMSIProperties.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $formMSIProperties.FindName($_.Name) -Scope Global }

#############################################
############## Event Handlers ###############
#############################################
#region Event Handlers

#### Form Load #####
$formMSIProperties.Add_Loaded({
    # Check if the script is running as an administrator
    if (($currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator))) {

      Write-Warning "The script is running as an administrator."
      Write-Warning "Drag and Drog will not work while running as an administrator."

      # Clear the listbox
      $lsbox_FilePath.Items.Clear()

      # Add a warning message to the listbox
      $lsbox_FilePath.Items.Add("WARNING: Running as Administrator | Drag and Drop will not work.")

      # Make the warning message bold and yellow
      $lsbox_FilePath.Background = [System.Windows.Media.Brushes]::Yellow
      $lsbox_FilePath.FontWeight = 'Bold'
    }

    # Update Version Information
    $formMSIProperties.Title = "MSI Properties - Version $($ScriptVersion)"
    $MenuItem_Version.Header = "Version $($ScriptVersion)"

    # Check if the FilePath parameter is provided to script
    if ($FilePath) {
      # Check if $FilePath is locked
      if (Test-FileLock -Path $FilePath) {
        Write-Warning "The file is locked: [$FilePath]"

        # Clear the listbox
        $lsbox_FilePath.Items.Clear()

        # Add an error message to the listbox
        $lsbox_FilePath.Items.Add("ERROR: The file is locked:`n[$FilePath]")
        
        # Make the Error message bold, red and yellow
        $lsbox_FilePath.Background = [System.Windows.Media.Brushes]::Red
        $lsbox_FilePath.Foreground = [System.Windows.Media.Brushes]::Yellow
        $lsbox_FilePath.FontWeight = 'Bold'
        $lsbox_FilePath.FontSize = 16
      }
      else {
        # Get the MSI file properties
        $FileMSIInfo = Get-MsiProperties -Path $FilePath

        # Get the File Hash Information
        $HashInfo = Get-FileHashInformation -Path $FilePath

        # Populate the textboxes
        Set-TextboxInformation -MSIPropertiesInfo $FileMSIInfo -FileHashInfo $HashInfo

        # Enable the Copy buttons
        Enable-AllButtons

        # Clear the listbox and add the filename
        $lsbox_FilePath.Items.Clear()
        $lsbox_FilePath.Items.Add($FilePath)
      }
    }
  })

#### Listbox Drag and Drop ####
$lsbox_FilePath.Add_Drop({
    $filename = $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)
    Write-Host "File Dropped: [$filename]"
    if ($filename) {
      # Check if $FilePath is locked
      if (Test-FileLock -Path "$($filename)") {
        Write-Warning "The file is locked: [$filename]"
  
        # Clear the listbox
        $lsbox_FilePath.Items.Clear()
  
        # Add an error message to the listbox
        $lsbox_FilePath.Items.Add("ERROR: The file is locked:`n[$filename]")
          
        # Make the Error message bold, red and yellow
        $lsbox_FilePath.Background = [System.Windows.Media.Brushes]::Red
        $lsbox_FilePath.Foreground = [System.Windows.Media.Brushes]::Yellow
        $lsbox_FilePath.FontWeight = 'Bold'
        $lsbox_FilePath.FontSize = 16
      }
      else {
        # Get the MSI file properties
        $FileMSIInfo = Get-MsiProperties -Path $filename

        # Get the File Hash Information
        $HashInfo = Get-FileHashInformation -Path $filename

        # Populate the textboxes
        Set-TextboxInformation -MSIPropertiesInfo $FileMSIInfo -FileHashInfo $HashInfo

        # Enable the Copy buttons
        Enable-AllButtons

        # Clear the listbox
        $lsbox_FilePath.Items.Clear()

        # Reset the listbox font style
        $lsbox_FilePath.ClearValue([System.Windows.Controls.Control]::BackgroundProperty)
        $lsbox_FilePath.ClearValue([System.Windows.Controls.Control]::ForegroundProperty)
        $lsbox_FilePath.ClearValue([System.Windows.Controls.Control]::FontWeightProperty)
        $lsbox_FilePath.ClearValue([System.Windows.Controls.Control]::FontSizeProperty)

        # Add the filename to the listbox
        $lsbox_FilePath.Items.Add($filename[0])
      }
    }
  })

$lsbox_FilePath.Add_DragOver({
    # Check if the dragged data contains file drop data
    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {
      foreach ($File in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) {
        # Check if the file is an MSI file
        if (([System.IO.Path]::GetExtension($File)) -eq ".msi") {
          # Set the drag effect to Copy if the file is an MSI file
          $_.Effects = [System.Windows.DragDropEffects]::Copy
        }
        else {
          # Set the drag effect to None if the file is not an MSI file
          $_.Effects = [System.Windows.DragDropEffects]::None
        }
      }
    }
  })

#### Menu Items ####  
$MenuItem_Install.add_Click({
    Write-Host "Menu Item Install Clicked"
    # Set Script Name
    $SaveAsScriptName = $ScriptName

    # Create a new directory in the LOCALAPPDATA folder
    Write-Host "Creating GetMSIInformation folder in LOCALAPPDATA folder"
    $DestinationFolderPath = "$env:LOCALAPPDATA\GetMSIInformation"
    if (-not (Test-Path $DestinationFolderPath)) {
      $DestinationFolder = New-Item -ItemType Directory -Path $DestinationFolderPath -ErrorAction SilentlyContinue
    }
    else {
      $DestinationFolder = Get-Item -Path $DestinationFolderPath
    }

    # Check if the script is being Invoked from the Internet
    if ($PSCommandPath -ne "") {
      # Copy the script to the new directory
      Write-Host "Copying Script to GetMSIInfo Folder"
      Copy-Item "$PSScriptRoot\$([System.IO.Path]::GetFileName($PSCommandPath))" -Destination "$($DestinationFolder.FullName)\$($SaveAsScriptName)" -ErrorAction SilentlyContinue
    }
    else {
      Write-Host "PSCommandPath is not available."
      # Script URL
      $ScriptURL = "https://raw.githubusercontent.com/MichaelEscamilla/GetMSIInformation/main/GetMSIInformation.ps1"
      Write-Host "Downloading the script from URL: [$ScriptURL]"
      try {
        Invoke-WebRequest -Uri $ScriptURL -OutFile "$($DestinationFolder.FullName)\$($SaveAsScriptName)" -ErrorAction Stop
        Write-Host "Script downloaded successfully saved: [$($DestinationFolder.FullName)\$($SaveAsScriptName)]"
      }
      catch {
        Write-Host "Failed to download the script: $_"
      }
    }

    # Reg2CI (c) 2020 by Roger Zander
    # https://github.com/asjimene/GetMSIInfo/blob/master/GetMSIInfo.ps1

    # Check if the registry path for .msi file associations exists, if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi") -ne $true) {
      New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi" -Force -ErrorAction SilentlyContinue 
    }

    # Check if the 'shell' subkey exists under the .msi file associations, if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell") -ne $true) {
      New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell" -Force -ErrorAction SilentlyContinue 
    }

    # Check if the 'Get MSI Information' subkey exists under 'shell', if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName") -ne $true) {
      New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName" -Force -ErrorAction SilentlyContinue 
    }

    # Set the 'icon' value under 'Get MSI Information' to a powershell.exe icon
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName" -Name 'icon' -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -PropertyType String -Force -ErrorAction SilentlyContinue

    # Check if the 'command' subkey exists under 'Get MSI Information', if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName\command") -ne $true) {
      New-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName\command" -Force -ErrorAction SilentlyContinue 
    }

    # Set the default value of the 'Get MSI Information' key to "Get MSI Information".
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName" -Name '(default)' -Value "$RightClickMenuName" -PropertyType String -Force -ea SilentlyContinue;

    # Set the default value of the 'command' key to execute a PowerShell script with the .msi file as an argument.
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName\command" -Name '(default)' -Value "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$($DestinationFolder.FullName)\$($SaveAsScriptName)`" -FilePath '%1'" -PropertyType String -Force -ErrorAction SilentlyContinue;
    Write-Host "Installation Complete"
  })

$MenuItem_Uninstall.add_Click({
    Write-Host "Menu Item Uninstall Clicked"
    Write-Output "Removing Script from LOCALAPPDATA"

    # Remove the script folder from the LOCALAPPDATA folder
    Remove-item "$env:LOCALAPPDATA\GetMSIInformation" -Force -Recurse -ErrorAction SilentlyContinue

    # Reg2CI (c) 2020 by Roger Zander
    # https://github.com/asjimene/GetMSIInfo/blob/master/GetMSIInfo.ps1


    Write-Output "Cleaning Up Registry"
    # Remove the 'Get MSI Information' registry key if it exists
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName") -eq $true) { 
      Remove-Item "HKCU:\Software\Classes\SystemFileAssociations\.msi\shell\$RightClickMenuName" -force -Recurse -ea SilentlyContinue 
    }

    Write-Output "Uninstallation Complete!"
  })

$MenuItem_GitHub.add_Click({
    # Open Github Project Page
    Start-Process "https://github.com/MichaelEscamilla/GetMSIInformation"
  })

$MenuItem_About.add_Click({
    # Open Blog
    Start-Process "https://michaeltheadmin.com"
  })

#### Button Handlers ####
$btn_AllProperties.add_Click({
    $SelectedProperty = Get-MsiProperties -Path $lsbox_FilePath.Items[0] | Out-GridView -Title "MSI Database Properties for $($lsbox_FilePath.Items[0])" -OutputMode Single
    $SelectedProperty.Value | Set-Clipboard
  })

$Button_Copy_Handler = {
  # Get the button name
  $ButtonName = $_.Source.Name
  # Get the property name from the button name by parsing between the underscores
  $PropertyName = [regex]::Match($ButtonName, "_(.*?)_").Groups[1].Value
  # Get the variable for the textbox with the same name as the property name
  $TextboxVariable = Get-Variable -Name "txt_$($propertyName)" -ValueOnly -ErrorAction SilentlyContinue
  if ($TextboxVariable) {
    Write-Host "Textbox [$($TextboxVariable.Name)] Value Copied to Clipboard : [$($TextboxVariable.Text)]"
    # Copy the text from the textbox with the same name as the property name
    [System.Windows.Forms.Clipboard]::SetText($TextboxVariable.Text)
  }
  else {
    # Try getting a Listbox variable with the same name as the property name
    $ListboxVariable = Get-Variable -Name "lsbox_$($propertyName)" -ValueOnly -ErrorAction SilentlyContinue
    if ($ListboxVariable) {
      # Check if the item in the listbox contains spaces
      if ($lsbox_FilePath.Items[0] -match "\s") {
        # Copy the item in the listbox to the clipboard with quotes
        [System.Windows.Forms.Clipboard]::SetText("`"$($lsbox_FilePath.Items[0])`"")
        Write-Host "Copied to Clipboard: [`"$($lsbox_FilePath.Items[0])`"]"
      }
      else {
        # Copy the item in the listbox to the clipboard without quotes
        [System.Windows.Forms.Clipboard]::SetText($lsbox_FilePath.Items[0])
        Write-Host "Copied to Clipboard: [$($lsbox_FilePath.Items[0])]"
      }
    }
  }
}

# Get all button variables that contain the word "Copy"
$Buttons = Get-Variable -Name "*Copy" -ValueOnly -ErrorAction SilentlyContinue
foreach ($Button in $Buttons) {
  # Add a click event handler to the button
  $Button.add_Click($Button_Copy_Handler)
}

#endregion Event Handlers

#Show the WPF Window
$formMSIProperties.WindowStartupLocation = "CenterScreen"
$formMSIProperties.ShowDialog() | Out-Null