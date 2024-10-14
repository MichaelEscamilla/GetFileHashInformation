<#
.SYNOPSIS
This script provides a graphical user interface (GUI) for viewing and copying File Hash Information of files.

.DESCRIPTION
The script creates a WPF-based GUI that allows users to drag and drop files to view their File Hash Information.
The script supports MD5, SHA1, and SHA256 hash algorithms.
It also provides functionality to copy these properties to the clipboard and to clear the displayed information.
Additionally, the script includes options to install and uninstall a context menu item for files to retrieve their properties.

.PARAMETER FilePath
Optional parameter to specify the path of the file to automatically load the information for.

.NOTES
Author: Michael Escamilla
Date: 10-13-2024

Version History:
2024-10.13.0- Initial release of the GetFileHashInformation.ps1 script.
#>

param (
  [Parameter(Mandatory = $false)]
  [string]$FilePath
)

#############################################
################# Variables #################
#############################################
# Script Name
$Global:ScriptName = "GetFileHashInformation.ps1"
# Script Version
[System.Version]$Global:ScriptVersion = "2024.10.13.0"
# Right-Click Menu Name
$Global:RightClickMenuName = "Get File Hash Information"
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
    [hashtable]$FileHashInfo
  )

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
[xml]$XAMLformFileHashProperties = @"
<Window
  xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
  xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
  Name="form1"
  Width="900"
  Height="250"
  ResizeMode="NoResize"
  Title="Get File Hash Information"
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
                  Header="GitHub - GetFileHashInformation"/>
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

      <!-- Row -->
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
            <TextBlock Text="Drag and drop files here"/>
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
$readerformFileHashProperties = New-Object System.Xml.XmlNodeReader $XAMLformFileHashProperties

# Load the XAML content into a WPF window object using the XAML reader
[System.Windows.Window]$formFileHashProperties = [Windows.Markup.XamlReader]::Load($readerformFileHashProperties)

# Create Variables for all the controls in the XAML form
$XAMLformFileHashProperties.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $formFileHashProperties.FindName($_.Name) -Scope Global }

#############################################
############## Event Handlers ###############
#############################################
#region Event Handlers

#### Form Load #####
$formFileHashProperties.Add_Loaded({
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
    $formFileHashProperties.Title = "Get File Hash Information - Version $($ScriptVersion)"
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
       # Get the File Hash Information
        $HashInfo = Get-FileHashInformation -Path $FilePath

        # Populate the textboxes
        Set-TextboxInformation -FileHashInfo $HashInfo

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
        # Get the File Hash Information
        $HashInfo = Get-FileHashInformation -Path $filename

        # Populate the textboxes
        Set-TextboxInformation -FileHashInfo $HashInfo

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
        # Set the drag effect to Copy
        $_.Effects = [System.Windows.DragDropEffects]::Copy
      }
    }
  })

#### Menu Items ####  
$MenuItem_Install.add_Click({
    Write-Host "Menu Item Install Clicked"
    # Set Script Name
    $SaveAsScriptName = $ScriptName

    # Create a new directory in the LOCALAPPDATA folder
    Write-Host "Creating $([System.IO.Path]::GetFileNameWithoutExtension($ScriptName)) folder in LOCALAPPDATA folder"
    $DestinationFolderPath = "$env:LOCALAPPDATA\$([System.IO.Path]::GetFileNameWithoutExtension($ScriptName))"
    if (-not (Test-Path $DestinationFolderPath)) {
      $DestinationFolder = New-Item -ItemType Directory -Path $DestinationFolderPath -ErrorAction SilentlyContinue
    }
    else {
      $DestinationFolder = Get-Item -Path $DestinationFolderPath
    }

    # Check if the script is being Invoked from the Internet
    if ($PSCommandPath -ne "") {
      # Copy the script to the new directory
      Write-Host "Copying Script to $([System.IO.Path]::GetFileNameWithoutExtension($ScriptName)) Folder"
      Copy-Item "$PSScriptRoot\$([System.IO.Path]::GetFileName($PSCommandPath))" -Destination "$($DestinationFolder.FullName)\$($SaveAsScriptName)" -ErrorAction SilentlyContinue
    }
    else {
      Write-Host "PSCommandPath is not available."
      # Script URL
      $ScriptURL = "https://raw.githubusercontent.com/MichaelEscamilla/GetFileHashInformation/main/GetFileHashInformation.ps1"
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

    # Check if the registry path for * file associations exists, if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\*") -ne $true) {
      New-Item "HKCU:\Software\Classes\*" -Force -ErrorAction SilentlyContinue 
    }

    # Check if the 'shell' subkey exists under the * file associations, if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\*\shell") -ne $true) {
      New-Item "HKCU:\Software\Classes\*\shell" -Force -ErrorAction SilentlyContinue 
    }

    # Check if the 'Get File Hash Information' subkey exists under 'shell', if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\*\shell\$RightClickMenuName") -ne $true) {
      New-Item "HKCU:\Software\Classes\*\shell\$RightClickMenuName" -Force -ErrorAction SilentlyContinue 
    }

    # Set the 'icon' value under 'Get File Hash Information' to a powershell.exe icon
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\*\shell\$RightClickMenuName" -Name 'icon' -Value "C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe" -PropertyType String -Force -ErrorAction SilentlyContinue

    # Check if the 'command' subkey exists under 'Get File Hash Information', if not, create it.
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\*\shell\$RightClickMenuName\command") -ne $true) {
      New-Item "HKCU:\Software\Classes\*\shell\$RightClickMenuName\command" -Force -ErrorAction SilentlyContinue 
    }

    # Set the default value of the 'Get File Hash Information' key to "Get File Hash Information".
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\*\shell\$RightClickMenuName" -Name '(default)' -Value "$RightClickMenuName" -PropertyType String -Force -ea SilentlyContinue;

    # Set the default value of the 'command' key to execute a PowerShell script with the * file as an argument.
    New-ItemProperty -LiteralPath "HKCU:\Software\Classes\*\shell\$RightClickMenuName\command" -Name '(default)' -Value "C:\Windows\system32\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -Command `"$($DestinationFolder.FullName)\$($SaveAsScriptName)`" -FilePath '%1'" -PropertyType String -Force -ErrorAction SilentlyContinue;
    Write-Host "Installation Complete"
  })

$MenuItem_Uninstall.add_Click({
    Write-Host "Menu Item Uninstall Clicked"
    Write-Output "Removing Script from LOCALAPPDATA"

    # Remove the script folder from the LOCALAPPDATA folder
    Remove-item "$env:LOCALAPPDATA\$([System.IO.Path]::GetFileNameWithoutExtension($ScriptName))" -Force -Recurse -ErrorAction SilentlyContinue

    # Reg2CI (c) 2020 by Roger Zander
    # https://github.com/asjimene/GetMSIInfo/blob/master/GetMSIInfo.ps1


    Write-Output "Cleaning Up Registry"
    # Remove the 'Get FIle Hash Information' registry key if it exists
    if ((Test-Path -LiteralPath "HKCU:\Software\Classes\*\shell\$RightClickMenuName") -eq $true) { 
      Remove-Item "HKCU:\Software\Classes\*\shell\$RightClickMenuName" -force -Recurse -ea SilentlyContinue 
    }

    Write-Output "Uninstallation Complete!"
  })

$MenuItem_GitHub.add_Click({
    # Open Github Project Page
    Start-Process "https://github.com/MichaelEscamilla/GetFileHashInformation"
  })

$MenuItem_About.add_Click({
    # Open Blog
    Start-Process "https://michaeltheadmin.com"
  })

#### Button Handlers ####

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
$formFileHashProperties.WindowStartupLocation = "CenterScreen"
$formFileHashProperties.ShowDialog() | Out-Null