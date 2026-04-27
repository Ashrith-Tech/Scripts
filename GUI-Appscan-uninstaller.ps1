Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Suppress console output from loops and operations
$ErrorActionPreference = "Continue"
$WarningPreference = "SilentlyContinue"
$DebugPreference = "SilentlyContinue"
$VerbosePreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

function Start-Clean {
    Clear-Host
    $null = Get-Date
}

# XAML definition with an extra row for the "Select All Servers" button
$XAML = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        Title="Remote App Uninstaller" Height="450" Width="800" ResizeMode="CanMinimize">
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>  <!-- Row 0: Applications textbox -->
            <RowDefinition Height="Auto"/>  <!-- Row 1: Server file selection -->
            <RowDefinition Height="Auto"/>  <!-- Row 2: Progress bar -->
            <RowDefinition Height="Auto"/>  <!-- Row 3: Status label -->
            <RowDefinition Height="*"/>     <!-- Row 4: Servers ListBox -->
            <RowDefinition Height="Auto"/>  <!-- Row 5: Select All Servers button -->
            <RowDefinition Height="Auto"/>  <!-- Row 6: Execution Mode & Buttons -->
        </Grid.RowDefinitions>

        <!-- Row 0: Applications to be Removed -->
        <TextBlock Grid.Row="0" Text="Applications to be Removed:" FontWeight="Bold" Margin="0,0,0,5"/>
        <TextBox Grid.Row="0" Name="AppTextBox" Margin="0,20,0,10" Height="70" Background="White" 
                 TextWrapping="Wrap" AcceptsReturn="True" VerticalScrollBarVisibility="Auto"/>

        <!-- Row 1: Server List File -->
        <Grid Grid.Row="1" Margin="0,0,0,5">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <TextBlock Grid.Column="0" Text="Select Server List File:" FontWeight="Bold" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <TextBox Grid.Column="1" Name="FilePathBox" IsReadOnly="True" VerticalAlignment="Center" Margin="0,0,10,0"/>
            <Button Grid.Column="2" Name="BrowseButton" Content="Browse" Width="80"/>
        </Grid>

        <!-- Row 2: Progress Bar -->
        <ProgressBar Grid.Row="2" Name="ProgressBar" Height="20" Minimum="0" Maximum="100" Margin="0,5,0,5" Background="#F0F0F0"/>

        <!-- Row 3: Status Label -->
        <TextBlock Grid.Row="3" Name="StatusLabel" FontWeight="Bold" Foreground="Blue" HorizontalAlignment="Center" Margin="0,5,0,5"/>

        <!-- Row 4: Servers ListBox -->
        <TextBlock Grid.Row="4" Text="Servers to Process:" FontWeight="Bold" Margin="0,0,0,5"/>
        <ListBox Grid.Row="4" Name="ServerList" SelectionMode="Extended" Margin="0,20,0,5" Background="White"/>

        <!-- Row 5: Select All Servers Button -->
        <Button Grid.Row="5" Name="SelectAllServersButton" Content="Select All Servers" Width="120" Height="25" 
                HorizontalAlignment="Left" Margin="0,0,0,5" />

        <!-- Row 6: Execution Mode & Buttons -->
        <Grid Grid.Row="6" Margin="0,10,0,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="Auto"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <!-- Execution Mode -->
            <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                <TextBlock Text="Execution Mode:" FontWeight="Bold" Margin="0,0,10,0"/>
                <RadioButton Name="SequentialMode" Content="Sequential" IsChecked="True" Margin="5,0,0,0"/>
                <RadioButton Name="ParallelMode" Content="Parallel" Margin="10,0,0,0"/>
            </StackPanel>
            <!-- Buttons Panel -->
            <StackPanel Grid.Column="2" Orientation="Horizontal" HorizontalAlignment="Right" Name="ButtonsPanel">
                <Button Name="GetInstalledAppsButton" Content="Get Installed Apps" Width="120" Margin="0,0,5,0"/>
                <!-- "Scan Apps" button will be inserted here -->
                <Button Name="UninstallButton" Content="Uninstall" Width="120" Margin="0,0,5,0"/>
                <Button Name="ExitButton" Content="Exit" Width="80" Margin="5,0,0,0"/>
            </StackPanel>
        </Grid>
    </Grid>
</Window>
"@

# Load the XAML
try {
    [xml]$XAMLObject = $XAML
    $Reader = New-Object System.Xml.XmlNodeReader($XAMLObject)
    $Window = [Windows.Markup.XamlReader]::Load($Reader)
    Start-Clean
} catch {
    [System.Windows.MessageBox]::Show("Failed to load XAML: $_", "Error", "OK", "Error")
    exit
}

# Get UI elements
try {
    $FilePathBox              = $Window.FindName("FilePathBox")
    $BrowseButton             = $Window.FindName("BrowseButton")
    $ServerList               = $Window.FindName("ServerList")
    $AppTextBox               = $Window.FindName("AppTextBox")
    $UninstallButton          = $Window.FindName("UninstallButton")
    $ExitButton               = $Window.FindName("ExitButton")
    $ProgressBar              = $Window.FindName("ProgressBar")
    $StatusLabel              = $Window.FindName("StatusLabel")
    $SequentialMode           = $Window.FindName("SequentialMode")
    $ParallelMode             = $Window.FindName("ParallelMode")
    $ButtonsPanel             = $Window.FindName("ButtonsPanel")
    $SelectAllServersButton   = $Window.FindName("SelectAllServersButton")
    $GetInstalledAppsButton   = $Window.FindName("GetInstalledAppsButton")
} catch {
    [System.Windows.MessageBox]::Show("Failed to get UI elements: $_", "Error", "OK", "Error")
    exit
}

# Insert the "Scan Apps" button into the ButtonsPanel (between GetInstalledApps and Uninstall)
$ScanAppsButton = New-Object System.Windows.Controls.Button
$ScanAppsButton.Name = "ScanAppsButton"
$ScanAppsButton.Content = "Scan Apps"
$ScanAppsButton.Width = 120
$ScanAppsButton.Margin = New-Object System.Windows.Thickness(0,0,5,0)
$ButtonsPanel.Children.Insert(1, $ScanAppsButton)

# Set up the default applications list in the textbox
$DefaultApps = @(
    "Google Chrome",
    "Adobe Acrobat (64-bit)",
    "Adobe Reader",
    "Notepad++ (64-bit x64)",
    "7-Zip 24.09 (x64 edition)",
    "Silverlight"
)
$AppTextBox.Text = ($DefaultApps -join "`r`n")

# Set a tooltip for the Applications textbox
$AppTextBoxTooltip = New-Object System.Windows.Controls.ToolTip
$AppTextBoxTooltip.Content = "Enter one application name (or keyword) per line. The Scan Apps button will find installed apps matching these keywords."
$AppTextBox.ToolTip = $AppTextBoxTooltip

# Define log file path
$LogFile = "C:\temp\Uninstall_Applications_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Function to update status and progress
function Update-Status {
    param([string]$Message, [int]$Progress = $null)
    $StatusLabel.Dispatcher.Invoke([action]{
        $StatusLabel.Text = $Message
        [System.Windows.Forms.Application]::DoEvents()
    }, [System.Windows.Threading.DispatcherPriority]::Render)
    if ($null -ne $Progress) {
        $ProgressBar.Dispatcher.Invoke([action]{
            $ProgressBar.Value = $Progress
            [System.Windows.Forms.Application]::DoEvents()
        }, [System.Windows.Threading.DispatcherPriority]::Render)
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Set up context menu for the Applications textbox
$AppContextMenu = New-Object System.Windows.Controls.ContextMenu
$MenuItem1 = New-Object System.Windows.Controls.MenuItem
$MenuItem1.Header = "Add Application"
$CommonApps = @(
    "Google Chrome",
    "Mozilla Firefox",
    "Adobe Acrobat Reader DC",
    "Adobe Acrobat (64-bit)",
    "7-Zip 24.09 (x64 edition)",
    "Microsoft Teams",
    "Notepad++ (64-bit x64)",
    "Microsoft Silverlight",
    "VLC media player",
    "Zoom",
    "Java 8 Update 401",
    "Microsoft OneDrive"
)
foreach ($App in $CommonApps) {
    $AppItem = New-Object System.Windows.Controls.MenuItem
    $AppItem.Header = $App
    $AppItem.Add_Click({
        $SelectedApp = $this.Header.ToString()
        $CurrentApps = $AppTextBox.Text -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
        if ($CurrentApps -notcontains $SelectedApp) {
            if ([string]::IsNullOrWhiteSpace($AppTextBox.Text)) {
                $AppTextBox.Text = $SelectedApp
            } else {
                $AppTextBox.Text += "`r`n" + $SelectedApp
            }
        }
    })
    $MenuItem1.Items.Add($AppItem)
}
$MenuItem2 = New-Object System.Windows.Controls.MenuItem
$MenuItem2.Header = "Clear All"
$MenuItem2.Add_Click({ $AppTextBox.Clear() })
$AppContextMenu.Items.Add($MenuItem1)
$AppContextMenu.Items.Add($MenuItem2)
$AppTextBox.ContextMenu = $AppContextMenu

# Define registry paths for installed applications
$RegistryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
)

# Function to get the silent uninstall command from the uninstall string
function Get-SilentUninstallCommand {
    param([string]$UninstallCmd)
    if ($UninstallCmd -match "unins000.exe") {
        return "$UninstallCmd /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-"
    }
    elseif ($UninstallCmd -match "MsiExec.exe") {
        $UninstallCmd = $UninstallCmd -replace "/I", "/X"
        if ($UninstallCmd -notmatch "/quiet") {
            return "$UninstallCmd /quiet /norestart"
        }
    }
    else {
        if ($UninstallCmd -notmatch "/S|/silent|/quiet") {
            return "$UninstallCmd /S"
        }
    }
    return $UninstallCmd
}

# Browse button: load servers from a file
$BrowseButton.Add_Click({
    try {
        $FileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $FileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        $FileDialog.Title = "Select Server List File"
        if ($FileDialog.ShowDialog() -eq "OK") {
            $FilePathBox.Text = $FileDialog.FileName
            $ServerList.Items.Clear()
            if (Test-Path $FileDialog.FileName) {
                $Servers = Get-Content $FileDialog.FileName -ErrorAction Stop
                foreach ($Server in $Servers) {
                    if (-not [string]::IsNullOrWhiteSpace($Server)) {
                        $ServerList.Items.Add($Server.Trim())
                    }
                }
                Update-Status "Server list loaded successfully!"
            } else {
                Update-Status "File not found!"
            }
        }
    }
    catch {
        Update-Status "Error loading server list: $($_.Exception.Message)"
    }
})

# "Select All Servers" button: select every server in the list
$SelectAllServersButton.Add_Click({
    if ($ServerList.Items.Count -eq 0) {
        Update-Status "No servers to select. Please load a server list first."
        return
    }
    $ServerList.SelectedItems.Clear()
    for ($i = 0; $i -lt $ServerList.Items.Count; $i++) {
        $ServerList.SelectedItems.Add($ServerList.Items[$i])
    }
    Update-Status "All servers selected."
})

# "Scan Apps" button: memory-friendly scan over selected servers
$ScanAppsButton.Add_Click({
    $SelectedServers = $ServerList.SelectedItems
    if (-not $SelectedServers -or $SelectedServers.Count -eq 0) {
        Update-Status "Please select one or more servers from the list first."
        return
    }
    # Get current keywords (each line) from the AppTextBox
    $Keywords = $AppTextBox.Text -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
    if ($Keywords.Count -eq 0) {
        Update-Status "No keywords found in the Applications list. Please add at least one line."
        return
    }
    Update-Status "Scanning for matching applications on selected servers..."
    # List to store only matched friendly names
    $MatchedApps = New-Object System.Collections.Generic.List[String]
    try {
        foreach ($Server in $SelectedServers) {
            $InstalledApps = Invoke-Command -ComputerName $Server -ScriptBlock {
                $Apps = @()
                $Paths = @(
                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                )
                foreach ($Path in $Paths) {
                    if (Test-Path $Path) {
                        Get-ChildItem -Path $Path -ErrorAction SilentlyContinue | ForEach-Object {
                            $AppName = $_.GetValue("DisplayName")
                            if ($AppName) {
                                $Apps += $AppName
                            }
                        }
                    }
                }
                return $Apps
            } -ErrorAction Stop
            foreach ($Keyword in $Keywords) {
                $LocalMatches = $InstalledApps | Where-Object { $_ -ilike "*$Keyword*" }
                foreach ($Match in $LocalMatches) {
                    if (-not $MatchedApps.Contains($Match)) {
                        [void]$MatchedApps.Add($Match)
                    }
                }
            }
            # Free memory from this server's app list
            $InstalledApps = $null
            [System.GC]::Collect()
        }
        # Append newly found matches to the AppTextBox if not already listed
        $ExistingEntries = $AppTextBox.Text -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
        $AddedCount = 0
        foreach ($AppName in $MatchedApps) {
            if ($ExistingEntries -notcontains $AppName) {
                $AppTextBox.Text += "`r`n$AppName"
                $AddedCount++
            }
        }
        if ($AddedCount -gt 0) {
            Update-Status "Scan complete. Added $AddedCount new application(s)."
        } else {
            Update-Status "Scan complete. No additional applications found."
        }
    }
    catch {
        Update-Status "Error during scan: $($_.Exception.Message)"
    }
})

# "Get Installed Apps" button: display a selection dialog for one server
$GetInstalledAppsButton.Add_Click({
    $Server = $ServerList.SelectedItem
    if (-not $Server) {
        Update-Status "Please select a server from the list first."
        return
    }
    Update-Status "Retrieving installed applications from $Server..."
    try {
        $InstalledApps = Invoke-Command -ComputerName $Server -ScriptBlock {
            $Apps = @()
            $RegistryPaths = @(
                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
            )
            foreach ($Path in $RegistryPaths) {
                if (Test-Path $Path) {
                    Get-ChildItem -Path $Path -ErrorAction SilentlyContinue | ForEach-Object {
                        $AppName = $_.GetValue("DisplayName")
                        $DisplayVersion = $_.GetValue("DisplayVersion")
                        $UninstallString = $_.GetValue("UninstallString")
                        if ($AppName -and $UninstallString) {
                            $Apps += [PSCustomObject]@{
                                Name            = $AppName
                                Version         = $DisplayVersion
                                UninstallString = $UninstallString
                            }
                        }
                    }
                }
            }
            return $Apps | Sort-Object Name
        } -ErrorAction Stop

        # Create selection dialog
        $AppSelectDialog = New-Object System.Windows.Window
        $AppSelectDialog.Title = "Select Applications to Uninstall"
        $AppSelectDialog.Width = 500
        $AppSelectDialog.Height = 600
        $AppSelectDialog.WindowStartupLocation = "CenterScreen"
        $AppSelectDialog.ResizeMode = "CanResize"
        $DockPanel = New-Object System.Windows.Controls.DockPanel
        $AppSelectDialog.Content = $DockPanel
        $TextBlock = New-Object System.Windows.Controls.TextBlock
        $TextBlock.Text = "Applications installed on $Server"
        $TextBlock.FontWeight = "Bold"
        $TextBlock.Margin = New-Object System.Windows.Thickness(5,5,5,5)
        [System.Windows.Controls.DockPanel]::SetDock($TextBlock, "Top")
        $DockPanel.Children.Add($TextBlock)
        $ButtonPanel = New-Object System.Windows.Controls.StackPanel
        $ButtonPanel.Orientation = "Horizontal"
        $ButtonPanel.HorizontalAlignment = "Right"
        $ButtonPanel.Margin = New-Object System.Windows.Thickness(5,5,5,5)
        [System.Windows.Controls.DockPanel]::SetDock($ButtonPanel, "Bottom")
        $DockPanel.Children.Add($ButtonPanel)
        $AddButton = New-Object System.Windows.Controls.Button
        $AddButton.Content = "Add Selected"
        $AddButton.Width = 100
        $AddButton.Margin = New-Object System.Windows.Thickness(5,0,5,0)
        $ButtonPanel.Children.Add($AddButton)
        $CancelButton = New-Object System.Windows.Controls.Button
        $CancelButton.Content = "Cancel"
        $CancelButton.Width = 80
        $CancelButton.Margin = New-Object System.Windows.Thickness(5,0,5,0)
        $ButtonPanel.Children.Add($CancelButton)
        $ListBox = New-Object System.Windows.Controls.ListBox
        $ListBox.SelectionMode = "Extended"
        $ListBox.Margin = New-Object System.Windows.Thickness(5,5,5,5)
        [System.Windows.Controls.DockPanel]::SetDock($ListBox, "Top")
        $DockPanel.Children.Add($ListBox)
        $FilterPanel = New-Object System.Windows.Controls.StackPanel
        $FilterPanel.Orientation = "Horizontal"
        $FilterPanel.Margin = New-Object System.Windows.Thickness(5,5,5,5)
        [System.Windows.Controls.DockPanel]::SetDock($FilterPanel, "Top")
        $DockPanel.Children.Add($FilterPanel)
        $FilterLabel = New-Object System.Windows.Controls.TextBlock
        $FilterLabel.Text = "Filter:"
        $FilterLabel.VerticalAlignment = "Center"
        $FilterLabel.Margin = New-Object System.Windows.Thickness(0,0,5,0)
        $FilterPanel.Children.Add($FilterLabel)
        $FilterTextBox = New-Object System.Windows.Controls.TextBox
        $FilterTextBox.Width = 350
        $FilterPanel.Children.Add($FilterTextBox)
        foreach ($App in $InstalledApps) {
            if ([string]::IsNullOrWhiteSpace($App.Version)) {
                $ListBox.Items.Add($App.Name)
            }
            else {
                $ListBox.Items.Add("$($App.Name) (v$($App.Version))")
            }
        }
        $FilterTextBox.Add_TextChanged({
            $Filter = $FilterTextBox.Text.ToLower()
            $ListBox.Items.Clear()
            if ([string]::IsNullOrWhiteSpace($Filter)) {
                foreach ($App in $InstalledApps) {
                    if ([string]::IsNullOrWhiteSpace($App.Version)) {
                        $ListBox.Items.Add($App.Name)
                    }
                    else {
                        $ListBox.Items.Add("$($App.Name) (v$($App.Version))")
                    }
                }
            }
            else {
                foreach ($App in $InstalledApps) {
                    $ItemText = if ([string]::IsNullOrWhiteSpace($App.Version)) { $App.Name } else { "$($App.Name) (v$($App.Version))" }
                    if ($ItemText.ToLower().Contains($Filter)) {
                        $ListBox.Items.Add($ItemText)
                    }
                }
            }
        })
        $AddButton.Add_Click({
            $SelectedItems = $ListBox.SelectedItems
            if ($SelectedItems.Count -gt 0) {
                $CurrentApps = $AppTextBox.Text -split "`r?`n" | Where-Object { $_.Trim() -ne "" }
                $NewApps = @()
                foreach ($Item in $SelectedItems) {
                    $AppName = $Item -replace "\s+\(v.*\)$", ""
                    if ($CurrentApps -notcontains $AppName) {
                        $NewApps += $AppName
                    }
                }
                if ($NewApps.Count -gt 0) {
                    if ([string]::IsNullOrWhiteSpace($AppTextBox.Text)) {
                        $AppTextBox.Text = $NewApps -join "`r`n"
                    }
                    else {
                        $AppTextBox.Text += "`r`n" + ($NewApps -join "`r`n")
                    }
                }
                $AppSelectDialog.DialogResult = $true
                $AppSelectDialog.Close()
            }
        })
        $CancelButton.Add_Click({
            $AppSelectDialog.DialogResult = $false
            $AppSelectDialog.Close()
        })
        $AppSelectDialog.ShowDialog() | Out-Null
        Update-Status "Ready"
    }
    catch {
        Update-Status "Error: $($_.Exception.Message)"
    }
})

# "Uninstall" button: uninstall apps on each server
$UninstallButton.Add_Click({
    Update-Status "Starting uninstallation process..." 0
    # Here, using all servers; change to $ServerList.SelectedItems if needed.
    $Servers = @($ServerList.Items)
    if ($Servers.Count -eq 0) {
        Update-Status "No servers loaded. Please select a server list file."
        return
    }
    # Get applications from the textbox; remove version info if present.
    $AppsToRemove = $AppTextBox.Text -split "`r?`n" | ForEach-Object { $_ -replace "\s+\(v.*\)$", "" } | Where-Object { $_.Trim() -ne "" }
    if ($AppsToRemove.Count -eq 0) {
        Update-Status "No applications specified. Please enter or scan for apps to uninstall."
        return
    }
    # Disable buttons during operation
    $UninstallButton.IsEnabled         = $false
    $BrowseButton.IsEnabled            = $false
    $GetInstalledAppsButton.IsEnabled  = $false
    $ScanAppsButton.IsEnabled          = $false
    $SelectAllServersButton.IsEnabled  = $false
    try {
        $Results = @()
        $totalOperations = $Servers.Count * $AppsToRemove.Count
        $currentOperation = 0
        foreach ($Server in $Servers) {
            $currentProgress = [int](($currentOperation / $totalOperations) * 100)
            Update-Status "Checking installed apps on $Server... ($($currentOperation+1) of $totalOperations)" $currentProgress
            Start-Sleep -Milliseconds 100
            $InstalledApps = Invoke-Command -ComputerName $Server -ScriptBlock {
                $ApplicationsFound = @()
                $RegPaths = @(
                    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                )
                foreach ($Path in $RegPaths) {
                    if (Test-Path $Path) {
                        Get-ChildItem -Path $Path -ErrorAction SilentlyContinue | ForEach-Object {
                            $AppName = $_.GetValue("DisplayName")
                            $DisplayVersion = $_.GetValue("DisplayVersion")
                            $UninstallString = $_.GetValue("UninstallString")
                            if ($AppName -and $UninstallString) {
                                $ApplicationsFound += [PSCustomObject]@{
                                    Name            = $AppName
                                    Version         = $DisplayVersion
                                    UninstallString = $UninstallString
                                }
                            }
                        }
                    }
                }
                return $ApplicationsFound
            } -ErrorAction SilentlyContinue
            foreach ($App in $AppsToRemove) {
                $currentOperation++
                $progress = [int](($currentOperation / $totalOperations) * 100)
                Update-Status "Processing server $Server - App: $App" $progress
                $AppEntry = $InstalledApps | Where-Object { $_.Name -eq $App }
                if ($AppEntry) {
                    $VersionInfo = if ([string]::IsNullOrWhiteSpace($AppEntry.Version)) { "" } else { " (v$($AppEntry.Version))" }
                    Update-Status "Uninstalling $App$VersionInfo on $Server..." $progress
                    [System.Windows.Forms.Application]::DoEvents()
                    $SilentCmd = Get-SilentUninstallCommand -UninstallCmd $AppEntry.UninstallString
                    if (-not [string]::IsNullOrWhiteSpace($SilentCmd)) {
                        Invoke-Command -ComputerName $Server -ScriptBlock {
                            param ($SilentCmd, $AppName)
                            Start-Process -FilePath "cmd.exe" -ArgumentList "/c $SilentCmd" -NoNewWindow -Wait
                        } -ArgumentList $SilentCmd, $App -ErrorAction SilentlyContinue
                        Start-Sleep -Seconds 5
                        # Verify uninstallation
                        $Verified = Invoke-Command -ComputerName $Server -ScriptBlock {
                            param ($AppName)
                            $Found = $false
                            $RegPaths = @(
                                "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
                                "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
                            )
                            foreach ($Path in $RegPaths) {
                                if (Test-Path $Path) {
                                    if (Get-ChildItem -Path $Path | Where-Object { $_.GetValue("DisplayName") -eq $AppName }) {
                                        $Found = $true
                                        break
                                    }
                                }
                            }
                            return -not $Found
                        } -ArgumentList $App -ErrorAction SilentlyContinue
                        if ($Verified) {
                            Update-Status "$App removed successfully from $Server!" $progress
                            # Cleanup traces
                            Invoke-Command -ComputerName $Server -ScriptBlock {
                                param ($AppName)
                                $PathsToDelete = @(
                                    "C:\Program Files\$AppName",
                                    "C:\Program Files (x86)\$AppName",
                                    "$env:APPDATA\$AppName",
                                    "$env:LOCALAPPDATA\$AppName"
                                )
                                foreach ($p in $PathsToDelete) {
                                    if (Test-Path $p) {
                                        Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
                                    }
                                }
                            } -ArgumentList $App -ErrorAction SilentlyContinue
                            $Status = "Uninstalled"
                        } else {
                            Update-Status "Failed to remove $App from $Server!" $progress
                            $Status = "Failed"
                        }
                    } else {
                        Update-Status "No valid uninstall command found for $App on $Server" $progress
                        $Status = "Failed - No Valid Command"
                    }
                } else {
                    Update-Status "$App not found on $Server" $progress
                    $Status = "Not Installed"
                }
                $Results += [PSCustomObject]@{
                    Server      = $Server
                    Application = $App
                    Version     = if ($AppEntry) { $AppEntry.Version } else { "N/A" }
                    Status      = $Status
                    Timestamp   = Get-Date
                }
            }
        }
        $Results | Export-Csv -Path $LogFile -NoTypeInformation -Force
        Update-Status "Uninstallation complete! Results saved to: $LogFile" 100
    }
    catch {
        Update-Status "Error: $($_.Exception.Message)" 100
    }
    finally {
        $UninstallButton.IsEnabled         = $true
        $BrowseButton.IsEnabled            = $true
        $GetInstalledAppsButton.IsEnabled  = $true
        $ScanAppsButton.IsEnabled          = $true
        $SelectAllServersButton.IsEnabled  = $true
    }
})

# Exit button: close the GUI
$ExitButton.Add_Click({ $Window.Close() })

# On window load, update status and clear console
$Window.Add_Loaded({
    Update-Status "Ready - Select a server list file to begin" 0
    Clear-Host
})

# Show the window
$Window.ShowDialog() | Out-Null
