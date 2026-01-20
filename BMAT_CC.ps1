#region FUNCTIONS

Function Write-HostAndLog {
    [cmdletbinding()]
    param (
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [parameter(Mandatory = $true)]
        [ValidateSet("WARNING", "INFO", "ERROR")]
        [ValidateNotNullOrEmpty()]
        [string]$Category,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$LogfilePath,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [switch]$SuppressDisplay,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [switch]$SuppressLogFile,
        [parameter(Mandatory = $false, HelpMessage = "Debug Logging Level")]
        [ValidateSet($true, $false)]
        [string]$DetailedLogging = $true
    )
    IF ($AsTask) {
        $SuppressDisplay = $true 
    }
    IF ($NoLog) {
        $SuppressLogFile = $true 
    }
    IF ($DetailedLogging -eq $false) {
        $SuppressLogFile = $true
        $SuppressDisplay = $true
    }
    IF (-not $SuppressDisplay) {
        IF ($Category -eq "WARNING") {
            Write-Host $Message -ForegroundColor Yellow
        } ELSEIF ($Category -eq "ERROR") {
            Write-Host $Message -ForegroundColor DarkYellow
        } ELSEIF ($Category -eq "INFO") {
            Write-Host $Message -ForegroundColor Green
        }
    }
    IF (-not $SuppressLogFile) {
        Add-Content -Value ("{0} - {1}: {2}" -f ($(Get-Date), $Category, $Message)) -Path $LogfilePath
    }
}

Function Test-FileLock {
    Param(
        [parameter(Mandatory = $True)]
        [string]$Path
    )
    $OFile = New-Object System.IO.FileInfo $Path
    If ((Test-Path -LiteralPath $Path -PathType Leaf) -eq $False) {
        Return $False 
    } Else {
        Try {
            $OStream = $OFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            If ($OStream) {
                $OStream.Close() 
            }
            Return $False
        } Catch {
            Return $True 
        }
    }
}

Function Get-ConfigurationVariables {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $true)]
        [string]$ConfigFileName
    )
    $ConfigFilePath = Join-Path -Path $BMATFolder -ChildPath $ConfigFileName
    If (Test-Path -Path $ConfigFilePath) {
        $ConfigurationContent = Get-Content -Path $ConfigFilePath -Raw | ConvertFrom-Json -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    } Else {
        Return $false
    }
    If ($ConfigurationContent) {
        foreach ($Property in $ConfigurationContent.psobject.Properties) {
            if ($Property.Name -eq '_comment') { continue }
            try {
                Set-Variable -Name $Property.Name -Value $Property.Value -Scope Global -Force
            } catch {
                Write-Host ("Failure with step. Error message:{0}" -f $($_.Exception.Message)) -ForegroundColor DarkYellow
                Return $false
            }
        }
        try {
            Set-Variable -Name 'GameName' -Value 'fallout4' -Scope Global -Force
        } catch {
            Write-Host ("Failure with step. Error message:{0}" -f $($_.Exception.Message)) -ForegroundColor DarkYellow
            Return $false
        }
        Return $true
    } Else {
        Return $false
    }
}

Function Convert-FromVDFToJSON {
    Param(
        [parameter(Mandatory = $True)]
        $Content
    )
    $jsonText = ($Content -join "`n").Trim()
    $jsonText = $jsonText -replace '"\s*\{', '": {'
    $jsonText = $jsonText -replace '"[ \t]+"', '": "'
    $jsonText = $jsonText -replace '(?<="|\})\s+(?=")', ",`n"
    if (-not $jsonText.StartsWith("{")) { $jsonText = "{" + $jsonText + "}" }
    try {
        Return $jsonText | ConvertFrom-Json
    } catch {
        Write-Host ("Failure with step. Error message:{0}" -f $($_.Exception.Message)) -ForegroundColor DarkYellow
        Write-Host ("JSON Conversion failed. Current output state:`n{0}" -f $($jsonText | Out-String)) -ForegroundColor DarkYellow
    }
}

function Set-Logging {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder
    )
    $LogFileName = "{0}-{1}.log" -f ($LogFileName, (Get-Date -UFormat "%d-%m-%Y_%H-%M-%S"))
    $Global:LogFilePath = Join-Path -Path $LogsFolder -ChildPath $LogFileName
    If (Test-Path -LiteralPath $LogsFolder) {
        $LogFiles = Get-ChildItem $LogsFolder -Filter *.log | Where-Object LastWriteTime -lt (Get-Date).AddDays(-1)
        IF ($LogFiles) {
            $LogArchiveDestinationPath = Join-Path -Path $LogsFolder -ChildPath ("Log_Archive_{0}.zip" -f (Get-Date -format "ddMMyyyy"))
            If (Test-Path -LiteralPath $LogArchiveDestinationPath) {
                Compress-Archive -Path $LogFiles.FullName -DestinationPath $LogArchiveDestinationPath -CompressionLevel Optimal -Update
                ForEach ($File in $LogFiles) {
                    Remove-Item -LiteralPath $File.FullName
                }
            } Else {
                Compress-Archive -Path $LogFiles.FullName -DestinationPath $LogArchiveDestinationPath -CompressionLevel Optimal
                ForEach ($File in $LogFiles) {
                    Remove-Item -LiteralPath $File.FullName
                }
            }
        }
        $ArchiveFiles = Get-ChildItem $LogsFolder -Filter *.zip | Where-Object LastWriteTime -lt (Get-Date).AddDays(-3)
        IF ($ArchiveFiles) {
            ForEach ($File in $ArchiveFiles) {
                Remove-Item -LiteralPath $File.FullName
            }
        }
    } Else {
        Write-Host ("BMAT logs folder `"{0}`" not found." -f ($LogsFolder)) -ForegroundColor DarkYellow
        Return $false
    }
    If (Test-Path -LiteralPath $LogFilePath) {
        Try {
            Write-HostAndLog -Message "Initiating data gathering and analysis stage of the process" -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            Return $true
        } Catch {
            Write-Host ("Failure with step. Error message:{0}" -f $($_.Exception.Message)) -ForegroundColor DarkYellow
            Return $false
        }
    } Else {
        Try {
            Write-HostAndLog -Message "Log file created. Initiating data gathering and analysis stage of the process" -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            Return $true
        } Catch {
            Write-Host ("Failure with step. Error message:{0}" -f $($_.Exception.Message)) -ForegroundColor DarkYellow
            Return $false
        }
    }
}

function Update-ConfigFile {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $true)]
        [string]$ConfigFileName,
        [parameter(Mandatory = $true)]
        [string]$Property,
        [parameter(Mandatory = $true)]
        [string]$Value
    )
    $ConfigFilePath = Join-Path -Path $BMATFolder -ChildPath $ConfigFileName
    $RefinedValue = $Value.replace('\', '\\')
    Try {
        $ConfigurationContent = Get-Content -Path $ConfigFilePath -Raw
        $ConfigurationContent -replace (('"' + $Property + '": "",'), ('"' + $Property + '": "' + $RefinedValue + '",')) | Out-File -FilePath $ConfigFilePath -Force
    } Catch {
        Return $false
    }
    If ((Get-Content -Path $ConfigFilePath -Raw | ConvertFrom-Json).$Property -ne $Value) {
        Return $false
    } else {
        Return $true
    }
}

function Get-FileType {
    param (
        [string]$FileName
    )
    $Extension = [System.IO.Path]::GetExtension($FileName).ToLower()
    $FileType = switch -Wildcard ($Extension) {
        ".exe" { "Executable File" }
        ".log" { "Log File" }
        ".txt" { "Text Document" }
        ".json" { "Configuration File" }
        ".esl" { "Light Master Plugin" }
        ".esm" { "Master Plugin" }
        ".esp" { "Mod Plugin" }
        ".ba2" { "Game Archive" }
        ".bsa" { "Game Archive" }
        ".png" { "Image File" }
        ".jpg" { "Image File" }
        Default { "Unknown File Type" }
    }
    Return $FileType
}

function Get-FileWindow {
    param (
        [parameter(Mandatory = $True)]
        [string]$FileName,
        [parameter(Mandatory = $false)]
        [string]$GameName = $null,
        [parameter(Mandatory = $false)]
        [string]$Title = $null,
        [parameter(Mandatory = $false)]
        [string]$Filter = $null,
        [parameter(Mandatory = $false)]
        [string]$Message = $null
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $FileType = Get-FileType -FileName $FileName
    If ([string]::IsNullOrEmpty($Title)) {
        If ([string]::IsNullOrEmpty($GameName)) {
            $Title = "Where is $FileName file?"
        } Else {
            $Title = "Where is $($GameName.ToUpper()) $FileName file?"
        }
    }
    If ([string]::IsNullOrEmpty($Filter)) {
        If ([string]::IsNullOrEmpty($GameName)) {
            $Filter = "$FileType ($FileName) | $FileName"
        } Else {
            $Filter = "$($GameName.ToUpper()) $FileType ($FileName) | $FileName"
        }
    }
    If ([string]::IsNullOrEmpty($Message)) {
        $Message = "BMAT failed to find {0} file on its own. Let BMAT know where it is." -f ($FileName)
    }
    Write-HostAndLog -Message $Message -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true

    # Create the TopMost owner window
    $TopWindow = New-Object System.Windows.Forms.Form
    $TopWindow.TopMost = $true

    $FileBrowserDialog = New-Object System.Windows.Forms.OpenFileDialog
    $FileBrowserDialog.Title = $Title
    $FileBrowserDialog.Filter = $Filter
    $FileBrowserDialog.InitialDirectory = 'C:\'
    $FileBrowserDialog.Multiselect = $false
    $FileBrowserDialog.ShowHelp = $false

    $FilePath = $null
    Do {
        # Pass $TopWindow so the dialog stays in front
        $Result = $FileBrowserDialog.ShowDialog($TopWindow)

        if ($Result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            [System.Windows.Forms.MessageBox]::Show($TopWindow, "No File Selected. Please select $FileName file!", "File Selection Error", "OK", [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
            $FilePath = $null
        } ElseIf ($FileBrowserDialog.FileName -notmatch [regex]::Escape($FileName)) {
            [System.Windows.Forms.MessageBox]::Show($TopWindow, "Not Correct File Selected. Please select $FileName file!", "File Selection Error", "OK", [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
            $FilePath = $null
        } Else { 
            $FilePath = $FileBrowserDialog.FileName
        }
    } While ($null -eq $FilePath) # Loop until a valid path is set
    $FileBrowserDialog.Dispose()
    $TopWindow.Dispose()
    Return $FilePath
}

function Get-FolderWindow {
    param (
        [parameter(Mandatory = $true)]
        [string]$FolderName,
        [parameter(Mandatory = $false)]
        [boolean]$ShowNewFolderButton = $false,
        [parameter(Mandatory = $false)]
        [string]$GameName = $null,
        [parameter(Mandatory = $false)]
        [string]$Title = $null,
        [parameter(Mandatory = $false)]
        [string]$Message = $null
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    If ([string]::IsNullOrEmpty($Title)) {
        If ([string]::IsNullOrEmpty($GameName)) {
            $Title = "Where is `"..\$FolderName`" folder?"
        } Else {
            $Title = "Where is $($GameName.ToUpper()) `"..\$FolderName`" folder?"
        }
    }
    If ([string]::IsNullOrEmpty($Message)) {
        $Message = "BMAT failed to find `"..\{0}`" folder on its own. Let BMAT know where it is." -f ($FolderName)
    }

    Write-HostAndLog -Message $Message -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true

    # Create the TopMost owner window
    $TopWindow = New-Object System.Windows.Forms.Form
    $TopWindow.TopMost = $true

    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.Description = $Title
    $FolderBrowserDialog.ShowNewFolderButton = $ShowNewFolderButton

    $FolderPath = $null
    Do {
        $Result = $FolderBrowserDialog.ShowDialog($TopWindow)
        if ($Result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            [System.Windows.Forms.MessageBox]::Show($This, "No Folder Selected. Please select `"..\$FolderName`" folder!", "Folder Selection Error", "OK", [System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $FolderPath = $null
        } ElseIf ($FolderBrowserDialog.SelectedPath -notmatch [regex]::Escape($FolderName)) {
            [System.Windows.Forms.MessageBox]::Show($This, "Not Correct Folder Selected. Please select `"..\$FolderName`" folder!", "Folder Selection Error", "OK", [System.Windows.Forms.MessageBoxIcon]::Exclamation) | Out-Null
            $FolderPath = $null
        } Else { 
            $FolderPath = $FolderBrowserDialog.SelectedPath
        }
    } While ($null -eq $FolderPath) # Loop until a valid path is set
    $FolderPath = $FolderBrowserDialog.SelectedPath
    $FolderBrowserDialog.Dispose()
    $TopWindow.Dispose()
    Return $FolderPath
}

function Set-FolderWindow {
    param (
        [parameter(Mandatory = $true)]
        [string]$Title,
        [parameter(Mandatory = $false)]
        [boolean]$ShowNewFolderButton = $true
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
    
    # Create the TopMost owner window
    $TopWindow = New-Object System.Windows.Forms.Form
    $TopWindow.TopMost = $true

    $FolderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowserDialog.Description = $Title
    $FolderBrowserDialog.ShowNewFolderButton = $ShowNewFolderButton

    $FolderPath = $null
    Do {
        $Result = $FolderBrowserDialog.ShowDialog($TopWindow)
        if ($Result -eq [System.Windows.Forms.DialogResult]::Cancel) {
            [System.Windows.Forms.MessageBox]::Show($This, "No Folder Selected. Please select a folder!", "Folder Selection Error", "OK", [System.Windows.Forms.MessageBoxIcon]::Exclamation)
            $FolderPath = $null
        } Else { 
            $FolderPath = $FolderBrowserDialog.SelectedPath
        }
    } While ($null -eq $FolderPath) # Loop until a valid path is set

    $FolderPath = $FolderBrowserDialog.SelectedPath
    $FolderBrowserDialog.Dispose()
    $TopWindow.Dispose()
    Return $FolderPath
}

function Set-Folder {
    param (
        [parameter(Mandatory = $True)]
        [string]$FolderName,
        [parameter(Mandatory = $True)]
        [string]$ParentFolder
    )
    $FolderPath = Join-Path -Path $ParentFolder -ChildPath $FolderName
    If (-not (Test-Path -Path $FolderPath)) {
        Write-Host ("Creating folder `"{0}`"" -f ($FolderPath)) -ForegroundColor Green
        Try {
            New-Item -Path $FolderPath -ItemType Directory -ErrorAction Stop
        } Catch {
            Write-Host ("Failure with step. Error message:{0}" -f $($_.Exception.Message)) -ForegroundColor DarkYellow
            Return $false
        }
    }
    Return $true
}

function Get-DualDecisionWindow {
    Param(
        [parameter(Mandatory = $true)]
        [string]$LButtonLabel,
        [parameter(Mandatory = $true)]
        [string]$RButtonLabel,
        [parameter(Mandatory = $false)]
        [string]$Title,
        [parameter(Mandatory = $true)]
        [string]$Message,
        [parameter(Mandatory = $false)]
        [string]$BMATFolder,
        [parameter(Mandatory = $false)]
        [ValidateSet("Information", "Warning", "Error", "Question")]
        [ValidateNotNullOrEmpty()]
        [string]$Category
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Create the Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.AutoSize = $true
    $form.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.TopMost = $true # Highly recommended for "Decision" windows

    # Set form icon
    $GUIIcon = $null
    Try {
        $GUIIcon = Get-BMATIcon -Type BMATCC
    } Catch {
        $GUIIcon = $null
    }
    if (![string]::IsNullOrEmpty($Category)) {
        switch ($Category) {
            "Information" { $form.Icon = [System.Drawing.SystemIcons]::Information }
            "Warning" { $form.Icon = [System.Drawing.SystemIcons]::Warning }
            "Error" { $form.Icon = [System.Drawing.SystemIcons]::Error }
            "Question" { $form.Icon = [System.Drawing.SystemIcons]::Question }
        }
    } elseif ($null -ne $GUIIcon) {
        $form.Icon = $GUIIcon
    } else {
        $form.Icon = [System.Drawing.SystemIcons]::Information
    }

    # Setup the TableLayoutPanel
    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.AutoSize = $true
    $table.ColumnCount = 1
    $table.RowCount = 2
    $table.Padding = New-Object System.Windows.Forms.Padding(30)

    # Force the column to span the full width of the table
    $columnStyle = New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)
    [void]$table.ColumnStyles.Add($columnStyle)

    # Message Label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $label.AutoSize = $true
    # Using 'None' to keep it centered in the table cell
    $label.Anchor = [System.Windows.Forms.AnchorStyles]::None 
    $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 20)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter # Ensures text looks good if it wraps

    # Creating button Container
    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.AutoSize = $true
    $buttonPanel.Anchor = [System.Windows.Forms.AnchorStyles]::None
    $buttonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $buttonPanel.WrapContents = $false

    # Creating the buttons
    $btn1 = New-Object System.Windows.Forms.Button
    $btn1.Text = $LButtonLabel
    $btn1.Width = 100
    $btn1.Height = 30
    $btn1.DialogResult = [System.Windows.Forms.DialogResult]::Yes

    $btn2 = New-Object System.Windows.Forms.Button
    $btn2.Text = $RButtonLabel
    $btn2.Width = 100
    $btn2.Height = 30
    $btn2.DialogResult = [System.Windows.Forms.DialogResult]::No

    # Window assembly
    $buttonPanel.Controls.Add($btn1)
    $buttonPanel.Controls.Add($btn2)
    $table.Controls.Add($label, 0, 0)
    $table.Controls.Add($buttonPanel, 0, 1)
    $form.Controls.Add($table)

    # Ensure minimum width
    $form.MinimumSize = New-Object System.Drawing.Size(400, 0)

    $result = $form.ShowDialog()
    Return $result
}

function Get-MessageWindow {
    Param(
        [parameter(Mandatory = $false)]
        [string]$Title,
        [parameter(Mandatory = $true)]
        [string]$Message,
        [parameter(Mandatory = $false)]
        [ValidateSet("Information", "Warning", "Error", "Question")]
        [ValidateNotNullOrEmpty()]
        [string]$Category
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Create the Form
    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.AutoSize = $true
    $form.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.MaximizeBox = $false
    $form.TopMost = $true

    # Set form icon
    $GUIIcon = $null
    Try {
        $GUIIcon = Get-BMATIcon -Type BMATCC
    } Catch {
        $GUIIcon = $null
    }
    if (![string]::IsNullOrEmpty($Category)) {
        switch ($Category) {
            "Information" { $form.Icon = [System.Drawing.SystemIcons]::Information }
            "Warning" { $form.Icon = [System.Drawing.SystemIcons]::Warning }
            "Error" { $form.Icon = [System.Drawing.SystemIcons]::Error }
            "Question" { $form.Icon = [System.Drawing.SystemIcons]::Question }
        }
    } elseif ($null -ne $GUIIcon) {
        $form.Icon = $GUIIcon
    } else {
        $form.Icon = [System.Drawing.SystemIcons]::Information
    }

    # Setup the TableLayoutPanel
    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.AutoSize = $true
    $table.ColumnCount = 1
    $table.RowCount = 2
    $table.Padding = New-Object System.Windows.Forms.Padding(30)

    # Force the column to span the full width of the table
    $columnStyle = New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)
    [void]$table.ColumnStyles.Add($columnStyle)

    # Message Label
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $label.AutoSize = $true
    # Using 'None' to keep it centered in the table cell
    $label.Anchor = [System.Windows.Forms.AnchorStyles]::None 
    $label.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 20)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter # Ensures text looks good if it wraps

    # Creating button Container
    $buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
    $buttonPanel.AutoSize = $true
    $buttonPanel.Anchor = [System.Windows.Forms.AnchorStyles]::None
    $buttonPanel.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $buttonPanel.WrapContents = $false

    # Creating the button
    $btn1 = New-Object System.Windows.Forms.Button
    $btn1.Text = 'OK'
    $btn1.Width = 100
    $btn1.Height = 30
    $btn1.DialogResult = [System.Windows.Forms.DialogResult]::OK

    # Window assembly
    $buttonPanel.Controls.AddRange($btn1)
    $table.Controls.Add($label, 0, 0)
    $table.Controls.Add($buttonPanel, 0, 1)
    $form.Controls.Add($table)

    # Ensure minimum width
    $form.MinimumSize = New-Object System.Drawing.Size(400, 0)

    $result = $form.ShowDialog()
    Return $result
}

Function Get-GameFolderSteam {
    Param(
        [parameter(Mandatory = $True)]
        [string]$SteamAppID
    )
    $SteamAppPath = $null
    $SteamRegistryKeys = 'HKLM:\SOFTWARE\WOW6432Node\Valve\Steam', 'HKLM:\SOFTWARE\Valve\Steam'
    foreach ($Key in $SteamRegistryKeys) {
        $InstallDir = (Get-ItemProperty -LiteralPath $Key -ErrorAction SilentlyContinue).InstallPath
        if ($InstallDir) {
            $FullExePath = Join-Path $InstallDir "Steam.exe"
            if (Test-Path -Path $FullExePath) {
                $SteamAppPath = $FullExePath
                break
            }
        }
    }
    If ($SteamAppPath) {
        $SteamLibraryVdfPath = Join-Path -Path (Split-Path -LiteralPath $SteamAppPath) -ChildPath '\steamapps\libraryfolders.vdf'
        $LibraryVDFContent = Get-Content -Path $SteamLibraryVdfPath -ErrorAction SilentlyContinue
    }
    If ($LibraryVDFContent) {
        $SteamFoldersJson = Convert-FromVDFToJSON -Content $LibraryVDFContent
        foreach ($Folder in $SteamFoldersJson.libraryfolders.psobject.Properties) {
            $CurrentPath = $Folder.Value.path
            $Apps = $Folder.Value.apps
            if ($Apps.psobject.Properties.Name -contains $SteamAppID) {
                $ManifestPath = Join-Path $CurrentPath "steamapps\appmanifest_$SteamAppID.acf"
                $ManifestContent = Get-Content $ManifestPath -Raw
                $GameSubfolder = ([regex]::Match($manifestContent, '"installdir"\s+"([^"]+)"')).Groups[1].Value
                $SteamFolder = Join-Path -Path 'steamapps\common' -ChildPath $GameSubfolder
                $GamePath = Join-Path -Path $CurrentPath -ChildPath $SteamFolder
                if (Test-Path -LiteralPath $GamePath) {
                    Return $GamePath
                }
            }
        }
    }
}

Function Get-GameFolderGoG {
    Param(
        [parameter(Mandatory = $True)]
        [string]$GoGAppID
    )
    $GogRegistryPath = "HKLM:\SOFTWARE\WOW6432Node\GOG.com\Games"
    $GameRegistryPath = Join-Path -Path $GogRegistryPath -ChildPath $GoGAppID
    if (Test-Path $GameRegistryPath) {
        $GameInfo = Get-ItemProperty -Path $GameRegistryPath -ErrorAction SilentlyContinue
        if ($GameInfo.path) {
            $GamePath = $GameInfo.path
            if (Test-Path -LiteralPath $GamePath) {
                Return $GamePath
            }
        }
    }
}

Function Get-GameFolder {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $true)]
        [string]$ConfigFileName
    )
    If ($GameName -eq 'fallout4') {
        $GameFile = 'Fallout4.exe'
        $SteamAppID = '377160'
        $GoGAppID = '1992028391'
        $OldGenThreshold = [version]"1.10.163.0"
        $AEThreshold = [version]"1.11.0"
    }
    $GameFolder = $null
    $ExePath = $null
    $GameFileFound = $false
    if (!([string]::IsNullOrEmpty($BMATFolder))) {
        $ExePath = Join-Path -Path $BMATFolder -ChildPath $GameFile
        if (!([string]::IsNullOrEmpty($ExePath)) -and (Test-Path -Path $ExePath)) {
            $GameFileFound = $true
            $GameFolder = Split-Path -Parent $ExePath
            break
        }
        $ParentPath = Split-Path -Path $BMATFolder -Parent
        if ($ParentPath -eq $BMATFolder) { break }
        $BMATFolder = $ParentPath
    }
    if (-not $GameFileFound) {
        $GamePath = Get-GameFolderSteam -SteamAppID $SteamAppID
        if (!([string]::IsNullOrEmpty($GamePath)) -and (Test-Path -LiteralPath $GamePath)) {
            $GameFileFound = $true
            $GameFolder = $GamePath
            $ExePath = Join-Path -Path $GameFolder -ChildPath $GameFile
        }
    }
    if (-not $GameFileFound) {
        $GamePath = Get-GameFolderGoG -GoGAppID $GoGAppID
        if (!([string]::IsNullOrEmpty($GamePath)) -and (Test-Path -LiteralPath $GamePath)) {
            $GameFileFound = $true
            $GameFolder = $GamePath
            $ExePath = Join-Path -Path $GameFolder -ChildPath $GameFile
        }
    }
    if (-not $GameFileFound) {
        $ExePath = Get-FileWindow -FileName $GameFile -GameName $GameName
        $GameFolder = = Split-Path -Parent $ExePath
    }
    [version]$GameVersion = (Get-Item -Path $ExePath).VersionInfo.ProductVersion -replace '[^\d\.]'
    If ($GameVersion -gt $AEThreshold) {
        $Generation = "Anniversary Edition"
    } Elseif ($GameVersion -gt $OldGenThreshold) {
        $Generation = "Next-Gen"
    } Else {
        $Generation = "Old-Gen (Pre-April 2024)"
    }
    Write-HostAndLog -Message ("{0} game (Gen: {1}) found in folder `"{2}`"." -f ((Get-Culture).TextInfo.ToTitleCase($GameName), $Generation, $GameFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    Return $GameFolder
}

function Update-GameFolder2Config {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $true)]
        [string]$ConfigFileName
    )
    $GameFolder = Get-GameFolder -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName
    $GameFolderResult = Update-ConfigFile -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName -Property "GameFolder" -Value $GameFolder
    Get-ConfigurationVariables -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName | Out-Null
    if ($GameFolder -notmatch "\w" -and -not $GameFolderResult) {
        Return $false
    } else {
        Return $true
    }
}

function Get-ModsLoadOrder {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $true)]
        [string]$ConfigFileName
    )
    $Loadorder = Join-Path $GameName -ChildPath 'loadorder.txt'
    $LoadOrderFilePath = Join-Path -Path $env:LOCALAPPDATA -ChildPath $Loadorder
    If ([string]::IsNullOrEmpty($LoadOrderFilePath) -or -not (Test-Path -LiteralPath $LoadOrderFilePath)) {
        $LoadOrderFilePath = Get-FileWindow -FileName 'loadorder.txt' -GameName $GameName
    }
    $ConfigUpdateResult = Update-ConfigFile -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName -Property "LoadOrderFilePath" -Value $LoadOrderFilePath
    $LoadOrder = Get-Content -Path $LoadOrderFilePath | Where-Object { $_ -notmatch "\*|#" }
    $GameCorePluginFiles = @('Fallout4.esm', 'DLCRobot.esm', 'DLCworkshop01.esm', 'DLCworkshop02.esm', 'DLCworkshop03.esm', 'DLCCoast.esm', 'DLCNukaWorld.esm')
    $LoadOrder = $GameCorePluginFiles + $CCPlugings + ($LoadOrder | Where-Object { $GameCorePluginFiles -notcontains $_ -and $CCPlugings -notcontains $_ })
    If (-not $ConfigUpdateResult) {
        Return $false
    } Else {
        Return $LoadOrder
    }
}

function Get-CCModsLoadOrder {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $true)]
        [string]$ConfigFileName
    )
    $LoadOrder = Get-ModsLoadOrder -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName
    $IndexedLoadOrder = @()
    If (-not $LoadOrder) {
        Write-Host "Failure with loading game load order!" -ForegroundColor DarkYellow
        Return $false
    } Else {
        $CCLoadOrder = $LoadOrder | Where-Object { $_ -match "^cc\w{9}\-" }
        $i = 0
        ForEach ($Plugin in $LoadOrder) {
            $Properties = [ordered]@{
                "Index"   = $i;
                "Plugin"  = $Plugin;
                "Enabled" = $true;
            }
            $IndexedLoadOrder += New-Object PSObject -prop $Properties
            $i++
        }
        $IndexedCCLoadOrder = $IndexedLoadOrder | Where-Object { $CCLoadOrder -contains $_.Plugin }
    }
    Return $IndexedCCLoadOrder
}

function Get-CCModsName {
    $CatalogPath = Join-Path -Path $GameName -ChildPath 'ContentCatalog.txt'
    $ContentCatalogPath = Join-Path -Path $env:LOCALAPPDATA -ChildPath $CatalogPath
    If ((![string]::IsNullOrEmpty($ContentCatalogPath)) -and (Test-Path -LiteralPath $ContentCatalogPath)) {
        $ContentCatalog = Get-Content -Path $ContentCatalogPath -Raw | ConvertFrom-Json
        $Report = New-Object System.Collections.Generic.List[PSCustomObject]
        foreach ($Key in $ContentCatalog.PSObject.Properties.Name) {
            if ($Key -eq "ContentCatalog") { continue }
            $ModTitle = $ContentCatalog.$Key.Title
            foreach ($FileName in $ContentCatalog.$Key.Files) {
                If ($FileName -match ".\esp|\.esl|\.esm") {
                    If ([string]::IsNullOrEmpty($ModTitle)) {
                        $Report.Add([PSCustomObject]@{
                                Plugin  = $FileName
                                ModName = $FileName -replace (".\esp|\.esl|\.esm", "")
                            })
                    } else {
                        $Report.Add([PSCustomObject]@{
                                Plugin  = $FileName
                                ModName = $ModTitle
                            })
                    }
                }
            }
        }
    }
    Return $Report
}

function Get-BSArchPath {
    param (
        [parameter(Mandatory = $True)]
        [string]$BMATFolder
    )
    $BSArchexe = "BSArch.exe"
    If ([string]::IsNullOrEmpty($BSArchPath) -or -not (Test-Path -LiteralPath $BSArchPath)) {
        $BSArchPath = (Get-ChildItem -LiteralPath $BMATFolder -Recurse -Filter $BSArchexe | Sort-Object -Property LastWriteTime | Select-Object -Last 1).FullName
        If (![string]::IsNullOrEmpty($BSArchPath)) {
            $Output = & $BSArchPath
            If ([string]::IsNullOrEmpty($Output -match 'BSArch [\w]+ x64')) {
                Return $BSArchPath
            }
        }
    }
    If ([string]::IsNullOrEmpty($BSArchPath) -or -not (Test-Path -LiteralPath $BSArchPath)) {
        Do {
            $BSArchPath = Get-FileWindow -FileName $BSArchexe
            If (![string]::IsNullOrEmpty($BSArchPath)) {
                $Output = & $BSArchPath
                If (!([string]::IsNullOrEmpty($Output -match 'BSArch [\w]+ x64'))) {
                    Break
                }
            } else {
                Write-HostAndLog -Message "BSArch.exe should be x64 version. Please try again." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            }
        } while ([string]::IsNullOrEmpty($Output -match 'BSArch [\w]+ x64'))
    }
    Return $BSArchPath
}

function Update-BSArchPath2Config {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $true)]
        [string]$ConfigFileName
    )
    $BSArchPath = Get-BSArchPath -BMATFolder $BMATFolder
    if ($BSArchPath -notmatch "\w") {
        Return $false
    }
    $BSArchPathResult = Update-ConfigFile -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName -Property "BSArchPath" -Value $BSArchPath
    Get-ConfigurationVariables -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName | Out-Null
    if ($BSArchPath -notmatch "\w" -and -not $BSArchPathResult) {
        Return $false
    } else {
        Return $true
    }
}

function Set-BMATFolders {
    Param(
        [parameter(Mandatory = $True)]
        [string]$BMATFolder,
        [parameter(Mandatory = $True)]
        [string]$ConfigFileName,
        [parameter(Mandatory = $True)]
        [AllowEmptyString()]
        [string]$BA2MergingStagingFolder,
        [parameter(Mandatory = $True)]
        [AllowEmptyString()]
        [string]$BA2MergingTempFolder
    )
    $Result = @()
    $Result += Set-Folder -FolderName 'Tracker' -ParentFolder $BMATFolder
    $Global:TrackerFolder = Join-Path -Path $BMATFolder -ChildPath 'Tracker'
    $Result += Set-Folder -FolderName 'Logs' -ParentFolder $BMATFolder
    $Global:LogsFolder = Join-Path -Path $BMATFolder -ChildPath 'Logs'

    $SelectedStagingFolderPath = $null
    If ($null -eq $BA2MergingStagingFolder -or $BA2MergingStagingFolder -eq "") {
        Write-Host "Let BMAT know which folder it can use to store the original mods files for later use." -ForegroundColor Green
        Write-Host "Do not select the game folder or any of the Vortex or MO2 mod manager mods downloads or staging folders!" -ForegroundColor Yellow
        Write-Host "If using BMAT CC only version along with BMAT choose separate folder for BMAT CC!" -ForegroundColor Yellow
        Write-Host "The folder selection window might be hidden behind another window if you have clicked elsewhere during the process." -ForegroundColor Green
        $SelectedStagingFolderPath = Set-FolderWindow -Title "Select/Create BA2 merge staging folder for BMAT"
        If ($SelectedStagingFolderPath -match "\w") {
            $ConfigUpdateResult = Update-ConfigFile -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName -Property "BA2MergingStagingFolder" -Value $SelectedStagingFolderPath
            If (-not $ConfigUpdateResult) {
                $Result += $false
            }
            $Result += $true
        } else {
            $Result += $false
        }
    } else {
        $SelectedStagingFolderPath = $BA2MergingStagingFolder
    }
    $Result += Set-Folder -FolderName 'Staging' -ParentFolder $SelectedStagingFolderPath
    $Global:StagingWorkingFolder = Join-Path -Path $SelectedStagingFolderPath -ChildPath 'Staging'

    $SelectedTempFolderPath = $null
    If ($null -eq $BA2MergingTempFolder -or $BA2MergingTempFolder -eq "") {
        Write-Host "Let BMAT know which folder it can use to extract and repackage mod files. Preferably choose SSD or NVME disk." -ForegroundColor Green
        Write-Host "Do not select the game folder or any of the Vortex or MO2 mod manager mods downloads or staging folders!" -ForegroundColor Yellow
        Write-Host "The folder selection window might be hidden behind another window if you have clicked elsewhere during the process." -ForegroundColor Green
        $SelectedTempFolderPath = Set-FolderWindow -Title "Select/Create a working/temp folder BMAT"
        If ($SelectedTempFolderPath -match "\w") {
            $ConfigUpdateResult = Update-ConfigFile -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName -Property "BA2MergingTempFolder" -Value $SelectedTempFolderPath
            If (-not $ConfigUpdateResult) {
                $Result += $false
            }
            $Result += $true
        } else {
            $Result += $false
        }
    } else {
        $SelectedTempFolderPath = $BA2MergingTempFolder
    }
    $Result += Set-Folder -FolderName 'Output' -ParentFolder $SelectedTempFolderPath
    $Global:OutputWorkingFolder = Join-Path -Path $SelectedTempFolderPath -ChildPath 'Output'
    $Result += Set-Folder -FolderName 'Temp' -ParentFolder $SelectedTempFolderPath
    $Global:TemporaryWorkingFolder = Join-Path -Path $SelectedTempFolderPath -ChildPath 'Temp'
    $Result += Set-Folder -FolderName 'Temp_Main' -ParentFolder $SelectedTempFolderPath
    $Global:TemporaryMainFolder = Join-Path -Path $SelectedTempFolderPath -ChildPath 'Temp_Main'
    $Result += Set-Folder -FolderName 'Temp_Textures' -ParentFolder $SelectedTempFolderPath
    $Global:TemporaryTexturesFolder = Join-Path -Path $SelectedTempFolderPath -ChildPath 'Temp_Textures'
    $Result += Set-Folder -FolderName 'CleanUp' -ParentFolder $SelectedTempFolderPath
    $Global:CleanUpFolder = Join-Path -Path $SelectedTempFolderPath -ChildPath 'CleanUp'

    If ($Result -contains $false) {
        Return $false
    } Else {
        Return $true
    }
}

function Set-FileClosed {
    param (
        [parameter(Mandatory = $True)]
        [string]$FileNamePath,
        [parameter(Mandatory = $True)]
        [string]$FileDescription,
        [parameter(Mandatory = $True)]
        [string]$BMATFolder
    )
    If ((Test-Path -LiteralPath $FileNamePath) -and (Test-FileLock -Path (Get-Item -Path $FileNamePath).FullName)) {
        Do {
            Get-DualDecisionWindow -LButtonLabel 'Yes' -RButtonLabel 'No' -Title 'Close Opened File!' -Message ("{0} `"{1}`" is open. Please close it and confirm here when it is done." -f ($FileDescription, $FileNamePath)) -BMATFolder $BMATFolder -Category Warning | Out-Null
        } While (Test-FileLock -Path (Get-Item -Path $FileNamePath).FullName)
    }
}

function Set-BMATFilesClosed {
    param (
        [parameter(Mandatory = $True)]
        [string]$BMATFolder
    )
    $MergedFilesTrackerPath = Join-Path -Path $TrackerFolder -ChildPath $MergedFilesTrackerName
    Set-FileClosed -FileNamePath $MergedFilesTrackerPath -FileDescription 'BMAT tracker file' -BMATFolder $BMATFolder
}

function Get-BMATIcon {
    param (
        [parameter(Mandatory = $True)]
        [ValidateSet("BMAT", "BMATCC")]
        [string]$Type
    )
    switch ($Type) {
        "BMAT" {
            $IconBase64String = ''
        }
        "BMATCC" {
            $IconBase64String = 'AAABAAEAEBAAAAAAIABCAQAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAQAAAAEAgGAAAAH/P/YQAAAAFzUkdCAK7OHOkAAAAEZ0FNQQAAsY8L/GEFAAAACXBIWXMAAA7DAAAOwwHHb6hkAAAA10lEQVQ4T2NgwAOmXlvyv+IQBKPLEQVgBiS25JFvgI+p5v9ATyswDcLoavACkIbuA5PhXiDJAJjt6OEAcwlRhoEVGqgQ74LnS/X+Y8MOupr/nyzWA9v6ZKbK/yczVDANAgmia0THIANANEjt/TZFhCEggcfTVY6ha0DHE9O0IK6AGoBiCMwgYjGKRhiA2XSo0+r/mYlmGC7oiHcjzoCWGPf/L5Yh2E3R7mD2vfmGxBsAol8u0wWzYfxdLdbEGYAPk2wAzCswmmQDYM4H0U9nqf6vCnFAMQAAwopq1Oon7YcAAAAASUVORK5CYII='
        }
    }
    $IconBytes = [System.Convert]::FromBase64String($IconBase64String)
    $MemoryStream = [System.IO.MemoryStream]::new($IconBytes)
    $GuiIcon = [System.Drawing.Icon]::FromHandle(([System.Drawing.Bitmap]::FromStream($MemoryStream)).GetHicon())
    Return $GuiIcon
}

function Set-EmptyESPPlugin {
    param (
        [string]$Path,
        [string]$Name
    )
    $DestinationPath = Join-Path -Path $Path -ChildPath $Name
    $DummyEspBase64 = 'VEVTNCoAAAAAAgAAAAAAAAAAAACDAAAASEVEUgwAAACAPwAAAAAACAAAQ05BTQgAREVGQVVMVABJTlRWBAABAAAA'
    $EspBytes = [System.Convert]::FromBase64String($DummyEspBase64)
    [System.IO.File]::WriteAllBytes($DestinationPath, $EspBytes)
}

function Set-ConfigDefault {
    param (
        [string]$ConfigPath
    )
    Write-Host "Creating default configuration file..." -ForegroundColor Gray
    $DefaultConfig = @"
{
  "_comment": "The following property defines the folder where the original mods ba2 files will be moved to, to be use by BMAT for the BA2 merges and separate from your mod manager and game folder so the game doesn't try to double load them. If empty BMAT will ask once during run and will store the value here.",
  "BA2MergingStagingFolder": "",
  "_comment": "If you have a faster disk where you want to do the BA2 files extract and re-archive use this property which defines the folder where the ba2 files will be extracted, cleaned up and re-archived. If this property is empty the temporary folders will be created under 'BA2MergingStagingFolder' folder. BMAT automatically cleans up after itself to keep the disk usage low. It is recommended to keep this folder on a separate disk with sufficient disk space. If empty BMAT will ask once during run and will store the value here.",
  "BA2MergingTempFolder": "",
  "_comment": "The following property defines the path to the loadorder.txt file. If empty BMAT will ask once during run and will store the value here.",
  "LoadOrderFilePath": "",
  "_comment": "The following property defines the game folder path. If empty BMAT will ask once during run and will store the value here.",
  "GameFolder": "",
  "_comment": "The following property defines the path to xEdit's BSArch.exe or BSArch64.exe. If empty BMAT will ask once during run and will store the value here.",
  "BSArchPath": "",
  "_comment": "The following property defines the name of the new re-package mod name, the esp file name, and the ba2 file names.",
  "BA2MergedModName": "BMAT_BA2_Merge",
  "_comment": "The following property defines the which will be used by BMAT. 2 logs are always maintained by the script where one is current, and one is the previous log file.",
  "LogFileName": "bmat_ba2_merge",
  "_comment": "The following property defines the name of the csv file where BMAT will keep track of any processed mods and ba2 files. Once you start using BMAT deletion or manual modification of this file is not recommended.",
  "MergedFilesTrackerName": "bmat_merged-files-tracker.csv",
  "_comment": "List files extensions where BMAT will check if any of those are found in the mod ba2 files and report on them to flag a dirty ba2 file.",
  "NonSupportedFilesFilter": "tga,psd,ini,exe,dll,yml",
  "_comment": "(This option is experimental and should be used with caution) If this property is set to true BMAT will try and detect if in the original BA2 files there are any folders which are in the wrong place and will move them where they should be. For example, if there are any Textures folders in the original Main type BA2 file those will be moved to a new Textures BA2 merge. If there are any Main type (like Main, Interface, Materials, Meshes, Scripts, Sound, Sounds, Voices, Animations, Music, Video) folders in the original Textures BA2 file, those will be moved to to a new Main BA2 merge. If the original BA2 files contain some other folders (my_mod_folder1\\main\\my_mod_folder2\\my_mod_file) which contain Main or Textures folders, those will be relocated where they should be.",
  "ModFolderRearange": "false"
}
"@
    Try {
        $DefaultConfig | Out-File $ConfigPath -Encoding utf8
    } Catch {
        Write-Host "Failure with creating BMAT configuration file! Exiting..." -ForegroundColor DarkYellow
        Return $false
    }
    Return $true
}

function Get-RandomTip {
    param (
        [int]$ChancePercentage = 10 # 10% chance
    )
    $RandomTipsRoll = Get-Random -Minimum 1 -Maximum 101 # 1 to 100
    if ($RandomTipsRoll -le $ChancePercentage) {
        $Tips = @(
            "Enjoying BMAT? A 'Like' on YouTube helps a lot! My YouTube channel is https://www.youtube.com/@RJDevDen",
            "Enjoying my content? Subscription on YouTube helps a lot! My YouTube channel is https://www.youtube.com/@RJDevDen",
            "Don't forget to Endorse BMAT on NexusMods if it saved you time! Link to Nexus is https://www.nexusmods.com/fallout4/mods/89306",
            "Want to support development? Coffee is always appreciated at Ko-fi. My Ko-fi page is https://ko-fi.com/rjdevden",
            "Check out my YouTube channel for more guides and tips and tricks! My YouTube channel is https://www.youtube.com/@RJDevDen",
            "Want to know more about Fallout 4 game engine limits? Check out this article: https://www.nexusmods.com/fallout4/articles/6251",
            "Don't forget to install the merged mod in your mod manager and enabled its plugin, or you will have issues with your game!",
            "Always test thoroughly in game after BA2 merge!",
            "BMAT merges BA2 files not mods plugins! Do not disabled other mods plugins!",
            "Want to know more about BMAT? Check out the BMAT guides on YouTube: https://www.youtube.com/playlist?list=PLnRJa1RqXU3rI2LYPPeHMsHmbQKEk9L1D",
            "Game modding requires knowledge, patience and a lot of testing!",
            "You like my creations and want to support me? A coffee helps a lot! My Ko-fi page is https://ko-fi.com/rjdevden",
            "You like my mods and want to support me? Endorse my mods on Nexus https://www.nexusmods.com/profile/rjshadowface/mods",
            "Your game is crashing? Check if you are above the BA2 limit https://www.nexusmods.com/fallout4/articles/6251",
            "Your game is crashing? Did you forget to enable the BMAT merged mod plugins?",
            "Your game is crashing? Check the BMAT log file for errors.",
            "Want to know how many BA2 slots your game is using? Check out this article: https://www.nexusmods.com/fallout4/articles/6251",
            "Want to know which mods files have been merged? Check the BMAT tracker file.",
            "Always make sure the merged mod plugins are placed correctly in your load order! Consult the BMAT tracker file for the merged mods load order indexes.",
            "Not every mod would work when its BA2 files are merged and loaded by another plugin! Always test and don't assume it will work!",
            "Have issues with a merge? Want to revert mods files as they were? Start BMAT again and do a restore.",
            "Want to know others experience with BMAT? Check the mod page posts section https://www.nexusmods.com/fallout4/mods/89306?tab=posts",
            "Don't just throw random mods at your game and expect them to work! Plan ahead or use pre-created and tested mods collections.",
            "Don't just throw random mods at your game and expect them to work! Mods compatibility is very important.",
            "Don't just throw random mods at your game and expect them to work! Always read mods description, posts, and bugs pages on Nexus!"
        )
        $SelectedBMATTip = $Tips | Get-Random
    }
    Return $SelectedBMATTip
}

#endregion FUNCTIONS

$BMATVersion = "1.0.35"

#region SETUP
if ([System.IO.Path]::GetExtension([System.Environment]::GetCommandLineArgs()[0]) -eq '.exe') {
    # If running as an EXE, get the directory of the process
    $ProcessPath = [System.Environment]::ProcessPath
    If (![string]::IsNullOrEmpty($ProcessPath)) {
        $BMATFolder = Split-Path -Parent ([System.Environment]::ProcessPath)
    }
} elseif ($PSScriptRoot) {
    # If running as a PS1 script
    $BMATFolder = $PSScriptRoot
} else {
    # Fallback for older PS versions or manual console runs
    $BMATFolder = Get-Location
}
# Ensure the path is valid before proceeding
if ([string]::IsNullOrWhiteSpace($BMATFolder)) {
    $BMATFolder = (Get-Item -Path ".").FullName
}
# Set the Location so relative paths work
Set-Location -Path $BMATFolder
# Verify the file exists for debugging
$targetFile = Get-ChildItem -Path $BMATFolder -File | Where-Object { $_.Name -match "BMAT_CC.(exe|ps1)" }
if ($null -eq $targetFile) {
    Write-Error "Could not find BMAT_CC file in $BMATFolder"
}

$StartTime = Get-Date
$GameSupportedLooseFolders = @("Main", "Interface", "Materials", "Meshes", "MeshesExtra", "Scripts", "Sound", "Sounds", "Voices_en", "Voices", "Animations", "Music", "Video", "lodsettings", "vis", "Strings", "F4SE", "Notes", "Textures")
$TexturesFileTypeRegex = "Textures"
$TexturesLooseFolderRegex = '^' + $TexturesFileTypeRegex + '$'
$AllLooseFolderInclusiveRegex = '\\' + ($GameSupportedLooseFolders -join ('(\\)?|\\')) + '(\\)?'
$AllLossFoldersForMerge = ($AllLossFoldersForMerge -Replace ("\s", "")) -Split (',')
$NonSupportedFilesFilter = (($NonSupportedFilesFilter -Replace ("\s", "")) -Split (',') | ForEach-Object { '.' + $_ }) -Join ('|')

$Global:ConfigFileName = 'Config.json'
$ConfigPath = Join-Path -Path $BMATFolder -ChildPath $ConfigFileName
if (-not (Test-Path $ConfigPath)) {
    $DefaultConfigResult = Set-ConfigDefault -ConfigPath $ConfigPath
    If (-not $DefaultConfigResult) {
        Write-Host "Failure with creating BMAT configuration file! Exiting..." -ForegroundColor DarkYellow
        Pause
        Break
    }
    Write-Host "`nThank you for using BMAT CC (BA2 Merging Automation Tool) for Creation Club mods!" -ForegroundColor Cyan
    Write-Host "`nIt looks like this is your first time running the tool.`n"
    Write-Host "If you find this tool helpful, please consider:"
    Write-Host "`t- Endorsing my tools/mods on NexusMods: " -NoNewline
    Write-Host "https://www.nexusmods.com/profile/rjshadowface/mods" -ForegroundColor Cyan
    Write-Host "`t- Subscribing to my YouTube channel: " -NoNewline
    Write-Host "https://www.youtube.com/@RJDevDen" -ForegroundColor Cyan
    Write-Host "`t- Buying me a coffee: " -NoNewline
    Write-Host "https://ko-fi.com/rjdevden`n" -ForegroundColor Cyan
    Pause
}
Write-Host "`n--- BMAT CC: BA2 Merging Automation Tool for Creation Club mods v$BMATVersion ---" -ForegroundColor Cyan
Write-Host "`tDeveloped by RJ (RJDevDen)`n" -ForegroundColor Gray

$GetConfigResult = Get-ConfigurationVariables -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName
If (-not $GetConfigResult) {
    Write-Host "Failure with fetching BMAT configuration! Exiting..." -ForegroundColor DarkYellow
    Pause
    Break
}

$SetFoldersResult = Set-BMATFolders -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName -BA2MergingStagingFolder $BA2MergingStagingFolder -BA2MergingTempFolder $BA2MergingTempFolder
If (-not $SetFoldersResult) {
    Write-Host "Failure setting up BMAT folders! Exiting..." -ForegroundColor DarkYellow
    Pause
    Break
}

$SetLoggingResult = Set-Logging -BMATFolder $BMATFolder
If (-not $SetLoggingResult) {
    Write-Host "Failure with setting up BMAT logging! Exiting..." -ForegroundColor DarkYellow
    Pause
    Break
}

Set-BMATFilesClosed -BMATFolder $BMATFolder

$SetGameFolderResult = Update-GameFolder2Config -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName
If (-not $SetGameFolderResult) {
    Write-HostAndLog -Message "Failure with finding and storing game folder in config file! Exiting..." -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
    Pause
    Break
} Else {
    $GameDataFolder = Join-Path -Path $GameFolder -ChildPath 'Data'
}

$Global:OSDetails = Get-ComputerInfo
Write-HostAndLog -Message ("OS runtime environment: {0} {1}" -f ($OSDetails.OSName, $OSDetails.OsArchitecture)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true

$SetBSArchPathResult = Update-BSArchPath2Config -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName
If (-not $SetBSArchPathResult) {
    Write-HostAndLog -Message "Failure with finding and storing BSArch path in config file! Exiting..." -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
    Pause
    Break
}

$ConfigResult = Get-ConfigurationVariables -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName
If (-not $ConfigResult) {
    Write-Host "Failure with fetching BMAT configuration! Exiting..." -ForegroundColor DarkYellow
    Pause
    Break
}

$IndexedLoadOrder = Get-CCModsLoadOrder -BMATFolder $BMATFolder -ConfigFileName $ConfigFileName
If (-not $IndexedLoadOrder) {
    Write-HostAndLog -Message "Failure with fetching CC mods load order! Exiting..." -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
    Pause
    Break
}
Try {
    Add-Content -Path $LogFilePath -Value ("{0} - {1}: Load Order:`n{2}" -f ($(Get-Date), 'INFO', ($IndexedLoadOrder | Out-String)))
} Catch {
    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
}

$CCModsNames = Get-CCModsName

#endregion SETUP

#region RESTORE
$RA_Response = ""
If (Test-Path -Path $StagingWorkingFolder) {
    If (Get-ChildItem -LiteralPath $StagingWorkingFolder) {
        Write-HostAndLog -Message "There are mods in BMAT staging folder." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        $RA_Response = Get-DualDecisionWindow -LButtonLabel 'Analysis' -RButtonLabel 'Restore' -Title 'BMAT Analysis or Restore?' -Message "There are Creation Club mods files in the BMAT staging folder '$StagingWorkingFolder'.`n`nDo you want to proceed with the BA2 files analysis prior the merge or you want to restore all previously merged mods BA2 files to your game folder?`n`nIf you have just updated BMAT (maybe as part of mod collection update) please perform a restore before proceeding with any mods files merge or re-merge to avoid unnecessary issues!" -BMATFolder $BMATFolder
        switch ($RA_Response) {
            'Yes' {
                $RA_Response = 'a'
                Write-HostAndLog -Message "BMAT will proceed with the BA2 analysis prior the merge process." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            }
            'No' {
                $RA_Response = 'r'
                Write-HostAndLog -Message "BMAT will now restore all the mods BA2 files to the game folder and will remove its current tracker file which will reinitialise the BA2 merge process." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            }
        }
    } Else {
        Write-HostAndLog -Message "There are no mods in BMAT staging folder. BMAT will proceed with the BA2 files analysis prior the merge." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        $RA_Response = 'a'
    }
} Else {
    Write-HostAndLog -Message ("There are no mods in BMAT staging folder. BMAT will proceed with the BA2 files analysis." -f ($ModFileDetails.Mod_Source_Path)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    $RA_Response = 'a'
}

If ($RA_Response -eq "r") {
    IF ($BA2MergingTempFolder -eq "") {
        $BA2MergingTempFolder = $BA2MergingStagingFolder
    }
    $CCBA2InToolStaging = Get-ChildItem -LiteralPath $StagingWorkingFolder -Recurse -File | Where-Object { $_.Extension -eq ".ba2" -and $_.Directory.BaseName -eq 'CC_Mods' } | Select-Object -Property Name, Directory, FullName
    $ModFoldersInToolStaging = Get-ChildItem -LiteralPath $StagingWorkingFolder
    If ($CCBA2InToolStaging.Count -eq 0) {
        Write-HostAndLog -Message "No mods BA2 files were found in BMAT staging folder." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    } Else {
        If ($CCBA2InToolStaging.Count -gt 0) {
            $Counter = 1
            ForEach ($File in $CCBA2InToolStaging) {
                Write-HostAndLog -Message ("Restoring back BA2 file `"{0}`" to `"{1}`"." -f ($File.Name, $GameDataFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                Write-Progress -Id 0 -Activity "BA2 Restore..." -Status ('    {0}% complete' -f ([math]::Round((($Counter / ($CCBA2InToolStaging).Count) * 100)))) -PercentComplete (($Counter / $CCBA2InToolStaging.Count) * 100)
                $Counter++
                If (Test-Path -LiteralPath $File.FullName.Replace($File.Directory.FullName, $GameDataFolder)) {
                    Write-HostAndLog -Message ("BA2 file `"{0}`" is already in the game 'Data' folder. Removing BA2 file from BMAT staging." -f ($File.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    Try {
                        $ProgressPreference = 'SilentlyContinue'
                        Remove-Item -LiteralPath $File.FullName -ErrorAction Stop
                        $ProgressPreference = 'Continue'
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                } Else {
                    Try {
                        Move-Item -LiteralPath $File.FullName -Destination $GameDataFolder -ErrorAction Stop
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
            }
        }
        ForEach ($Folder in $ModFoldersInToolStaging) {
            IF ((Get-ChildItem -LiteralPath $Folder.FullName | Measure-Object).Count -eq 0) {
                Try {
                    $ProgressPreference = 'SilentlyContinue'
                    Remove-Item -LiteralPath $Folder.FullName -ErrorAction Stop
                    $ProgressPreference = 'Continue'
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
            } Else {
                Write-HostAndLog -Message ("Something went wrong. Folder `"{0}`" is not empty. Move the files manually to `"{1}`" if required." -f ($Folder.FullName, $Folder.FullName.Replace($StagingWorkingFolder, $GameDataFolder))) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
            }
        }
    }
    $MergedFilesTrackerPath = Join-Path -Path $TrackerFolder -ChildPath $MergedFilesTrackerName
    If (Test-Path -LiteralPath $MergedFilesTrackerPath) {
        Write-HostAndLog -Message ("Removing tracker file `"{0}`"." -f ($MergedFilesTrackerPath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            $ProgressPreference = 'SilentlyContinue'
            Remove-Item -LiteralPath $MergedFilesTrackerPath -ErrorAction Stop
            $ProgressPreference = 'Continue'
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
    } Else {
        Write-HostAndLog -Message "Tracker file was already removed." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }
    If (Test-Path $OutputWorkingFolder) {
        Write-HostAndLog -Message ("Removing BMAT output folder `"{0}`"." -f ($OutputWorkingFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            $ProgressPreference = 'SilentlyContinue'
            Remove-Item -LiteralPath $OutputWorkingFolder -Recurse -ErrorAction Stop
            $ProgressPreference = 'Continue'
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
    } Else {
        Write-HostAndLog -Message "BMAT output folder was already removed." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }
    If (Test-Path $TemporaryWorkingFolder) {
        Write-HostAndLog -Message ("Removing BMAT temp folder `"{0}`"." -f ($TemporaryWorkingFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            $ProgressPreference = 'SilentlyContinue'
            Remove-Item -LiteralPath $TemporaryWorkingFolder -Recurse -ErrorAction Stop
            $ProgressPreference = 'Continue'
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
    } Else {
        Write-HostAndLog -Message "BMAT temp folder was already removed." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }
    If (Test-Path $TemporaryMainFolder) {
        Write-HostAndLog -Message ("Removing BMAT temp main folder `"{0}`"." -f ($TemporaryMainFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            $ProgressPreference = 'SilentlyContinue'
            Remove-Item -LiteralPath $TemporaryMainFolder -Recurse -ErrorAction Stop
            $ProgressPreference = 'Continue'
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
    } Else {
        Write-HostAndLog -Message "BMAT temp main folder was already removed." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }
    If (Test-Path $TemporaryTexturesFolder) {
        Write-HostAndLog -Message ("Removing BMAT temp textures folder `"{0}`"." -f ($TemporaryTexturesFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            $ProgressPreference = 'SilentlyContinue'
            Remove-Item -LiteralPath $TemporaryTexturesFolder -Recurse -ErrorAction Stop
            $ProgressPreference = 'Continue'
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
    } Else {
        Write-HostAndLog -Message "BMAT temp textures folder was already removed." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }
    If (Test-Path $CleanUpFolder) {
        Write-HostAndLog -Message ("Removing BMAT cleanup folder `"{0}`"." -f ($CleanUpFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            $ProgressPreference = 'SilentlyContinue'
            Remove-Item -LiteralPath $CleanUpFolder -Recurse -ErrorAction Stop
            $ProgressPreference = 'Continue'
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
    } Else {
        Write-HostAndLog -Message "BMAT cleanup folder was already removed." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }
    If (Test-Path $StagingWorkingFolder) {
        Write-HostAndLog -Message ("Removing BMAT staging sub-folder `"{0}`"." -f ($StagingWorkingFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            $ProgressPreference = 'SilentlyContinue'
            Remove-Item -LiteralPath $StagingWorkingFolder -Recurse -ErrorAction Stop
            $ProgressPreference = 'Continue'
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
    } Else {
        Write-HostAndLog -Message "BMAT staging sub-folder was already removed." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }
    Get-MessageWindow -Title "Remove BMAT merged mod!" -Message "To complete the restore process you need to manually remove the created by BMAT merge mod '$BA2MergedModName' from yor mod manager!" -Category Warning | Out-Null

    $Tip = Get-RandomTip -ChancePercentage 50
    If ($Tip -match "\w") {
        Write-Host "`n[TIP] $Tip`n" -ForegroundColor Yellow
    }

    Write-HostAndLog -Message "BMAT restoration finished. Closing BMAT..." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    Write-Host "`nProcess Complete. Thank you for using BMAT!" -ForegroundColor Cyan
    Pause
    Break
}
#endregion RESTORE

#region ANALYSIS
If ($RA_Response -eq "a") {
    Write-HostAndLog -Message "Starting Creation Club mods files analysis" -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    
    # Creating candidates list with CC mods
    $FullModList = @()
    $CCPlugings = $IndexedLoadOrder.Plugin
    $Counter = 1
    ForEach ($Plugin in $CCPlugings) {
        $ModName = ($Plugin -Split ('\.'))[0]
        Write-Progress -Id 0 -Activity "Checking CC Mods Details..." -Status ('    {0}% complete' -f ([math]::Round((($Counter / ($CCPlugings).Count) * 100)))) -PercentComplete (($Counter / ($CCPlugings).Count) * 100)
        $Counter++
        Try {
            $BA2FilesList = Get-ChildItem -LiteralPath $GameDataFolder -ErrorAction Stop | Where-Object { $_.Extension -eq '.ba2' -and $_.BaseName -match ([regex]::Escape($ModName) + " - (main|textures)") }
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
        $FileDetails = @()
        ForEach ($File in $BA2FilesList) {
            $LoadOrderPosition = ($IndexedLoadOrder | Where-Object { $_.Plugin -eq $Plugin }).Index
            IF ($File.Name -match "Main") {
                $Type = "Main"
            } ELSEIF ($File.Name -match "Textures") {
                $Type = "Textures"
            } Else {
                $Type = "UNKNOWN"
            }

            $FileBytes = Get-Content -LiteralPath $File.FullName -Encoding Byte -TotalCount 12
            $ArchiveType = [System.Text.Encoding]::UTF8.GetString($FileBytes[-4..-1])
            If ($ArchiveType -eq 'GNRL') {
                $DiscoveredType = 'Main'
                If ($Type -ne $DiscoveredType) {
                    Write-HostAndLog -Message ("BA2 file `"{0}`" part of mod `"{1}`" was named as `"{2}`" type, but it was packaged as `"{3}`"" -f ($File.Name, $ModName, $Type, $DiscoveredType)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
                }
            } elseif ($ArchiveType -eq 'DX10') {
                $DiscoveredType = 'Textures'
                If ($Type -ne $DiscoveredType) {
                    Write-HostAndLog -Message ("BA2 file `"{0}`" part of mod `"{1}`" was named as `"{2}`" type, but it was packaged as `"{3}`"" -f ($File.Name, $ModName, $Type, $DiscoveredType)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
                }
            } else {
                $DiscoveredType = 'UNKNOWN'
                Write-HostAndLog -Message ("BA2 file `"{0}`" part of mod `"{1}`" was named as `"{2}`" type, but its packaging type is `"UNKNOWN`"" -f ($File.Name, $ModName, $Type)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
            }

            $Properties = [ordered]@{
                "Name"                = $File.Name;
                "Folder"              = $File.Directory.Name;
                "Folder_Path"         = $File.Directory.FullName;
                "Path"                = $File.FullName;
                "Modified"            = [datetime]($File.LastWriteTime);
                "Size_Bytes"          = $File.Length;
                "Type"                = If ($Type -notmatch $DiscoveredType) {
                    $DiscoveredType 
                } Else {
                    $Type 
                };
                "Included_In_BA2"     = $null;
                "Plugin_Name"         = $Plugin;
                "Load_Order_Position" = $LoadOrderPosition;
                "Plugin_Mod"          = $Plugin;
                "Mod_Name"            = $ModName;
            }
            $FileDetails += New-Object PSObject -prop $Properties
        }
        $Properties = [ordered]@{
            "Mod_Name"              = $ModName;
            "Mod_Id"                = $null;
            "isPrimary"             = $null;
            "fileType"              = $null;
            "Mod_Display_Name"      = ($CCModsNames | Where-Object { $_.Plugin -eq $Plugin }).ModName;
            "Mod_Secondary_Name"    = $null;
            "Enabled"               = $true;
            "Game"                  = $GameName;
            "State"                 = $null;
            "Mod_Size_Bytes"        = ($FileDetails.Size_Bytes + (Get-Item -LiteralPath (Join-Path -Path $GameDataFolder -ChildPath $Plugin)).Length | Measure-Object -Sum).Sum;
            "Version"               = $null;
            "Mod_Version"           = $null;
            "Mod_File_Name"         = $null;
            "Mod_Source_Path"       = $GameDataFolder;
            "Mod_Target_Path"       = $null;
            "Collections"           = $null;
            "BA2_Files_Details"     = $FileDetails;
            "Plugins_List"          = $Plugin;
            "PreCombines"           = $false;
            "Consistency"           = "Clean";
            "Not_Supported_Files"   = $null;
            "Repackage_Group_Name"  = $null;
            "Loose_Folders_Details" = $null;
            "CC_Mod"                = $true;
        }
        $FullModList += New-Object PSObject -prop $Properties
    }

    # Checking if any mods are in the full list
    If (($FullModList | Measure-Object).Count -lt 1) {
        Write-HostAndLog -Message "No merge candidates identified! Exiting..." -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        Pause
        Break
    }

    # Importing the tracker file and updating the list of mods with their location in the repackage staging folder
    $MergedFilesTrackerPath = Join-Path -Path $TrackerFolder -ChildPath $MergedFilesTrackerName
    If (Test-Path -LiteralPath $MergedFilesTrackerPath) {
        $BA2RepackageTracker = Import-Csv -Path $MergedFilesTrackerPath
        $DateUpdatedTracker = @()
        $StagedFiles = @()
        ForEach ($Line in $BA2RepackageTracker) {
            $Line.Mod_BA2_Modified = [datetime]::parseexact($Line.Mod_BA2_Modified, 'dd-MM-yyyy-HH-mm-ss', $null)
            $DateUpdatedTracker += $Line
            $FileStagingPath = Join-Path -Path $Line.Mod_Target_Path -ChildPath $Line.Mod_BA2_File
            If ((!([string]::IsNullOrEmpty($FileStagingPath))) -and (Test-Path -LiteralPath $FileStagingPath)) {
                $StagedFiles += $Line.Mod_BA2_File
            }
        }
        $BA2RepackageTracker = $DateUpdatedTracker
        If ($StagedFiles.Count -eq 0) {
            $BA2RepackageTracker = $null
            Write-HostAndLog -Message ("Found BMAT tracker file, but the associated files in BMAT staging are missing. This tracker file will be removed as it looks like it is a leftover from a clean up.") -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
            Write-HostAndLog -Message ("Removing tracker file `"{0}`"." -f ($MergedFilesTrackerPath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            Try {
                $ProgressPreference = 'SilentlyContinue'
                Remove-Item -LiteralPath $MergedFilesTrackerPath -ErrorAction Stop
                $ProgressPreference = 'Continue'
            } Catch {
                Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            }
        }
    }
    IF ($BA2RepackageTracker) {
        Write-HostAndLog -Message "Importing the tracker file and updating the list of mods with their location in the repackage staging folder" -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        Try {
            Add-Content -Path $LogFilePath -Value ("{0} - {1}: Input tracker details are:`n{2}" -f ($(Get-Date), 'INFO', ($BA2RepackageTracker | ConvertTo-Json | Out-String)))
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
        $Counter = 1
        ForEach ($File in $BA2RepackageTracker) {
            Write-Progress -Id 0 -Activity "Updating Mod Files Repackage Location Details..." -Status ('    {0}% complete' -f ([math]::Round((($Counter / ($BA2RepackageTracker).Count) * 100)))) -PercentComplete (($Counter / ($BA2RepackageTracker).Count) * 100)
            $Counter++
            IF ($null -ne ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name })) {
                If ($null -ne ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name -and (($_.BA2_Files_Details -ne $null -and $File.Mod_BA2_File -ne $null) -and $_.BA2_Files_Details.Name -contains $File.Mod_BA2_File) })) {
                    $ModObj = (($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name -and $_.BA2_Files_Details.Name -eq $File.Mod_BA2_File }).BA2_Files_Details | Where-Object { $_.Name -eq $File.Mod_BA2_File })
                    $ModObj.Path = (Join-Path -Path $File.Mod_Target_Path -ChildPath $File.Mod_BA2_File)
                    $ModObj.Included_In_BA2 = $File.Included_In_BA2
                    $ModObj.Folder = $File.Mod_Folder
                    $ModObj.Folder_Path = $File.Mod_Folder_Path
                } ElseIf ($null -ne ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name -and (($_.BA2_Files_Details.Name -eq $null -and $File.Mod_BA2_File -ne $null -and $File.Mod_BA2_File -ne '')) })) {
                    $Properties = [ordered]@{
                        "Name"                = $File.Mod_BA2_File;
                        "Folder"              = $File.Mod_Folder;
                        "Folder_Path"         = $File.Mod_Folder_Path;
                        "Path"                = (Join-Path -Path $File.Mod_Target_Path -ChildPath $File.Mod_BA2_File);
                        "Modified"            = [datetime]($File.Mod_BA2_Modified);
                        "Size_Bytes"          = $File.Mod_BA2_Size_Bytes;
                        "Type"                = $File.Mod_BA2_Type;
                        "Included_In_BA2"     = $File.Included_In_BA2;
                        "Plugin_Name"         = $File.Plugin_Name;
                        "Load_Order_Position" = $File.LoadOrder;
                        "Plugin_Mod"          = $File.Plugin_Mod;
                    }
                    ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).BA2_Files_Details += New-Object PSObject -prop $Properties
                } ElseIf ($null -ne ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name -and (($_.BA2_Files_Details.Name -ne $null -and $File.Mod_BA2_File -ne $null) -and $_.BA2_Files_Details.Name -notcontains $File.Mod_BA2_File) })) {
                    $Properties = [ordered]@{
                        "Name"                = $File.Mod_BA2_File;
                        "Folder"              = $File.Mod_Folder;
                        "Folder_Path"         = $File.Mod_Folder_Path;
                        "Path"                = (Join-Path -Path $File.Mod_Target_Path -ChildPath $File.Mod_BA2_File);
                        "Modified"            = [datetime]($File.Mod_BA2_Modified);
                        "Size_Bytes"          = $File.Mod_BA2_Size_Bytes;
                        "Type"                = $File.Mod_BA2_Type;
                        "Included_In_BA2"     = $File.Included_In_BA2;
                        "Plugin_Name"         = $File.Plugin_Name;
                        "Load_Order_Position" = $File.LoadOrder;
                        "Plugin_Mod"          = $File.Plugin_Mod;
                    }
                    ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).BA2_Files_Details += New-Object PSObject -prop $Properties
                } Else {
                    Write-HostAndLog -Message ("Unhandled match and update from tracker file.`n`nFullList match:`n{0}`nTracker match:`n{1}" -f ((($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }) | ConvertTo-Json | Out-String), ($File | ConvertTo-Json | Out-String))) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
            }
        }
    }
    Try {
        Add-Content -Path $LogFilePath -Value ("{0} - {1}: Full Mods Details (if tracker is available the ba2 files details will be updated with the details from it):`n{2}" -f ($(Get-Date), 'INFO', (ConvertTo-Json ($FullModList | Where-Object { $_.Game -eq $GameName }))))
    } Catch {
        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
    }

    # Creating change tracker and checking if any of the mods and the ba2 files were added, removed, or changed which would require BA2 repackaging.
    Write-HostAndLog -Message "Creating change tracker and checking if any mods files were changed which would require BA2 repackaging." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    $BA2FilesCandidatesList = $FullModList
    If ($BA2RepackageTracker) {
        # Creating change tracker
        $ModsChangeTracker = @()
        $Counter = 1   
        ForEach ($FileInTracker in $BA2RepackageTracker) {
            Write-Progress -Id 0 -Activity 'Existing Mod Check...' -Status ('    {0}% complete' -f ([math]::Round((($Counter / $BA2RepackageTracker.Count) * 100)))) -PercentComplete (($Counter / $BA2RepackageTracker.Count) * 100)
            $Counter++

            Write-HostAndLog -Message ("Checking file `"{0}`" part of mod `"{1}`" for changes which would require BA2 repackaging." -f ($FileInTracker.Mod_BA2_File, $FileInTracker.Mod_Display_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true

            $NewMod = $null
            $BARepackage = $false
            $ModDetail = $BA2FilesCandidatesList | Where-Object { $_.Mod_Name -eq $FileInTracker.Mod_Name -and $_.BA2_Files_Details.Name -eq $FileInTracker.Mod_BA2_File }
            $BA2FileDetail = $ModDetail.BA2_Files_Details | Where-Object { $_.Name -eq $FileInTracker.Mod_BA2_File }

            If ($FullModList.Mod_Name -notcontains $FileInTracker.Mod_Name) {
                Write-HostAndLog -Message ("File `"{0}`" part of mod `"{1}`" was removed from the game, so the BA2 will be re-packaged again and the mod files will be removed from the repackaging folder." -f ($FileInTracker.Mod_BA2_File, $FileInTracker.Mod_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $Status = "Removed_from_Game"
                $BARepackage = $true
            } ElseIf (($FullModList | Where-Object { $_.Mod_Name -eq $FileInTracker.Mod_Name }) -and ($FullModList | Where-Object { $_.Mod_Name -eq $FileInTracker.Mod_Name }).Count -gt 1) {
                If (($null -ne $BA2FileDetail -and $BA2FileDetail.Size_Bytes -eq $FileInTracker.Mod_BA2_Size_Bytes) -and ($null -ne $BA2FileDetail -and $BA2FileDetail.Modified.DateTime -eq $FileInTracker.Mod_BA2_Modified.DateTime)) {
                    Write-HostAndLog -Message ("File `"{0}`" part of mod `"{1}`" is already in the tracker. No changes detected, so no BA2 re-package is required for this mod." -f ($FileInTracker.Mod_BA2_File, $FileInTracker.Mod_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    $Status = "Identical_to_Tracker"
                    $BARepackage = $false
                } Else {
                    Write-HostAndLog -Message ("File `"{0}`" part of mod `"{1}`" is already in the tracker, but it was updated, so the BA2 will be re-packaged again." -f ($FileInTracker.Mod_BA2_File, $FileInTracker.Mod_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    $Status = "Updated_Changed"
                    $BARepackage = $true
                }
            } ElseIF (($null -ne $BA2FileDetail -and $BA2FileDetail.Size_Bytes -ne $FileInTracker.Mod_BA2_Size_Bytes) -or ($null -ne $BA2FileDetail -and $BA2FileDetail.Modified.DateTime -ne $FileInTracker.Mod_BA2_Modified.DateTime)) {
                Write-HostAndLog -Message ("File `"{0}`" part of mod `"{1}`" is already in the tracker, but it was updated, so the BA2 will be re-packaged again." -f ($FileInTracker.Mod_BA2_File, $FileInTracker.Mod_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $Status = "Updated_Changed"
                $BARepackage = $true
            } ElseIf (($null -ne $BA2FileDetail -and $BA2FileDetail.Size_Bytes -eq $FileInTracker.Mod_BA2_Size_Bytes) -and ($null -ne $BA2FileDetail -and $BA2FileDetail.Modified.DateTime -eq $FileInTracker.Mod_BA2_Modified.DateTime)) {
                Write-HostAndLog -Message ("File `"{0}`" part of mod `"{1}`" is already in the tracker. No changes detected, so no BA2 re-package is required for this mod." -f ($FileInTracker.Mod_BA2_File, $FileInTracker.Mod_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $Status = "Identical_to_Tracker"
                $BARepackage = $false
            } Else {
                Write-HostAndLog -Message ("File `"{0}`" part of mod `"{1}`" in the tracker, but the file change status is currently not handled by BMAT. Collect logs and send to BMAT author. FileInTracker Output:`n{2}`nMod Details:`n{3}" -f ($FileInTracker.Mod_BA2_File, $FileInTracker.Mod_Name, ($FileInTracker | Out-String), (($FullModList | Where-Object { ($_.Mod_Name -eq $FileInTracker.Mod_Name) -and ($_.Mod_Name -eq $FileInTracker.Mod_Name) }) | ConvertTo-Json))) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                $Status = "Unhandled_Status"
                $BARepackage = $false
            }
            $ModBA2File = $FileInTracker.Mod_BA2_File
            $ModLooseFolder = $null
            $LOIndex = $FileInTracker.LoadOrder
            $PluginName = $FileInTracker."Plugin_Name"
            $PluginMod = $FileInTracker.Plugin_Mod
            If ($Status -eq "Updated_Changed") {
                $OldMod = $FileInTracker
                $NewMod = $ModDetail
                If ($null -eq $LOIndex) {
                    $LOIndex = ($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $FileInTracker.Mod_BA2_File }).Load_Order_Position
                }
                If ($null -eq $NewMod.Mod_Target_Path -or $NewMod.Mod_Target_Path -eq "") {
                    If (($FullModList | Where-Object { $_.Mod_Name -eq $FileInTracker.Mod_Name }).CC_Mod -eq $true) {
                        $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                    } else {
                        $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath $NewMod.Mod_Name
                    }
                    $NewMod.Mod_Target_Path = $ModTargetPath
                }
            }
            $Properties = [ordered]@{
                "Mod_Display_Name" = $FileInTracker.Mod_Display_Name;
                "Mod_Name"         = $FileInTracker.Mod_Name;
                "Mod_BA2_File"     = $ModBA2File;
                "Mod_Loose_Folder" = $ModLooseFolder;
                "Mod_Folder"       = $FileInTracker.Mod_Folder;
                "Mod_Folder_Path"  = $FileInTracker.Mod_Folder_Path;
                "Mod_Source_Path"  = $FileInTracker.Mod_Source_Path;
                "Mod_Target_Path"  = $FileInTracker.Mod_Target_Path;
                "Status"           = $Status;
                "Updated_By"       = $NewMod.Mod_Name;
                "In_BA2"           = $FileInTracker.Included_In_BA2;
                "Group"            = $FileInTracker.Repackaging_Group;
                "BA2_Re_Merge"     = $BARepackage;
                "Plugin_Name"      = $PluginName;
                "LO_Index"         = $LOIndex;
                "Plugin_Mod"       = $PluginMod;
                "Modified"         = $FileInTracker.Mod_BA2_Modified;
                "Size"             = $FileInTracker.Mod_BA2_Size_Bytes;
                "Type"             = $FileInTracker.Mod_BA2_Type;
            }
            $ModsChangeTracker += New-Object PSObject -prop $Properties
        }

        # Checking for newly added mods
        $NewModsList = $BA2FilesCandidatesList | Where-Object { $BA2RepackageTracker.Mod_Name -notcontains $_.Mod_Name -and $_.Enabled -eq $true }
        If ($null -ne $NewModsList) {
            $Counter = 1
            ForEach ($Mod in $NewModsList) {
                Write-HostAndLog -Message ("Checking newly processed mod `"{0}`"." -f ($Mod.Mod_Display_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                Write-Progress -Id 0 -Activity 'New Mod Check...' -Status ('    {0}% complete' -f ([math]::Round((($Counter / $NewModsList.Count) * 100)))) -PercentComplete (($Counter / $NewModsList.Count) * 100)
                $Counter++
                If ($Mod.CC_Mod -eq $true) {
                    $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                } else {
                    $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath $Mod.Mod_Name
                }
                $Mod.Mod_Target_Path = $ModTargetPath
                ForEach ($File in $Mod.BA2_Files_Details) {
                    Write-HostAndLog -Message ("Checking file `"{0}`"." -f ($File.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    IF ($File.Type -eq "Main" -and $Mod.CC_Mod -eq $True) {
                        $MergeGroup = $BA2MergedModName + "_UserAdded_CC"
                        $BA2Package = $MergeGroup + " - Main.ba2"
                        $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                        $Mod.Mod_Target_Path = $ModTargetPath
                    } ElseIf ($File.Type -eq "Textures" -and $Mod.CC_Mod -eq $True) {
                        $MergeGroup = $BA2MergedModName + "_UserAdded_CC"
                        $BA2Package = $MergeGroup + " - Textures.ba2"
                        $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                        $Mod.Mod_Target_Path = $ModTargetPath
                    }
                    $Status = "New_Candidate"
                    $Properties = [ordered]@{
                        "Mod_Display_Name" = ($CCModsNames | Where-Object { $_.Plugin -eq $File."Plugin_Name" }).ModName;
                        "Mod_Name"         = $Mod.Mod_Name;
                        "Mod_BA2_File"     = $File.Name;
                        "Mod_Loose_Folder" = $null;
                        "Mod_Folder"       = $File.Folder;
                        "Mod_Folder_Path"  = $File.Folder_Path;
                        "Mod_Source_Path"  = $Mod.Mod_Source_Path;
                        "Mod_Target_Path"  = $Mod.Mod_Target_Path;
                        "Status"           = $Status;
                        "Updated_By"       = $null;
                        "In_BA2"           = $BA2Package;
                        "Group"            = $MergeGroup;
                        "BA2_Re_Merge"     = $true;
                        "Plugin_Name"      = $File."Plugin_Name";
                        "LO_Index"         = $File.Load_Order_Position;
                        "Plugin_Mod"       = $File.Plugin_Mod;
                        "Modified"         = $File.Modified;
                        "Size"             = $File.Size_Bytes;
                        "Type"             = $File.Type;
                    }
                    $ModsChangeTracker += New-Object PSObject -prop $Properties
                }
            }
        }

        # Checking for removed BA2 files as part of mod updates
        $ToBeRemovedObj = ($ModsChangeTracker | Group-Object -Property Mod_Name | Where-Object { $_.Group.Status -contains "Updated_Changed" -and $_.Group.Updated_By -eq $null }).Group | Where-Object { $_.Updated_By -eq $null }
        If ($null -ne $ToBeRemovedObj) {
            Write-HostAndLog -Message "Checking for any removed files from mods after updates" -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            $Counter = 1
            ForEach ($Obj in $ToBeRemovedObj) {
                Write-Progress -Id 0 -Activity 'Removed Files from Mods Check...' -Status ('    {0}% complete' -f ([math]::Round((($Counter / $ToBeRemovedObj.Count) * 100)))) -PercentComplete (($Counter / $ToBeRemovedObj.Count) * 100)
                $Counter++
                Write-HostAndLog -Message ("Mod `"{0}`" was updated but file `"{1}`" does not exist in the new mod version, so it will be removed from BMAT staging." -f ($Obj.Mod_Name, $Obj.Mod_BA2_File )) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $Obj.Status = "Removed_from_Mod"
                If ($ModsChangeTracker | Where-Object { $_.Mod_Name -eq $Obj.Mod_Name -and $_.LO_Index -ne $null }) {
                    $Obj.LO_Index = ($ModsChangeTracker | Where-Object { $_.Mod_Name -eq $Obj.Mod_Name -and $_.LO_Index -ne $null })[0].LO_Index
                }
                If ($ModsChangeTracker | Where-Object { $_.Mod_Name -eq $Obj.Mod_Name -and $_.Updated_By -ne $null }) {
                    $Obj.Updated_By = ($ModsChangeTracker | Where-Object { $_.Mod_Name -eq $Obj.Mod_Name -and $_.Updated_By -ne $null })[0].Updated_By
                }
            }
        }

        # Checking for new BA2 files as part of mod updates
        $UpdateByModsList = ($BA2FilesCandidatesList | Where-Object { ($ModsChangeTracker | Where-Object { $_.Updated_By -ne $null }).Updated_By -contains $_.Mod_Name -and $_.Enabled -eq $true })
        If ($null -ne $UpdateByModsList) {
            $Counter = 1
            ForEach ($Mod in $UpdateByModsList) {
                Write-Progress -Id 0 -Activity 'Added Files to Mods Check...' -Status ('    {0}% complete' -f ([math]::Round((($Counter / $UpdateByModsList.Count) * 100)))) -PercentComplete (($Counter / $UpdateByModsList.Count) * 100)
                $Counter++
                If ($Mod.CC_Mod -eq $true) {
                    $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                } else {
                    $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath $Mod.Mod_Name
                }
                $Mod.Mod_Target_Path = $ModTargetPath
                If ($null -ne $Mod.BA2_Files_Details -and $Mod.BA2_Files_Details -ne "") {
                    ForEach ($File in $Mod.BA2_Files_Details) {
                        If (($ModsChangeTracker | Where-Object { $_.Updated_By -eq $Mod.Mod_Name }).Mod_BA2_File -notcontains $File.Name) {
                            Write-HostAndLog -Message ("BA2 file `"{0}`" had been found part of mod `"{1}`" which was not processed by BMAT before. Adding it now." -f ($File.Name, $Mod.Mod_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                            IF ($File.Type -eq "Main" -and $Mod.CC_Mod -eq $True) {
                                $MergeGroup = $BA2MergedModName + "_UserAdded_CC"
                                $BA2Package = $MergeGroup + " - Main.ba2"
                                $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                                $Mod.Mod_Target_Path = $ModTargetPath
                            } ElseIf ($File.Type -eq "Textures" -and $Mod.CC_Mod -eq $True) {
                                $MergeGroup = $BA2MergedModName + "_UserAdded_CC"
                                $BA2Package = $MergeGroup + " - Textures.ba2"
                                $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                                $Mod.Mod_Target_Path = $ModTargetPath
                            }
                            $Status = "New_Candidate"
                            $Properties = [ordered]@{
                                "Mod_Display_Name" = ($CCModsNames | Where-Object { $_.Plugin -eq $File."Plugin_Name" }).ModName;
                                "Mod_Name"         = $Mod.Mod_Name;
                                "Mod_BA2_File"     = $File.Name;
                                "Mod_Loose_Folder" = $null;
                                "Mod_Folder"       = $File.Folder;
                                "Mod_Folder_Path"  = $File.Folder_Path;
                                "Mod_Source_Path"  = $Mod.Mod_Source_Path;
                                "Mod_Target_Path"  = $Mod.Mod_Target_Path;
                                "Status"           = $Status;
                                "Updated_By"       = $null;
                                "In_BA2"           = $BA2Package;
                                "Group"            = $MergeGroup;
                                "BA2_Re_Merge"     = $true;
                                "Plugin_Name"      = $File."Plugin_Name";
                                "LO_Index"         = $File.Load_Order_Position;
                                "Plugin_Mod"       = $File.Plugin_Mod;
                                "Modified"         = $File.Modified;
                                "Size"             = $File.Size_Bytes;
                                "Type"             = $File.Type;
                            }
                            $ModsChangeTracker += New-Object PSObject -prop $Properties
                        }
                    }
                }
            }
        }

        # Flagging a BA2 merge for repackage even if one of its BA2 files had changed or was removed, or disabled
        $Counter = 1
        Write-HostAndLog -Message ("Checking which BA2 merges need repackage or if any BA2 merges will be retired." -f ($Mod.Mod_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        ForEach ($BA2File in ($ModsChangeTracker.In_BA2 | Sort-Object -Unique)) {
            Write-Progress -Id 0 -Activity 'BA2 Merge Check...' -Status ('    {0}% complete' -f ([math]::Round((($Counter / ($ModsChangeTracker.In_BA2 | Sort-Object -Unique).Count) * 100)))) -PercentComplete (($Counter / ($ModsChangeTracker.In_BA2 | Sort-Object -Unique).Count) * 100)
            $Counter++
            If (($ModsChangeTracker | Where-Object { $_.In_BA2 -eq $BA2File }).BA2_Re_Merge -contains $true) {
                ForEach ($File in ($ModsChangeTracker | Where-Object { $_.In_BA2 -eq $BA2File })) {
                    If ((($ModsChangeTracker | Where-Object { $_.Mod_BA2_File -eq $File.Mod_BA2_File -and $_.Mod_Name -eq $File.Mod_Name }) | Get-Member).Name -contains "BA2_Re_Merge") {
                        ($ModsChangeTracker | Where-Object { $_.Mod_BA2_File -eq $File.Mod_BA2_File -and $_.Mod_Name -eq $File.Mod_Name })."BA2_Re_Merge" = $True
                    } Else {
                        ($ModsChangeTracker | Where-Object { $_.Mod_BA2_File -eq $File.Mod_BA2_File -and $_.Mod_Name -eq $File.Mod_Name }) | Add-Member -NotePropertyName "BA2_Re_Merge" -NotePropertyValue $True
                    }
                }
            }
        }
    } Else {
        $ModsChangeTracker = @()
        $Counter = 1
        ForEach ($Mod in ($BA2FilesCandidatesList | Where-Object { $_.Enabled -eq $true })) {
            Write-HostAndLog -Message ("Checking newly processed mod `"{0}`"." -f ($Mod.Mod_Display_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            Write-Progress -Id 0 -Activity 'New Mod Check...' -Status ('    {0}% complete' -f ([math]::Round((($Counter / ($BA2FilesCandidatesList | Where-Object { $_.Enabled -eq $true }).Count) * 100)))) -PercentComplete (($Counter / ($BA2FilesCandidatesList | Where-Object { $_.Enabled -eq $true }).Count) * 100)
            $Counter++
            If ($Mod.CC_Mod -eq $true) {
                $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
            } else {
                $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath $Mod.Mod_Name
            }
            $Mod.Mod_Target_Path = $ModTargetPath
            ForEach ($File in $Mod.BA2_Files_Details) {
                Write-HostAndLog -Message ("Checking file `"{0}`"." -f ($File.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                IF ($File.Type -eq "Main") {
                    If ($Mod.CC_Mod -eq $True) {
                        $MergeGroup = $BA2MergedModName + "_UserAdded_CC"
                        $BA2Package = $MergeGroup + " - Main.ba2"
                        $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                        $Mod.Mod_Target_Path = $ModTargetPath
                    } Else {
                        $MergeGroup = $BA2MergedModName + "_UserAdded"
                        $BA2Package = $MergeGroup + " - Main.ba2"
                    }
                } ElseIf ($File.Type -eq "Textures") {
                    If ($Mod.CC_Mod -eq $True) {
                        $MergeGroup = $BA2MergedModName + "_UserAdded_CC"
                        $BA2Package = $MergeGroup + " - Textures.ba2"
                        $ModTargetPath = Join-Path -Path $StagingWorkingFolder -ChildPath 'CC_Mods'
                        $Mod.Mod_Target_Path = $ModTargetPath
                    } Else {
                        $MergeGroup = $BA2MergedModName + "_UserAdded"
                        $BA2Package = $MergeGroup + " - Textures.ba2"
                    }
                }
                $Status = "New_Candidate"
                $Properties = [ordered]@{
                    "Mod_Display_Name" = ($CCModsNames | Where-Object { $_.Plugin -eq $File."Plugin_Name" }).ModName;
                    "Mod_Name"         = $Mod.Mod_Name;
                    "Mod_BA2_File"     = $File.Name;
                    "Mod_Loose_Folder" = $null;
                    "Mod_Folder"       = $File.Folder;
                    "Mod_Folder_Path"  = $File.Folder_Path;
                    "Mod_Source_Path"  = $Mod.Mod_Source_Path;
                    "Mod_Target_Path"  = $Mod.Mod_Target_Path;
                    "Status"           = $Status;
                    "Updated_By"       = $null;
                    "In_BA2"           = $BA2Package;
                    "Group"            = $MergeGroup;
                    "BA2_Re_Merge"     = $true;
                    "Plugin_Name"      = $File."Plugin_Name";
                    "LO_Index"         = $File.Load_Order_Position;
                    "Plugin_Mod"       = $File.Plugin_Mod;
                    "Modified"         = $File.Modified;
                    "Size"             = $File.Size_Bytes;
                    "Type"             = $File.Type;
                }
                $ModsChangeTracker += New-Object PSObject -prop $Properties
            }
        }
    }

    # Identifying any ba2 merges which would be re-packaged
    $MembersCounter = @()
    ForEach ($MergeFile in ($ModsChangeTracker | Group-Object -Property In_BA2)) {
        $BA2Files = @()
        ForEach ($File in $MergeFile.Group) {
            $Properties = [ordered]@{
                "BA2_File"         = $File.Mod_BA2_File;
                "Mod"              = $File.Mod_Display_Name;
                "Load_Order_Index" = $File.LO_Index;
            }
            $BA2Files += New-Object PSObject -prop $Properties
        }
        $BA2Files = $BA2Files | Sort-Object -Property LO_Index
        $Properties = [ordered]@{
            "BA2_Name"                    = $MergeFile.Name;
            "Qt_Total_Members"            = $MergeFile.Count;
            "Qt_Changed"                  = ($MergeFile.Group | Where-Object { $_.BA2_Re_Merge -eq $true }).Count;
            "Qt_Status_For_Merge"         = ($MergeFile.Group | Where-Object { $_.BA2_Re_Merge -eq $true -and $_.Status -match "Updated_Changed|New_Candidate" }).Count;
            "Qt_Status_Removed_from_Game" = ($MergeFile.Group | Where-Object { $_.BA2_Re_Merge -eq $true -and $_.Status -match "Removed_from_Game|Removed_from_Mod" }).Count;
        }
        $MembersCounter += New-Object PSObject -prop $Properties
    }

    Write-HostAndLog -Message "Data gathering and analysis stage of the process has completed" -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    Write-Host "`nMaximise this console window before progressing!" -ForegroundColor Cyan
    Pause
    Write-HostAndLog -Message ("The following list of potential mods ba2 files candidates were identified by BMAT:`n{0}" -f ($ModsChangeTracker | Format-Table -Property @{e = 'Mod_Display_Name'; width = 28 }, @{e = 'Mod_Name'; width = 28 }, @{e = 'Mod_BA2_File'; width = 23 }, @{e = 'Status'; width = 20 }, @{e = 'Updated_By'; width = 28 }, @{e = 'In_BA2'; width = 23 }, @{e = 'BA2_Re_Merge'; width = 5 }, @{e = 'Plugin_Name'; width = 23 }, @{e = 'LO_Index'; width = 5 } -Wrap | Out-String)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
} 
#endregion ANALYSIS

#region MERGE
if (-not $BA2FilesCandidatesList) {
    Write-HostAndLog -Message "Based on the current BMAT configuration and, and based on the mods data analysis there are no BA2 candidates found which can be merged. Exiting..." -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
    Get-MessageWindow -Title "No BA2 candidates!" -Message "Based on the current BMAT configuration and, and based on the mods data analysis there are no BA2 candidates found which can be merged. Exiting..." -Category Warning | Out-Null
    Pause
    Break
} elseif (($MembersCounter.Qt_Changed + $MembersCounter.Qt_Status_For_Merge + $MembersCounter.Qt_Status_Removed_from_Game | Measure-Object -Sum).Sum -eq 0) {
    Write-HostAndLog -Message "No BA2 candidates found which need merging or re-merging. Exiting..." -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
    Get-MessageWindow -Title "No BA2 candidates!" -Message "No BA2 candidates found which need merging or re-merging. Exiting..." -Category Warning | Out-Null
    Pause
    Break
} else {
    $ME_Response = Get-DualDecisionWindow -LButtonLabel 'Merge' -RButtonLabel 'Exit' -Title 'BMAT Merge or Exit?' -Message "BMAT had finished with the Creation Club mods files analysis. Check if you are happy with the planned actions.`nAny mods files marked with 'True' under 'BA2_Re_Merge' will be merged or re-merged into a BA2 merge file.`n`nDo you want to proceed with the BA2 files merge or you want to exit BMAT?" -BMATFolder $BMATFolder
    switch ($ME_Response) {
        'Yes' {
            $ME_Response = 'm'
            Write-HostAndLog -Message "BMAT will proceed with the BA2 merge or re-merge process." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        }
        'No' {
            Write-HostAndLog -Message "You chose to exit BMAT. Exiting..." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            Break
        }
    }
}

If ($ME_Response -eq "m") {
    Write-HostAndLog -Message "Starting Creation Club mods files merge." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true

    # Checking disk available space for merge to proceed
    $NewModsSpaceRq = (($FullModList | Where-Object { (($ModsChangeTracker | Where-Object { $_.Status -eq "New_Candidate" }).Mod_Name | Sort-Object -Unique) -contains $_.Mod_Name }).Mod_Size_Bytes | Measure-Object -Sum).Sum
    $OldChangedModsSpaceRq = (($BA2RepackageTracker | Where-Object { (($ModsChangeTracker | Where-Object { $_.Status -eq "Updated_Changed" }).Mod_Name | Sort-Object -Unique) -contains $_.Mod_Name }).Mod_BA2_Size_Bytes | Measure-Object -Sum).Sum
    $NewChangedModsSpaceRq = (($FullModList | Where-Object { (($ModsChangeTracker | Where-Object { $_.Status -eq "Updated_Changed" }).Mod_Name | Sort-Object -Unique) -contains $_.Mod_Name }).Mod_Size_Bytes | Measure-Object -Sum).Sum
    If ((Split-Path $GameDataFolder -Qualifier) -eq (Split-Path $BA2MergingStagingFolder -Qualifier) -and (Split-Path $GameDataFolder -Qualifier) -eq (Split-Path $BA2MergingTempFolder -Qualifier)) {
        If ($NewChangedModsSpaceRq -gt $OldChangedModsSpaceRq) {
            $Staging_ModsMergeSpaceRq = ($NewModsSpaceRq + $NewChangedModsSpaceRq) * 2 * 1.3
        } Else {
            $Staging_ModsMergeSpaceRq = ($NewModsSpaceRq + $OldChangedModsSpaceRq) * 2 * 1.3
        }
        $StagingDiskFreeSpace = (Get-PSDrive (Split-Path $BA2MergingStagingFolder -Qualifier).Replace(':', '')).Free
        If ($StagingDiskFreeSpace -gt $Staging_ModsMergeSpaceRq) {
            Write-HostAndLog -Message ("BMAT checked and there should be sufficient space on disk `"{0}`" to proceed with the current merge." -f (Split-Path $BA2MergingStagingFolder -Qualifier)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        } Else {
            Write-HostAndLog -Message ("BMAT checked and there is no sufficient space on disk `"{0}`" to proceed with the current merge. Additional {1} GB of free space is required." -f ((Split-Path $BA2MergingStagingFolder -Qualifier), [math]::Round((($Staging_ModsMergeSpaceRq - $StagingDiskFreeSpace) / [Math]::Pow(1024, 3)), 2))) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            Pause
            Break
        }
    } ElseIf ((Split-Path $GameDataFolder -Qualifier) -eq (Split-Path $BA2MergingStagingFolder -Qualifier) -and (Split-Path $GameDataFolder -Qualifier) -ne (Split-Path $BA2MergingTempFolder -Qualifier)) {
        If ($NewChangedModsSpaceRq -gt $OldChangedModsSpaceRq) {
            $Temp_ModsMergeSpaceRq = ($NewModsSpaceRq + $NewChangedModsSpaceRq) * 1 * 1.3
        } Else {
            $Temp_ModsMergeSpaceRq = ($NewModsSpaceRq + $OldChangedModsSpaceRq) * 1 * 1.3
        }
        $TempDiskFreeSpace = (Get-PSDrive (Split-Path $BA2MergingTempFolder -Qualifier).Replace(':', '')).Free
        If ($TempDiskFreeSpace -gt $Temp_ModsMergeSpaceRq) {
            Write-HostAndLog -Message ("BMAT checked and there should be sufficient space on disk `"{0}`" to proceed with the current merge." -f (Split-Path $BA2MergingTempFolder -Qualifier)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        } Else {
            Write-HostAndLog -Message ("BMAT checked and there is no sufficient space on disk `"{0}`" to proceed with the current merge. Additional {1} GB of free space is required." -f ((Split-Path $BA2MergingTempFolder -Qualifier), [math]::Round((($Temp_ModsMergeSpaceRq - $TempDiskFreeSpace) / [Math]::Pow(1024, 3)), 2))) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            Pause
            Break
        }
    } ElseIf ((Split-Path $GameDataFolder -Qualifier) -ne (Split-Path $BA2MergingStagingFolder -Qualifier) -and (Split-Path $BA2MergingStagingFolder -Qualifier) -eq (Split-Path $BA2MergingTempFolder -Qualifier)) {
        If ($NewChangedModsSpaceRq -gt $OldChangedModsSpaceRq) {
            $Both_ModsMergeSpaceRq = ($NewModsSpaceRq + $NewChangedModsSpaceRq) * 3 * 1.3
        } Else {
            $Both_ModsMergeSpaceRq = ($NewModsSpaceRq + $OldChangedModsSpaceRq) * 3 * 1.3
        }
        $StagingDiskFreeSpace = (Get-PSDrive (Split-Path $BA2MergingStagingFolder -Qualifier).Replace(':', '')).Free
        If ($StagingDiskFreeSpace -gt $Both_ModsMergeSpaceRq) {
            Write-HostAndLog -Message ("BMAT checked and there should be sufficient space on disk `"{0}`" to proceed with the current merge." -f (Split-Path $BA2MergingStagingFolder -Qualifier)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        } Else {
            Write-HostAndLog -Message ("BMAT checked and there is no sufficient space on disk `"{0}`" to proceed with the current merge. Additional {1} GB of free space is required." -f ((Split-Path $BA2MergingStagingFolder -Qualifier), [math]::Round((($Both_ModsMergeSpaceRq - $StagingDiskFreeSpace) / [Math]::Pow(1024, 3)), 2))) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            Pause
            Break
        }
    } ElseIf ((Split-Path $GameDataFolder -Qualifier) -ne (Split-Path $BA2MergingStagingFolder -Qualifier) -and (Split-Path $BA2MergingStagingFolder -Qualifier) -ne (Split-Path $BA2MergingTempFolder -Qualifier)) {
        If ($NewChangedModsSpaceRq -gt $OldChangedModsSpaceRq) {
            $Staging_ModsMergeSpaceRq = ($NewModsSpaceRq + $NewChangedModsSpaceRq) * 2 * 1.3
        } Else {
            $Staging_ModsMergeSpaceRq = ($NewModsSpaceRq + $OldChangedModsSpaceRq) * 2 * 1.3
        }
        $StagingDiskFreeSpace = (Get-PSDrive (Split-Path $BA2MergingStagingFolder -Qualifier).Replace(':', '')).Free
        If ($StagingDiskFreeSpace -gt $Staging_ModsMergeSpaceRq) {
            Write-HostAndLog -Message ("BMAT checked and there should be sufficient space on disk `"{0}`" to proceed with the current merge." -f (Split-Path $BA2MergingStagingFolder -Qualifier)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        } Else {
            Write-HostAndLog -Message ("BMAT checked and there is no sufficient space on disk `"{0}`" to proceed with the current merge. Additional {1} GB of free space is required." -f ((Split-Path $BA2MergingStagingFolder -Qualifier), [math]::Round((($Staging_ModsMergeSpaceRq - $StagingDiskFreeSpace) / [Math]::Pow(1024, 3)), 2))) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            Pause
            Break
        }
        If ($NewChangedModsSpaceRq -gt $OldChangedModsSpaceRq) {
            $Temp_ModsMergeSpaceRq = ($NewModsSpaceRq + $NewChangedModsSpaceRq) * 1 * 1.3
        } Else {
            $Temp_ModsMergeSpaceRq = ($NewModsSpaceRq + $OldChangedModsSpaceRq) * 1 * 1.3
        }
        $TempDiskFreeSpace = (Get-PSDrive (Split-Path $BA2MergingTempFolder -Qualifier).Replace(':', '')).Free
        If ($TempDiskFreeSpace -gt $Temp_ModsMergeSpaceRq) {
            Write-HostAndLog -Message ("BMAT checked and there should be sufficient space on disk `"{0}`" to proceed with the current merge." -f (Split-Path $BA2MergingTempFolder -Qualifier)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
        } Else {
            Write-HostAndLog -Message ("BMAT checked and there is no sufficient space on disk `"{0}`" to proceed with the current merge. Additional {1} GB of free space is required." -f ((Split-Path $BA2MergingTempFolder -Qualifier), [math]::Round((($Temp_ModsMergeSpaceRq - $TempDiskFreeSpace) / [Math]::Pow(1024, 3)), 2))) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            Pause
            Break
        }
    }

    # Cleaning up and mod files restoration to game folder where required
    Write-HostAndLog -Message "Initiating clean-up and repackaging stage of the process" -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    $BA2RepackageCandidates = $ModsChangeTracker | Where-Object { $_.BA2_Re_Merge -eq $true }
    If ($BA2RepackageTracker) {
        ForEach ($MergeFile in ($BA2RepackageCandidates | Group-Object -Property In_BA2)) {
            Write-HostAndLog -Message ("Starting working on `"{0}`" merge file" -f ($MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            If ($MergeFile.Group | Where-Object { $_.Status -eq "Updated_Changed" }) {
                ForEach ($ModFile in ($MergeFile.Group | Where-Object { $_.Status -eq "Updated_Changed" })) {
                    $OldMod = $BA2RepackageTracker | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name -and $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }
                    $NewMod = $FullModList | Where-Object { $_.Mod_Name -eq $ModFile.Updated_By }
                    Write-HostAndLog -Message ("Mod `"{0}`" was updated or changed. Removing mod files from BMAT staging folder. Merge file `"{1}`" will be repackaged." -f ($OldMod.Mod_Name, $MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    $OldModFilePath = Join-Path -Path $OldMod.Mod_Target_Path -ChildPath $OldMod.Mod_BA2_File
                    IF (Test-Path -LiteralPath $OldModFilePath) {
                        Try {
                            $ProgressPreference = 'SilentlyContinue'
                            Remove-Item -LiteralPath $OldModFilePath -Recurse -ErrorAction Stop
                            $ProgressPreference = 'Continue'
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    } Else {
                        Write-HostAndLog -Message ("File `"{0}`" was already removed." -f ($OldModFilePath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                    $OldModInTracker = ($ModsChangeTracker | Where-Object { $_.Mod_Name -eq $OldMod.Mod_Name -and $_.Mod_BA2_File -eq $OldMod.Mod_BA2_File })
                    $TrackerUpdateForModUpdate = @()
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Mod_Display_Name: `"{0}`"") -f ($OldModInTracker.Mod_Display_Name); "New Value" = ("Mod_Display_Name: `"{0}`"") -f ($NewMod.Mod_Display_Name); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Mod_Name: `"{0}`"") -f ($OldModInTracker.Mod_Name); "New Value" = ("Mod_Name: `"{0}`"") -f ($NewMod.Mod_Name); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Mod_Source_Path: `"{0}`"") -f ($OldModInTracker.Mod_Source_Path); "New Value" = ("Mod_Source_Path: `"{0}`"") -f ($NewMod.Mod_Source_Path); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Mod_Target_Path: `"{0}`"") -f ($OldModInTracker.Mod_Target_Path); "New Value" = ("Mod_Target_Path: `"{0}`"") -f ($NewMod.Mod_Target_Path); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = "Status: `"Updated_Changed`""; "New Value" = "Status: `"New_Candidate`""; }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Mod_Folder: `"{0}`"") -f ($OldModInTracker.Mod_Folder); "New Value" = ("Mod_Folder: `"{0}`"") -f (($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Folder); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Mod_Folder_Path: `"{0}`"") -f ($OldModInTracker.Mod_Folder_Path); "New Value" = ("Mod_Folder_Path: `"{0}`"") -f (($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Folder_Path); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Modified: `"{0}`"") -f ($OldModInTracker.Modified.ToString()); "New Value" = ("Modified: `"{0}`"") -f (($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Modified.ToString()); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Size: `"{0}`"") -f ($OldModInTracker.Size); "New Value" = ("Size: `"{0}`"") -f (($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Size_Bytes); }
                    $TrackerUpdateForModUpdate += New-Object PSObject -prop @{ "Old Value" = ("Plugin_Mod: `"{0}`"") -f ($OldModInTracker.Plugin_Mod); "New Value" = ("Plugin_Mod: `"{0}`"") -f (($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Plugin_Mod); }
                    Write-HostAndLog -Message ("Updating change tracker for file `"{0}`" with the new mod `"{1}`" details:`n{2}" -f ($ModFile.Mod_BA2_File, $NewMod.Mod_Name, ($TrackerUpdateForModUpdate | Format-Table -Wrap | Out-String))) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    $OldModInTracker.Mod_Display_Name = $NewMod.Mod_Display_Name
                    $OldModInTracker.Mod_Name = $NewMod.Mod_Name
                    $OldModInTracker.Mod_Source_Path = $NewMod.Mod_Source_Path
                    $OldModInTracker.Mod_Target_Path = $NewMod.Mod_Target_Path
                    $OldModInTracker.Status = "New_Candidate"
                    $OldModInTracker.Mod_Folder = ($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Folder
                    $OldModInTracker.Mod_Folder_Path = ($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Folder_Path
                    $OldModInTracker.Modified = ($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Modified
                    $OldModInTracker.Size = ($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Size_Bytes
                    $OldModInTracker.Plugin_Mod = ($NewMod.BA2_Files_Details | Where-Object { $_.Name -eq $ModFile.Mod_BA2_File }).Plugin_Mod
                }
            } ElseIf ($MergeFile.Group | Where-Object { $_.Status -eq "Removed_from_Game" }) {
                ForEach ($ModFile in ($MergeFile.Group | Where-Object { $_.Status -eq "Removed_from_Game" })) {
                    If ($null -eq ($FullModList | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name })) {
                        $ModFileDetails = $BA2RepackageTracker | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name -and $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }
                        Write-HostAndLog -Message ("Mod `"{0}`" was removed from game. Removing mod files from BMAT staging folder. Merge file `"{1}`" will be repackaged." -f ($ModFile.Mod_Name, $MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                        $ToBeRemovedFilePath = = Join-Path -Path $ModFile.Mod_Target_Path -ChildPath $ModFile.Mod_BA2_File
                        IF (Test-Path -LiteralPath $ToBeRemovedFilePath) {
                            Try {
                                $ProgressPreference = 'SilentlyContinue'
                                Remove-Item -LiteralPath $ToBeRemovedFilePath -Recurse -Force -Confirm:$false -ErrorAction Stop
                                $ProgressPreference = 'Continue'
                            } Catch {
                                Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                            }
                        } Else {
                            Write-HostAndLog -Message ("File `"{0}`" was already removed." -f ($ToBeRemovedFilePath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    }
                }
            } ElseIf ($MergeFile.Group | Where-Object { $_.Status -eq "Removed_from_Mod" }) {
                ForEach ($ModFile in ($MergeFile.Group | Where-Object { $_.Status -eq "Removed_from_Mod" })) {
                    Write-HostAndLog -Message ("Removing mod `"{0}`" files from mod repackage folder `"{1}`"." -f ($ModFile.Mod_Display_Name, $ModFile.Mod_Target_Path)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    $FilePath = Join-Path -Path $ModFile.Mod_Target_Path -ChildPath $ModFile.Mod_BA2_File
                    IF (Test-Path -LiteralPath $FilePath) {
                        Try {
                            $ProgressPreference = 'SilentlyContinue'
                            Remove-Item -LiteralPath $FilePath -Recurse -ErrorAction Stop
                            $ProgressPreference = 'Continue'
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    } Else {
                        Write-HostAndLog -Message ("File `"{0}`" was already removed." -f ($FilePath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
            } Else {
                Write-HostAndLog -Message ("Merge file `"{0}`" will be repackaged as one or more of the mods files in it had changed or a new mod files are to be added." -f ($MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            }
        }
    }

    # Creating file extract tracker
    $FileExtrackTracker = @()
    ForEach ($File in $ModsChangeTracker) {
        $Properties = [ordered]@{
            "Mod_Name"     = $File.Mod_Name;
            "Mod_BA2_File" = $File.Name;
            "Extracted"    = $False;
        }
        $FileExtrackTracker += New-Object PSObject -prop $Properties
    }

    # Moving mod files to staging, extracting BA2 files, cleaning BA2 files, and archive BA2 merge files
    $BA2RepackageCandidatesNonRetire = $ModsChangeTracker | Where-Object { $_.BA2_Re_Merge -eq $true -and $_.Status -ne "Removed_from_Game" -and $_.Status -ne "Removed_from_Mod" }
    IF ($BA2RepackageCandidatesNonRetire) {
        # Removing any duplicate records from the change tracker
        $NoDuplicateTracker = @()
        ForEach ($File in $BA2RepackageCandidatesNonRetire) {
            If (-not ($NoDuplicateTracker | Where-Object { $_.Mod_Name -eq $File.Mod_Name -and $_.Mod_BA2_File -eq $File.Mod_BA2_File })) {
                $NoDuplicateTracker += $File
            }
            If (-not ($NoDuplicateTracker | Where-Object { $_.Mod_Name -eq $File.Mod_Name -and $_.Mod_Loose_Folder -eq $File.Mod_Loose_Folder })) {
                $NoDuplicateTracker += $File
            }
        }
        $BA2RepackageCandidatesNonRetire = $NoDuplicateTracker

        # Moving mod files to staging, extracting BA2 files, cleaning BA2 files, and archive BA2 merge files
        $MergeGroupCandidates = ($BA2RepackageCandidatesNonRetire | Group-Object -Property "Group" | Where-Object { $_.Name -ne "" })
        $TotalDSSForCleanUp = @()
        $TotalProcessedModFilesRaw = @()
        $TotalProcessedModFilesPreArch = @()
        $L1Counter = 1
        ForEach ($MergeFile in $MergeGroupCandidates) {
            #$MergeFile = $MergeGroupCandidates[1]
            Write-HostAndLog -Message ("Starting processing group `"{0}`"." -f ($MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            Write-Progress -Id 0 -Activity "Processing BA2 Merge Groups..." -Status ('    {0}% complete' -f ([math]::Round((($L1Counter / ($MergeGroupCandidates | Measure-Object).Count) * 100)))) -PercentComplete $(($L1Counter / ($MergeGroupCandidates | Measure-Object).Count) * 100)
            $L1Counter++
            $L2Counter = 1
            ForEach ($ModFile in $MergeFile.Group) {
                #$ModFile = ($MergeFile.Group)[1]
                # Read-Host "Press Any Key To Continue"
                If ($ModFolderRearange -eq $true) {
                    Write-Progress -Id 1 -ParentId 0 -Activity "Moving, Extracting and Cleaning BA2 Files..." -Status ('    {0}% complete' -f ([math]::Round((($L2Counter / ($MergeFile.Group).Count) * 100)))) -PercentComplete $(($L2Counter / ($MergeFile.Group).Count) * 100)
                } Else {
                    Write-Progress -Id 1 -ParentId 0 -Activity "Moving and Extracting BA2 Files..." -Status ('    {0}% complete' -f ([math]::Round((($L2Counter / ($MergeFile.Group).Count) * 100)))) -PercentComplete $(($L2Counter / ($MergeFile.Group).Count) * 100)
                }

                $L2Counter++
                If ($ModFile.Status -eq "New_Candidate" -or $ModFile.Status -eq "Updated_Changed") {
                    $ModDetails = $FullModList | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name }
                    If ($null -eq $ModDetails) {
                        $ModDetails = $FullModList | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name }
                    }
                } Else {
                    $ModDetails = $BA2FilesCandidatesList | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name }
                }

                # Creating mods folders in the staging
                IF (Test-Path -LiteralPath $ModFile.Mod_Target_Path) {
                    Write-HostAndLog -Message ("Mod specific staging working folder found `"{0}`"." -f ($ModFile.Mod_Target_Path)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                } Else {
                    Write-HostAndLog -Message ("Creating mod specific staging working folder `"{0}`"." -f ($ModFile.Mod_Target_Path)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    Try {
                        New-Item -Path $ModFile.Mod_Target_Path -ItemType Directory -ErrorAction Stop
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
                Write-HostAndLog -Message ("Processing file `"{0}`" part of mod `"{1}`"." -f ($ModFile.Mod_BA2_File, $ModFile.Mod_Display_Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true

                # Moving BA2 files to staging
                # New mod to BMAT, file not in source place, but file is in target place

                $BA2FileSourcePath = Join-Path -Path $ModFile.Mod_Source_Path -ChildPath $ModFile.Mod_BA2_File
                $BA2FileTargetPath = Join-Path -Path $ModFile.Mod_Target_Path -ChildPath $ModFile.Mod_BA2_File
                if (![string]::IsNullOrWhiteSpace(($BA2RepackageTracker | Where-Object { $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }).Mod_Source_Path)) {
                    $TrackerBA2FileSourcePath = Join-Path -Path ($BA2RepackageTracker | Where-Object { $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }).Mod_Source_Path -ChildPath $ModFile.Mod_BA2_File
                } else {
                    $TrackerBA2FileSourcePath = $null
                }
                If (($ModFile.Status -eq "New_Candidate") -and -not (Test-Path -LiteralPath $BA2FileSourcePath) -and (Test-Path -LiteralPath $BA2FileTargetPath)) {
                    Write-HostAndLog -Message ("File `"{0}`" is already in the working folder." -f ($ModFile.Mod_BA2_File)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    # Mod file source and target paths are the same, and file path exists
                } ElseIf (($ModFile.Mod_Source_Path -eq $ModFile.Mod_Target_Path) -and (Test-Path -LiteralPath $BA2FileTargetPath)) {
                    Write-HostAndLog -Message ("File `"{0}`" is already in the working folder." -f ($ModFile.Mod_BA2_File)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    # If the ba2 file exist in both source and target and if their modify dates match
                } ElseIf ((((!([string]::IsNullOrEmpty($BA2FileSourcePath))) -and (Test-Path -LiteralPath $BA2FileSourcePath)) -and ((!([string]::IsNullOrEmpty($BA2FileTargetPath))) -and (Test-Path -LiteralPath $BA2FileTargetPath))) -and ($ModFile.Modified -eq ($BA2RepackageTracker | Where-Object { $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }).Mod_BA2_Modified)) {
                    Write-HostAndLog -Message ("File `"{0}`" found in the game folder as well as in the BMAT repackaging staging folder. The mod file in game folder will be removed." -f ($ModFile.Mod_BA2_File)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    Try {
                        $ProgressPreference = 'SilentlyContinue'
                        Remove-Item -LiteralPath $BA2FileSourcePath -ErrorAction Stop
                        $ProgressPreference = 'Continue'
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                    # If the file path from tracker does not exist and the file from the full list i.e. the updated/replaced file exist in source
                } ElseIf ($BA2RepackageTracker -and -not (Test-Path -LiteralPath $TrackerBA2FileSourcePath) -and (Test-Path -Path $BA2FileSourcePath)) {
                    Write-HostAndLog -Message ("Relocating file `"{0}`" to working folder." -f ($BA2FileSourcePath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    If (Test-Path -LiteralPath $BA2FileSourcePath) {
                        Try {
                            Move-Item -LiteralPath $BA2FileSourcePath -Destination $ModFile.Mod_Target_Path -Force -Confirm:$false -ErrorAction Stop
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    } Else {
                        Write-HostAndLog -Message ("File `"{0}`" doesn't exist" -f ($BA2FileSourcePath)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                    # If BA2 files names between the one in the full list and the one in the tracker list match, but their modify dates are different
                } ElseIf (($ModFile.Mod_BA2_File -eq ($BA2RepackageTracker | Where-Object { $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }).Mod_BA2_File) -and ($ModFile.Modified -ne ($BA2RepackageTracker | Where-Object { $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }).Mod_BA2_Modified)) {
                    Write-HostAndLog -Message ("Relocating file `"{0}`" to working folder." -f ($TrackerBA2FileSourcePath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    If (Test-Path -LiteralPath $TrackerBA2FileSourcePath) {
                        Try {
                            Move-Item -LiteralPath $TrackerBA2FileSourcePath -Destination $ModFile.Mod_Target_Path -Force -Confirm:$false -ErrorAction Stop
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    } Else {
                        Write-HostAndLog -Message ("File `"{0}`" doesn't exist" -f ($TrackerBA2FileSourcePath)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                    # If file not in source place, but file is in target place
                } ElseIf (-not (Test-Path -LiteralPath $BA2FileSourcePath) -and (Test-Path -LiteralPath $BA2FileTargetPath)) {
                    Write-HostAndLog -Message ("File `"{0}`" is already in the working folder." -f ($ModFile.Mod_BA2_File)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                } Else {
                    Write-HostAndLog -Message ("Relocating file `"{0}`" to working folder." -f ($BA2FileSourcePath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    If (Test-Path -LiteralPath $BA2FileSourcePath) {
                        Try {
                            Move-Item -LiteralPath $BA2FileSourcePath -Destination $ModFile.Mod_Target_Path -Force -Confirm:$false -ErrorAction Stop
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    } Else {
                        Write-HostAndLog -Message ("File `"{0}`" doesn't exist" -f ($BA2FileSourcePath)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }

                # Extracting ba2 files to work folder
                If (($FileExtrackTracker | Where-Object { $_.Mod_Name -eq $ModDetails.Mod_Name -and $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }).Extracted -eq $True) {
                    Write-HostAndLog -Message ("File `"{0}`" has already been extracted. Skipping..." -f ($BA2FileTargetPath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                } Else {
                    Write-HostAndLog -Message ("Extracting file `"{0}`"." -f ($BA2FileTargetPath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                    Try {
                        $SourceFile = Join-Path -Path $ModFile.Mod_Target_Path -ChildPath $ModFile.Mod_BA2_File
                        $Output = & $BSArchPath unpack $SourceFile $CleanUpFolder 2>&1
                        if ($LASTEXITCODE -ne 0) {
                            Throw "BSArch failed to unpack $SourceFile. Exit Code: $LASTEXITCODE"
                        }
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                    If ($Output -match "exception") {
                        Write-HostAndLog -Message "Failure with step. Error message: $($Output | Out-String)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                    If ($null -ne ($FileExtrackTracker | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name -and $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File })) {
                        ($FileExtrackTracker | Where-Object { $_.Mod_Name -eq $ModFile.Mod_Name -and $_.Mod_BA2_File -eq $ModFile.Mod_BA2_File }).Extracted = $True
                    }
                    Try {
                        Add-Content -Path $LogFilePath -Value ("{0} - {1}: Extracting command output:`n{2}" -f ($(Get-Date), 'INFO', ($Output | Out-String)))
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }

                # Tracking Mod files post extract
                $AllModsFilesPostExtract = Get-ChildItem -LiteralPath $CleanUpFolder -File -Recurse | Where-Object { $F4SECoreScripts -notcontains $_.Name -and $_.Name -notmatch "Thumbs(.+)?.db" -and $_.Name -ne "desktop.ini" -and $_.Extension -ne '.tga' -and $_.Extension -ne '.psd' }
                ForEach ($File in $AllModsFilesPostExtract) {
                    $Properties = [ordered]@{
                        "File_Name"        = $File.Name;
                        "File_Path"        = $File.FullName;
                        "Mod_Display_Name" = $ModFile.Mod_Display_Name;
                        "Mod_Name"         = $ModFile.Mod_Name;
                    }
                    $TotalProcessedModFilesRaw += New-Object PSObject -prop $Properties
                }

                # Moving BA2 files to their respective overwrite temp folder
                Try {
                    $BA2Folders = Get-ChildItem -LiteralPath $CleanUpFolder -Directory -ErrorAction Stop | Where-Object { (Get-ChildItem -LiteralPath $_.FullName -Recurse).FullName -match $AllLooseFolderInclusiveRegex }
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
                ForEach ($Folder in $BA2Folders) {
                    $BA2MainFoldersSubfolders = @()
                    $BA2TexturesFoldersSubfolders = @()
                    If ($ModFolderRearange -eq $true) {
                        ForEach ($FolderName in $GameSupportedLooseFolders) {
                            If (($Folder | Get-ChildItem -Directory -Recurse | Where-Object { $_.FullName -match "\\$FolderName\\?" }).Count -gt 0) {
                                $ExtractedFolder = Get-Item -LiteralPath ($Folder | Get-ChildItem -Recurse -Directory | Where-Object { $_.FullName -match "\\$FolderName\\?" } | ForEach-Object { $_ -Replace ("(?<=$FolderName)\\.+", "") } | Sort-Object -Unique)
                                If ($ExtractedFolder.Name -eq "Textures") {
                                    $BA2TexturesFoldersSubfolders += $ExtractedFolder
                                } Else {
                                    $BA2MainFoldersSubfolders += $ExtractedFolder
                                }
                            } ElseIf (($Folder | Where-Object { $_.FullName -match "\\$FolderName\\?" }).Count -gt 0) {
                                $ExtractedFolder = Get-Item -LiteralPath ($Folder | Where-Object { $_.FullName -match "\\$FolderName\\?" } | ForEach-Object { $_ -Replace ("(?<=$FolderName)\\.+", "") } | Sort-Object -Unique)
                                If ($ExtractedFolder.Name -eq "Textures") {
                                    $BA2TexturesFoldersSubfolders += $ExtractedFolder
                                } Else {
                                    $BA2MainFoldersSubfolders += $ExtractedFolder
                                }
                            }
                        }
                    } Else {
                        If ($ModFile.Type -eq 'Main') {
                            $BA2MainFoldersSubfolders = $Folder
                        } ElseIf ($ModFile.Type -eq 'Textures') {
                            $BA2TexturesFoldersSubfolders = $Folder
                        }
                    }
                    If ($BA2MainFoldersSubfolders.Count -gt 0) {
                        ForEach ($SubFolder in $BA2MainFoldersSubfolders) {
                            $FilesList = Get-ChildItem -LiteralPath $SubFolder.FullName -File -Recurse | Where-Object { $F4SECoreScripts -notcontains $_.Name -and $_.Name -notmatch "Thumbs(.+)?.db" -and $_.Name -ne "desktop.ini" -and $_.Extension -ne '.tga' -and $_.Extension -ne '.psd' }
                            $SourceFolder = $SubFolder.Parent.FullName
                            $DestinationFolder = $TemporaryMainFolder
                            ForEach ($File in $FilesList) {
                                $DestinationFile = ($File.FullName.ToLower()).replace(($SourceFolder.ToLower()), ($DestinationFolder.ToLower()))
                                $DestinationFolderPath = ($(If ($File.Mode -match "^d") { $File.Parent.FullName } Else { $File.DirectoryName }).ToLower()).replace(($SourceFolder.ToLower()), ($DestinationFolder.ToLower()))
                                If (Test-Path -LiteralPath $DestinationFile) {
                                    Try {
                                        $ProgressPreference = 'SilentlyContinue'
                                        Remove-Item -LiteralPath $DestinationFile -Recurse -Force -ErrorAction Stop
                                        $ProgressPreference = 'Continue'
                                    } Catch {
                                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                                    }
                                }
                                Try {
                                    $null = New-Item -Path  $DestinationFolderPath -Type Directory -Force
                                } Catch {
                                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                                }
                                Try {
                                    Copy-Item -LiteralPath $File.FullName -Destination $DestinationFolderPath -Force
                                } Catch {
                                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                                }
                            }
                        }
                    }
                    If ($BA2TexturesFoldersSubfolders.Count -gt 0) {
                        ForEach ($SubFolder in $BA2TexturesFoldersSubfolders) {
                            $FilesList = Get-ChildItem -LiteralPath $SubFolder.FullName -File -Recurse | Where-Object { $_.Extension -eq '.dds' }
                            $ExcludedTexturesFiles = Get-ChildItem -LiteralPath $SubFolder.FullName -File -Recurse | Where-Object { $_.Extension -ne '.dds' }
                            If ($null -ne $ExcludedTexturesFiles) {
                                Write-HostAndLog -Message ("The following files will be skipped from the Textures BA2 merge as non DDS files:`n{0}" -f ($ExcludedTexturesFiles.FullName | Out-String)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
                            }
                            $SourceFolder = $SubFolder.Parent.FullName
                            $DestinationFolder = $TemporaryTexturesFolder
                            ForEach ($File in $FilesList) {
                                If ($ModFolderRearange -eq $true) {
                                    $DestinationFile = ($File.FullName.ToLower()).replace(($SourceFolder.ToLower()), ($DestinationFolder.ToLower())) -replace ((".+\\textures\\"), ($DestinationFolder.ToLower() + "\textures\"))
                                    $DestinationFolderPath = ($(If ($File.Mode -match "^d") { $File.Parent.FullName } Else { $File.DirectoryName }).ToLower()).replace(($SourceFolder.ToLower()), ($DestinationFolder.ToLower())) -replace ((".+\\textures\\"), ($DestinationFolder.ToLower() + "\textures\"))
                                } Else {
                                    $DestinationFile = ($File.FullName.ToLower()).replace(($SourceFolder.ToLower()), ($DestinationFolder.ToLower()))
                                    $DestinationFolderPath = ($(If ($File.Mode -match "^d") { $File.Parent.FullName } Else { $File.DirectoryName }).ToLower()).replace(($SourceFolder.ToLower()), ($DestinationFolder.ToLower()))
                                }
                                If (Test-Path -LiteralPath $DestinationFile) {
                                    Try {
                                        $ProgressPreference = 'SilentlyContinue'
                                        Remove-Item -LiteralPath $DestinationFile -Recurse -Force -ErrorAction Stop
                                        $ProgressPreference = 'Continue'
                                    } Catch {
                                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                                    }
                                }
                                Try {
                                    $null = New-Item -Path  $DestinationFolderPath -Type Directory -Force
                                } Catch {
                                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                                }
                                Try {
                                    Copy-Item -LiteralPath $File.FullName -Destination $DestinationFolderPath -Force
                                } Catch {
                                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                                }
                            }
                        }
                    }
                }
                ForEach ($Folder in $BA2Folders) {
                    Try {
                        $ProgressPreference = 'SilentlyContinue'
                        Remove-Item -LiteralPath $Folder.FullName -Recurse -Force -ErrorAction Stop
                        $ProgressPreference = 'Continue'
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
            }

            # Creating plugin for the respective group
            $PluginFilePath = Join-Path -Path $OutputWorkingFolder -ChildPath ("{0}.esp" -f ($MergeFile.Name))
            IF (Test-Path -LiteralPath $PluginFilePath) {
                Write-HostAndLog -Message ("File `"{0}`" already exists. Skipping step." -f ($PluginFilePath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            } ELSE {
                Write-HostAndLog -Message ("Creating file `"{0}`.esp`" in output folder." -f ($MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                Try {
                    Set-EmptyESPPlugin -Path $OutputWorkingFolder -Name ("{0}.esp" -f ($MergeFile.Name))
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
            }

            # Tracking Mod files pre compress
            $MainModsFilesPreCompact = Get-ChildItem -LiteralPath $TemporaryMainFolder -File -Recurse
            ForEach ($File in $MainModsFilesPreCompact) {
                $Properties = [ordered]@{
                    "File_Name" = $File.Name;
                    "File_Path" = $File.FullName;
                }
                $TotalProcessedModFilesPreArch += New-Object PSObject -prop $Properties
            }
            $TexturesModsFilesPreCompact = Get-ChildItem -LiteralPath $TemporaryTexturesFolder -File -Recurse
            ForEach ($File in $TexturesModsFilesPreCompact) {
                $Properties = [ordered]@{
                    "File_Name" = $File.Name;
                    "File_Path" = $File.FullName;
                }
                $TotalProcessedModFilesPreArch += New-Object PSObject -prop $Properties
            }

            #Pause

            # Archiving the Textures ba2 files
            IF ((Get-ChildItem -LiteralPath $TemporaryTexturesFolder | Measure-Object).Count -gt 0) {
                Write-HostAndLog -Message "Moving the cleaned Textures ba2 type folders to temp." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $TempTexturesFolders = Get-ChildItem -LiteralPath $TemporaryTexturesFolder
                ForEach ($Folder in $TempTexturesFolders) {
                    If ($Folder.Name -match $TexturesLooseFolderRegex) {
                        Try {
                            Move-Item -LiteralPath $Folder.FullName -Destination $TemporaryWorkingFolder -Force -Confirm:$false -ErrorAction Stop
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    } Else {
                        Try {
                            Copy-Item -LiteralPath $Folder.FullName -Destination $TemporaryMainFolder -Recurse -Container -Force -ErrorAction Stop
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                        Try {
                            $ProgressPreference = 'SilentlyContinue'
                            Remove-Item -LiteralPath $Folder.FullName -Recurse -Force -ErrorAction Stop
                            $ProgressPreference = 'Continue'
                        } Catch {
                            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                        }
                    }
                }
                Write-HostAndLog -Message ("Archiving the Textures BA2 files for group `"{0}`"." -f ($MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                Try {
                    $FileName = "{0} - Textures.BA2" -f $MergeFile.Name
                    $TargetFilePath = Join-Path -Path $OutputWorkingFolder -ChildPath $FileName
                    $Output = & $BSArchPath pack $TemporaryWorkingFolder $TargetFilePath -share -fo4dds -z 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Throw "BSArch failed to pack $TargetFilePath. Exit Code: $LASTEXITCODE"
                    }
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
                Try {
                    Add-Content -Path $LogFilePath -Value ("{0} - {1}: Archiving command output:`n{2}" -f ($(Get-Date), 'INFO', ($Output | Out-String)))
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
                If ($Output -match "exception") {
                    $DDSIssueFiles = $TotalDSSForCleanUp | Where-Object { $Output -match ([regex]::Escape($($_.File_Path.Replace($CleanUpFolder, '')))) }
                    Write-HostAndLog -Message ("There was an issue with mod `"{0}`" during BA2 merge! The affected files was `"{1}`"" -f ($DDSIssueFiles.Mod_Name, $DDSIssueFiles.File_Path)) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
                Write-HostAndLog -Message "Cleaning up temp folders." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $TempWorkingFolderList = Get-ChildItem -LiteralPath $("{0}" -f ($TemporaryWorkingFolder))
                ForEach ($Item in $TempWorkingFolderList) {
                    Try {
                        $ProgressPreference = 'SilentlyContinue'
                        Remove-Item -LiteralPath $Item.FullName -Recurse -Force -Confirm:$false -ErrorAction Stop
                        $ProgressPreference = 'Continue'
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
                $TempTexturesFolderList = Get-ChildItem -LiteralPath $("{0}" -f ($TemporaryTexturesFolder)) -Recurse
                ForEach ($Item in $TempTexturesFolderList) {
                    Try {
                        $ProgressPreference = 'SilentlyContinue'
                        Remove-Item -LiteralPath $Item.FullName -Recurse -Force -Confirm:$false -ErrorAction Stop
                        $ProgressPreference = 'Continue'
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
            } Else {
                Write-HostAndLog -Message "There aren't any BA2 Textures files in staging suitable to be archived." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            }

            # Archiving the main ba2 files
            IF ((Get-ChildItem -LiteralPath $TemporaryMainFolder | Measure-Object).Count -gt 0) {
                Write-HostAndLog -Message "Moving the cleaned Main ba2 type folders to temp." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $TempMainFolders = Get-ChildItem -LiteralPath $TemporaryMainFolder
                ForEach ($Folder in $TempMainFolders) {
                    Try {
                        Move-Item -LiteralPath $Folder.FullName -Destination $TemporaryWorkingFolder -Force -Confirm:$false -ErrorAction Stop
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
                Write-HostAndLog -Message ("Archiving the Main BA2 files for group `"{0}`"." -f ($MergeFile.Name)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                Try {
                    $FileName = "{0} - Main.BA2" -f $MergeFile.Name
                    $TargetFilePath = Join-Path -Path $OutputWorkingFolder -ChildPath $FileName
                    $Output = & $BSArchPath pack $TemporaryWorkingFolder $TargetFilePath -share -fo4 2>&1
                    if ($LASTEXITCODE -ne 0) {
                        Throw "BSArch failed to pack $TargetFilePath. Exit Code: $LASTEXITCODE"
                    }
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
                Try {
                    Add-Content -Path $LogFilePath -Value ("{0} - {1}: Archiving command output:`n{2}" -f ($(Get-Date), 'INFO', ($Output | Out-String)))
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }

                Write-HostAndLog -Message "Cleaning up temp folders." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                $TempWorkingFolderList = Get-ChildItem -LiteralPath $("{0}" -f ($TemporaryWorkingFolder))
                ForEach ($Item in $TempWorkingFolderList) {
                    Try {
                        $ProgressPreference = 'SilentlyContinue'
                        Remove-Item -LiteralPath $Item.FullName -Recurse -Force -Confirm:$false -ErrorAction Stop
                        $ProgressPreference = 'Continue'
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
                $TempMainFolderList = Get-ChildItem -LiteralPath $("{0}" -f ($TemporaryMainFolder)) -Recurse
                ForEach ($Item in $TempMainFolderList) {
                    Try {
                        $ProgressPreference = 'SilentlyContinue'
                        Remove-Item -LiteralPath $Item.FullName -Recurse -Force -Confirm:$false -ErrorAction Stop
                        $ProgressPreference = 'Continue'
                    } Catch {
                        Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                    }
                }
            } Else {
                Write-HostAndLog -Message "There aren't any BA2 Main files in staging suitable to be archived." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            }

            # Adding BA2 files and plugin files to a new mod zip archive
            IF (Get-ChildItem -LiteralPath $OutputWorkingFolder) {
                $ProcessedFiles = Get-ChildItem -LiteralPath $BA2MergingStagingFolder -Recurse | Where-Object { $_.Extension -eq ".ba2" }
                $MergedFiles = Get-ChildItem -LiteralPath $OutputWorkingFolder | Where-Object { $_.Extension -eq ".ba2" }
                $ModZipPath = Join-Path -Path $OutputWorkingFolder -ChildPath "$BA2MergedModName.zip"
                Write-HostAndLog -Message ("Combining the BA2 files and plugin files into a new mod zip file `"{0}`"." -f ($ModZipPath)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
                Try {
                    Compress-Archive -Path "$OutputWorkingFolder\*" -DestinationPath $ModZipPath -CompressionLevel NoCompression | Out-Null
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
                Try {
                    $ProgressPreference = 'SilentlyContinue'
                    Get-ChildItem -Path $OutputWorkingFolder -Exclude "$BA2MergedModName.zip" | Remove-Item -Force
                    $ProgressPreference = 'Continue'
                } Catch {
                    Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
                }
                Get-MessageWindow -Title "Install BMAT merged mod" -Message ("To complete the BMAT merge process install the newly created by BMAT mod '{0}' file using your mod manager Vortex or MO2.`n The new mod file can be found in '{1}'" -f ("$BA2MergedModName.zip", $OutputWorkingFolder)) -Category Information | Out-Null
            } Else {
                Write-HostAndLog -Message "There aren't any BA2 merged files or plugin files in staging to be added to the mod manager." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            }
        }

        # Check for missing mod content files
        $TotalProcessedModFilesRawUnique = ($TotalProcessedModFilesRaw | ForEach-Object { $_.File_Path.Replace($CleanUpFolder, '') }) | Sort-Object -Unique
        $TotalProcessedModFilesPreArchUnique = ($TotalProcessedModFilesPreArch | ForEach-Object { $_.File_Path -Replace ("($($TemporaryMainFolder.Replace('\','\\'))|$($TemporaryTexturesFolder.Replace('\','\\')))", '') }) | Sort-Object -Unique
        If ($TotalProcessedModFilesRawUnique.Count -ne $TotalProcessedModFilesPreArchUnique.Count) {
            Write-HostAndLog -Message ("There were {0} mods content files post BA2 files extract and {1} mods content files just before BA2 compression." -f ($TotalProcessedModFilesRawUnique.Count, $TotalProcessedModFilesPreArchUnique.Count)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
            $MissingFiles = Compare-Object $TotalProcessedModFilesRawUnique $TotalProcessedModFilesPreArchUnique | Where-Object { $_.SideIndicator -eq '<=' } | Select-Object -ExpandProperty InputObject
            Write-HostAndLog -Message ("The delta files are:`n{0}" -f ($MissingFiles | Out-String)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
        }
    }

    # Updating repackage register
    Write-HostAndLog -Message "Updating repackage tracker." -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    IF ($ModsChangeTracker) {
        # Checking BMAT tracker file is not opened
        If (Test-Path -LiteralPath $TrackerFolder) {
            Set-FileClosed -FileNamePath $MergedFilesTrackerPath -FileDescription 'BMAT tracker file' -BMATFolder $BMATFolder
        } Else {
            Write-HostAndLog -Message ("Creating BMAT tracker folder `"{0}`"" -f ($TrackerFolder)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
            Try {
                New-Item -Path $TrackerFolder -ItemType Directory -ErrorAction Stop
            } Catch {
                Write-HostAndLog -Message ("Failure with step. Error message:{0}" -f $($_.Exception.Message)) -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            }
        }
        $BackupTrackerName = "{0}-{1}.csv" -f (($MergedFilesTrackerName -split ("\."))[0], (Get-Date -UFormat "%d-%m-%Y-%H-%M-%S"))
        $TrackerFiles = Get-ChildItem (Split-Path -Parent $MergedFilesTrackerPath) -Filter *.csv | Where-Object { $_.BaseName -match "$($MergedFilesTrackerName -replace ('.csv',''))(-\d+){6}" } | Where-Object LastWriteTime -lt (Get-Date).AddDays(-3)
        IF ($TrackerFiles) {
            ForEach ($File in $TrackerFiles) {
                $ProgressPreference = 'SilentlyContinue'
                Remove-Item -LiteralPath $File.FullName
                $ProgressPreference = 'Continue'
            }
        }
        If (Test-Path -LiteralPath $MergedFilesTrackerPath) {
            Try {
                Rename-Item -Path $MergedFilesTrackerPath -NewName $BackupTrackerName -ErrorAction Stop
            } Catch {
                Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
            }
        }
        $TrackerUpdate = @()
        ForEach ($File in ($ModsChangeTracker | Where-Object { $_.Status -ne "Removed_from_Game" -and $_.Status -ne "Removed_from_Mod" })) {
            If ($null -ne ($TrackerUpdate | Where-Object { $_.Mod_Name -eq $File.Mod_Name -and ($File.Mod_BA2_File -ne $null -and $_.Mod_BA2_File -eq $File.Mod_BA2_File) })) {
                Write-HostAndLog -Message ("Duplicate record attempt for mod `"{0}`" and BA2 file: `"{1}`". Skipping..." -f ($File.Mod_Name, $File.Mod_BA2_File)) -Category WARNING -LogfilePath $LogFilePath -DetailedLogging $true
            } Else {
                IF ((($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).Collections | Measure-Object).Count -gt 1) {
                    $ModCollectionsTracker = ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).Collections -join "; "
                } Else {
                    $ModCollectionsTracker = ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).Collections
                }
                If (($File.Plugin_Mod).Count -gt 1) {
                    $PluginMod = $File.Plugin_Mod -join (';')
                } Else {
                    $PluginMod = $File.Plugin_Mod
                }
                $LoadOrder = $File.LO_Index
                $Properties = [ordered]@{
                    "Mod_Id"             = ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).Mod_Id;
                    "Mod_Name"           = $File.Mod_Name;
                    "Mod_Display_Name"   = $File.Mod_Display_Name;
                    "Mod_Version"        = ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).Version;
                    "Mod_Game"           = ($FullModList | Where-Object { $_.Mod_Name -eq $File.Mod_Name }).Game;
                    "Mod_Collections"    = $ModCollectionsTracker;
                    "Mod_Source_Path"    = $File.Mod_Source_Path;
                    "Mod_Target_Path"    = $File.Mod_Target_Path;
                    "Mod_BA2_File"       = $File.Mod_BA2_File;
                    "Mod_Loose_Folder"   = $File.Mod_Loose_Folder;
                    "Mod_Folder"         = $File.Mod_Folder;
                    "Mod_Folder_Path"    = $File.Mod_Folder_Path;
                    "Mod_BA2_Modified"   = $File.Modified.ToString('dd\-MM\-yyyy-HH\-mm\-ss');
                    "Mod_BA2_Size_Bytes" = $File.Size;
                    "Mod_BA2_Type"       = $File.Type;
                    "Repackaging_Group"  = $File.Group;
                    "Included_In_BA2"    = $File.In_BA2;
                    "Plugin_Name"        = $File.Plugin_Name;
                    "LoadOrder"          = $LoadOrder;
                    "Plugin_Mod"         = $PluginMod;
                }
                $TrackerUpdate += New-Object PSObject -prop $Properties
            }
        }
        $TrackerUpdate | Export-Csv -Path $MergedFilesTrackerPath -Encoding ascii -NoClobber
        Try {
            Add-Content -Path $LogFilePath -Value ("{0} - {1}: New tracker details are:`n{2}" -f ($(Get-Date), 'INFO', ($TrackerUpdate | ConvertTo-Json | Out-String)))
        } Catch {
            Write-HostAndLog -Message "Failure with step. Error message: $($_.Exception.Message)" -Category ERROR -LogfilePath $LogFilePath -DetailedLogging $true
        }
        Write-HostAndLog -Message ("BMAT reduced {0} BA2 files down to {1} BA2 files." -f ($ProcessedFiles.Count, $MergedFiles.Count)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
    }

    if (Test-Path $ModZipPath) {
        & explorer.exe (Split-Path -Path $ModZipPath)
    }
}
$EndTime = Get-Date
$LapsedTime = $EndTime - $StartTime
Write-HostAndLog -Message ("It took {0} hours, {1} minutes, and {2} seconds to complete the process this time." -f ($LapsedTime.Hours, $LapsedTime.Minutes, $LapsedTime.Seconds)) -Category INFO -LogfilePath $LogFilePath -DetailedLogging $true
$Tip = Get-RandomTip -ChancePercentage 50
If ($Tip -match "\w") {
    Write-Host "`n[TIP] $Tip`n" -ForegroundColor Yellow
}
Write-Host "`nProcess Complete. Thank you for using BMAT!" -ForegroundColor Cyan
Pause
#endregion MERGE