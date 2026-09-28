#Requires -Version 5.1
<#
    PDF TOOLBOX - window version

    A single window over the same engine as PDFToolbox.ps1. That file must sit
    in this folder: it is loaded for the Python helper, Ghostscript calls,
    Office conversion and the page organiser.

    Start it with "Launch PDF Toolbox Window.bat".

    Built by Ismail (toolkitpdf@gmail.com)
#>

$ErrorActionPreference = 'Stop'

# ----------------------------------------------------------------------
# Load the engine (the console script does not auto-start when dot-sourced)
# ----------------------------------------------------------------------
$engine = Join-Path $PSScriptRoot 'PDFToolbox.ps1'
if (-not (Test-Path -LiteralPath $engine)) {
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [System.Windows.Forms.MessageBox]::Show(
        "PDFToolbox.ps1 was not found in:`n$PSScriptRoot`n`nKeep both files in the same folder.",
        'PDF Toolbox', 'OK', 'Error') | Out-Null
    return
}
. $engine

try {
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Xaml -ErrorAction Stop
} catch {
    [void][System.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms')
    [System.Windows.Forms.MessageBox]::Show(
        "This machine cannot show the window version:`n$($_.Exception.Message)`n`nUse Launch PDFToolbox.bat for the console version instead.",
        'PDF Toolbox', 'OK', 'Error') | Out-Null
    return
}

# The script runs inside a PowerShell console. Hide that console so only the
# app window is on screen. Nothing is printed to it after this point, so
# anything that would have gone there goes to a message box instead.
function Hide-ConsoleWindow {
    try {
        if (-not ('PdfToolbox.ConsoleWindow' -as [type])) {
            Add-Type -Namespace PdfToolbox -Name ConsoleWindow -MemberDefinition @'
[System.Runtime.InteropServices.DllImport("kernel32.dll")]
public static extern System.IntPtr GetConsoleWindow();

[System.Runtime.InteropServices.DllImport("user32.dll")]
public static extern bool ShowWindow(System.IntPtr hWnd, int nCmdShow);
'@ -ErrorAction Stop
        }
        $handle = [PdfToolbox.ConsoleWindow]::GetConsoleWindow()
        if ($handle -ne [IntPtr]::Zero) { [void][PdfToolbox.ConsoleWindow]::ShowWindow($handle, 0) }
    } catch { }
}
Hide-ConsoleWindow

Initialize-Toolbox

# ======================================================================
# WINDOW
# ======================================================================
$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PDF Toolbox" Height="800" Width="1220" MinHeight="620" MinWidth="980"
        WindowStartupLocation="CenterScreen" Background="#14141F"
        FontFamily="Aptos Narrow, Aptos, Segoe UI" FontSize="13">
  <Window.Resources>
    <Style TargetType="Button">
      <Setter Property="Background" Value="#2B2B45"/>
      <Setter Property="Foreground" Value="#E9ECF5"/>
      <Setter Property="BorderBrush" Value="#3C3C5C"/>
      <Setter Property="Padding" Value="10,5"/>
      <Setter Property="Margin" Value="3"/>
    </Style>
    <Style TargetType="Label">
      <Setter Property="Foreground" Value="#9AA3B8"/>
    </Style>
    <Style TargetType="CheckBox">
      <Setter Property="Foreground" Value="#E9ECF5"/>
      <Setter Property="Margin" Value="3,6"/>
    </Style>
    <Style TargetType="TextBox">
      <Setter Property="Background" Value="#1D1D30"/>
      <Setter Property="Foreground" Value="#E9ECF5"/>
      <Setter Property="BorderBrush" Value="#3C3C5C"/>
      <Setter Property="Padding" Value="4,3"/>
      <Setter Property="Margin" Value="3"/>
    </Style>
    <Style TargetType="ComboBox">
      <Setter Property="Margin" Value="3"/>
      <Setter Property="Padding" Value="4,3"/>
      <Setter Property="Foreground" Value="#1B1B2C"/>
    </Style>
    <Style TargetType="ComboBoxItem">
      <Setter Property="Foreground" Value="#1B1B2C"/>
    </Style>
  </Window.Resources>

  <Grid Margin="14" TextElement.Foreground="#E9ECF5">
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- header -->
    <Border Grid.Row="0" Background="#1D1D30" CornerRadius="6" Padding="14,10" Margin="0,0,0,12">
      <DockPanel>
        <StackPanel>
          <TextBlock Text="PDF TOOLBOX" FontSize="19" FontWeight="Bold" Foreground="#4FC3DC"/>
          <TextBlock x:Name="SubTitle" FontSize="12" Foreground="#9AA3B8"/>
        </StackPanel>
        <TextBlock x:Name="VersionText" DockPanel.Dock="Right" HorizontalAlignment="Right"
                   VerticalAlignment="Bottom" Foreground="#6D7390"/>
      </DockPanel>
    </Border>

    <!-- body -->
    <Grid Grid.Row="1">
      <Grid.ColumnDefinitions>
        <ColumnDefinition Width="232"/>
        <ColumnDefinition Width="*"/>
      </Grid.ColumnDefinitions>

      <Border Grid.Column="0" Background="#1D1D30" CornerRadius="6" Padding="6" Margin="0,0,12,0">
        <DockPanel>
          <TextBlock DockPanel.Dock="Top" Text="Task" Margin="8,6,0,6" Foreground="#9AA3B8"/>
          <ListBox x:Name="TaskList" Background="Transparent" BorderThickness="0" Foreground="#E9ECF5"/>
        </DockPanel>
      </Border>

      <Grid Grid.Column="1">
        <Grid.RowDefinitions>
          <RowDefinition Height="2*" MinHeight="180"/>
          <RowDefinition Height="Auto"/>
          <RowDefinition Height="*" MinHeight="120"/>
        </Grid.RowDefinitions>

        <!-- files -->
        <Border Grid.Row="0" Background="#1D1D30" CornerRadius="6" Padding="10">
          <DockPanel>
            <DockPanel DockPanel.Dock="Top" Margin="0,0,0,8">
              <TextBlock x:Name="FilesHeading" Text="Files" FontWeight="Bold" VerticalAlignment="Center"/>
              <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                <Button x:Name="PreviewFileButton" Content="Preview"/>
                <Button x:Name="AddFilesButton" Content="Add files"/>
                <Button x:Name="AddFolderButton" Content="Add folder"/>
                <Button x:Name="RemoveButton" Content="Remove"/>
                <Button x:Name="ClearButton" Content="Clear"/>
                <Button x:Name="MoveUpButton" Content="Up"/>
                <Button x:Name="MoveDownButton" Content="Down"/>
              </StackPanel>
            </DockPanel>
            <TextBlock x:Name="DropHint" DockPanel.Dock="Bottom" Margin="2,6,0,0" Foreground="#6D7390"
                       Text="Drag files or a folder in from Explorer to add them. Drag a row to reorder it. Double-click a row to page through it."/>
            <ListView x:Name="FileList" Background="#F7F8FC" BorderBrush="#C7CDDD" Foreground="#1B1B2C"
                      AllowDrop="True" SelectionMode="Extended">
              <ListView.Resources>
                <!-- the window-wide TextBlock style is light, which would be
                     invisible on this light list, so override it in here -->
                <Style TargetType="TextBlock">
                  <Setter Property="Foreground" Value="#1B1B2C"/>
                </Style>
                <Style TargetType="GridViewColumnHeader">
                  <Setter Property="Foreground" Value="#1B1B2C"/>
                  <Setter Property="Background" Value="#E4E8F2"/>
                  <Setter Property="Padding" Value="6,4"/>
                  <Setter Property="FontWeight" Value="SemiBold"/>
                </Style>
              </ListView.Resources>
              <ListView.ItemContainerStyle>
                <Style TargetType="ListViewItem">
                  <Setter Property="Foreground" Value="#1B1B2C"/>
                  <Setter Property="Padding" Value="2"/>
                </Style>
              </ListView.ItemContainerStyle>
              <ListView.View>
                <GridView>
                  <GridViewColumn Header="File" Width="430" DisplayMemberBinding="{Binding Name}"/>
                  <GridViewColumn Header="Pages" Width="70" DisplayMemberBinding="{Binding Pages}"/>
                  <GridViewColumn Header="Size" Width="90" DisplayMemberBinding="{Binding Size}"/>
                  <GridViewColumn Header="Folder" Width="260" DisplayMemberBinding="{Binding Folder}"/>
                </GridView>
              </ListView.View>
            </ListView>
          </DockPanel>
        </Border>

        <!-- options -->
        <Border Grid.Row="1" Background="#1D1D30" CornerRadius="6" Padding="10" Margin="0,12,0,12">
          <StackPanel>
            <TextBlock x:Name="OptionsHeading" Text="Options" FontWeight="Bold" Margin="0,0,0,6"/>

            <StackPanel x:Name="PanelOcr" Visibility="Collapsed">
              <WrapPanel>
                <Label Content="These files are" VerticalAlignment="Center"/>
                <ComboBox x:Name="OcrMode" Width="240"/>
                <Label Content="Compression" VerticalAlignment="Center" Margin="14,0,0,0"/>
                <ComboBox x:Name="OcrOptimize" Width="150"/>
                <Label Content="At a time" VerticalAlignment="Center" Margin="14,0,0,0"/>
                <ComboBox x:Name="OcrJobs" Width="70"/>
                <CheckBox x:Name="OcrDeskew" Content="Straighten and rotate pages" Margin="16,6,0,6"/>
              </WrapPanel>
            </StackPanel>

            <StackPanel x:Name="PanelMerge" Visibility="Collapsed">
              <WrapPanel>
                <Label Content="Output file name" VerticalAlignment="Center"/>
                <TextBox x:Name="MergeName" Width="260" Text="Combined.pdf"/>
                <Button x:Name="MergeArrangeButton" Content="Arrange visually" Margin="14,3,3,3"/>
                <TextBlock Text="Merged in the order shown above. Drag rows to reorder." VerticalAlignment="Center" Foreground="#6D7390" Margin="14,0,0,0"/>
              </WrapPanel>
            </StackPanel>

            <StackPanel x:Name="PanelPages" Visibility="Collapsed">
              <WrapPanel>
                <Label x:Name="PagesLabel" Content="Pages" VerticalAlignment="Center"/>
                <TextBox x:Name="PagesSpec" Width="230"/>
                <Button x:Name="PickPagesButton" Content="Pick from previews"/>
                <Label x:Name="RotateLabel" Content="Turn" VerticalAlignment="Center" Margin="14,0,0,0"/>
                <ComboBox x:Name="RotateDegrees" Width="190"/>
              </WrapPanel>
              <TextBlock x:Name="PagesHint" Foreground="#6D7390" Margin="4,2,0,0"
                         Text="Examples: 1,3,5   1-10   1-5,10,15-20. Works on the first file in the list."/>
            </StackPanel>

            <StackPanel x:Name="PanelInsert" Visibility="Collapsed">
              <WrapPanel>
                <Label Content="Insert this" VerticalAlignment="Center"/>
                <TextBox x:Name="InsertSource" Width="330" IsReadOnly="True"/>
                <Button x:Name="InsertBrowseButton" Content="Choose PDF or image"/>
                <Label Content="Position" VerticalAlignment="Center" Margin="14,0,0,0"/>
                <ComboBox x:Name="InsertPosition" Width="150"/>
                <Label Content="Before page" VerticalAlignment="Center" Margin="10,0,0,0"/>
                <TextBox x:Name="InsertBefore" Width="60"/>
              </WrapPanel>
              <TextBlock Foreground="#6D7390" Margin="4,2,0,0"
                         Text="Inserted into the first file in the list. An image is sized to match the document."/>
            </StackPanel>

            <StackPanel x:Name="PanelConvert" Visibility="Collapsed">
              <WrapPanel>
                <Label Content="Convert" VerticalAlignment="Center"/>
                <ComboBox x:Name="ConvertKind" Width="330"/>
                <TextBlock x:Name="ConvertHint" VerticalAlignment="Center" Foreground="#6D7390" Margin="14,0,0,0"/>
              </WrapPanel>
            </StackPanel>

            <StackPanel x:Name="PanelCompress" Visibility="Collapsed">
              <WrapPanel>
                <Label Content="How small" VerticalAlignment="Center"/>
                <ComboBox x:Name="CompressLevel" Width="330"/>
              </WrapPanel>
            </StackPanel>

            <StackPanel x:Name="PanelSearch" Visibility="Collapsed">
              <WrapPanel>
                <Label Content="Find this text" VerticalAlignment="Center"/>
                <TextBox x:Name="SearchTerm" Width="330"/>
                <TextBlock Text="Only searchable text is found. Scanned files are reported separately."
                           VerticalAlignment="Center" Foreground="#6D7390" Margin="14,0,0,0"/>
              </WrapPanel>
            </StackPanel>

            <StackPanel x:Name="PanelPassword" Visibility="Collapsed">
              <WrapPanel>
                <Label Content="Password" VerticalAlignment="Center"/>
                <PasswordBox x:Name="PasswordBox" Width="220" Margin="3" Background="#1D1D30" Foreground="#E9ECF5"/>
                <TextBlock Text="Saves an unlocked copy. The original is left alone."
                           VerticalAlignment="Center" Foreground="#6D7390" Margin="14,0,0,0"/>
              </WrapPanel>
            </StackPanel>

            <TextBlock x:Name="PanelNone" Visibility="Collapsed" Foreground="#9AA3B8"
                       Text="Nothing to set for this task. Press Run."/>

            <DockPanel Margin="0,8,0,0">
              <Label Content="Save results to" VerticalAlignment="Center"/>
              <Button x:Name="OutputBrowseButton" Content="Change" DockPanel.Dock="Right"/>
              <TextBox x:Name="OutputFolder" IsReadOnly="True"/>
            </DockPanel>
          </StackPanel>
        </Border>

        <!-- log -->
        <Border Grid.Row="2" Background="#1D1D30" CornerRadius="6" Padding="10">
          <DockPanel>
            <TextBlock DockPanel.Dock="Top" Text="Progress" FontWeight="Bold" Margin="0,0,0,6"/>
            <TextBox x:Name="LogBox" Background="#F7F8FC" BorderBrush="#C7CDDD" Foreground="#1B1B2C"
                     IsReadOnly="True" VerticalScrollBarVisibility="Auto"
                     TextWrapping="NoWrap" HorizontalScrollBarVisibility="Auto"/>
          </DockPanel>
        </Border>

        <Border Grid.Row="0" Grid.RowSpan="3" x:Name="AboutPanel" Background="#1D1D30" CornerRadius="6"
                Padding="26,22" Visibility="Collapsed">
          <ScrollViewer VerticalScrollBarVisibility="Auto">
            <StackPanel>
              <TextBlock Text="PDF TOOLBOX" FontSize="26" FontWeight="Bold" Foreground="#4FC3DC"/>
              <TextBlock x:Name="AboutLine1" FontSize="14" Foreground="#9AA3B8" Margin="0,4,0,0"/>
              <TextBlock x:Name="AboutLine2" FontSize="14" Foreground="#9AA3B8" Margin="0,2,0,0"/>
              <Separator Margin="0,16" Background="#33334D"/>

              <TextBlock Text="WHAT IT DOES" FontWeight="Bold" Foreground="#4FC3DC" Margin="0,0,0,6"/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Text="OCR, merge, split, extract, delete, rotate, insert, rearrange, compress, search and count PDFs, remove known passwords, and convert both to and from PDF. Built for audit work: scanned bank statements, working papers and supporting documents."/>

              <TextBlock Text="WHAT IT RUNS ON" FontWeight="Bold" Foreground="#4FC3DC" Margin="0,18,0,6"/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Text="OCRmyPDF, Tesseract and Ghostscript do the OCR and compression. pypdf, pdfplumber, openpyxl and Pillow handle the page and table work. Microsoft Word, Excel and PowerPoint do the Office conversions. All open-source except Office. Everything runs on this machine; no file is uploaded anywhere."/>

              <TextBlock Text="KNOW THIS BEFORE YOU TRUST THE OUTPUT" FontWeight="Bold" Foreground="#E3B341" Margin="0,18,0,6"/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Margin="0,0,0,3" Text="A scanned PDF has no text until you OCR it. Text extraction, table extraction and search all return nothing until then."/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Margin="0,0,0,3" Text="PDF to Excel guesses where the columns are. Tie the totals back to the PDF before using any figure."/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Margin="0,0,0,3" Text="PDF to Word rebuilds the layout, it does not copy it. Complex tables usually need tidying."/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Margin="0,0,0,3" Text="OCR is not perfect. Spot-check figures on anything that matters."/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Margin="0,0,0,3" Text="Compressing re-writes the PDF. Check one file before replacing originals."/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Text="Every task writes a new file. Your originals are left alone."/>

              <TextBlock Text="UPDATES" FontWeight="Bold" Foreground="#4FC3DC" Margin="0,18,0,6"/>
              <StackPanel Orientation="Horizontal">
                <Button x:Name="AboutUpdateButton" Content="Check for a new version" Width="190"/>
                <TextBlock x:Name="AboutUpdateStatus" VerticalAlignment="Center" Margin="12,0,0,0" Foreground="#9AA3B8" TextWrapping="Wrap"/>
              </StackPanel>
              <TextBlock x:Name="AboutDownloadLine" Visibility="Collapsed" Margin="0,6,0,0" TextWrapping="Wrap">
                <Hyperlink x:Name="AboutDownloadLink" Foreground="#4FC3DC"><Run Text="Open the download page"/></Hyperlink>
              </TextBlock>

              <TextBlock Text="IF SOMETHING GOES WRONG" FontWeight="Bold" Foreground="#4FC3DC" Margin="0,18,0,6"/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Text="Run Check setup first: most faults are a missing install."/>
              <TextBlock TextWrapping="Wrap" Foreground="#E9ECF5" Margin="0,3,0,0">
                <Run Text="Still stuck, or want something added: "/><Hyperlink x:Name="AboutMailLink" Foreground="#4FC3DC"><Run x:Name="AboutMailText"/></Hyperlink>
              </TextBlock>
              <TextBlock x:Name="AboutContact" TextWrapping="Wrap" Foreground="#9AA3B8" Margin="0,3,0,0"/>
            </StackPanel>
          </ScrollViewer>
        </Border>

        <!-- check setup (full page, like About) -->
        <Border Grid.Row="0" Grid.RowSpan="3" x:Name="SetupPanel" Background="#1D1D30" CornerRadius="6"
                Padding="26,22" Visibility="Collapsed">
          <DockPanel>
            <StackPanel DockPanel.Dock="Top">
              <TextBlock Text="CHECK SETUP" FontSize="26" FontWeight="Bold" Foreground="#4FC3DC"/>
              <TextBlock Text="What this machine has installed, and what each part is used for."
                         FontSize="14" Foreground="#9AA3B8" Margin="0,4,0,0"/>
              <Separator Margin="0,16" Background="#33334D"/>
            </StackPanel>
            <StackPanel DockPanel.Dock="Bottom" Margin="0,16,0,0">
              <TextBlock x:Name="SetupSummary" TextWrapping="Wrap" Foreground="#E3B341"/>
              <StackPanel Orientation="Horizontal" Margin="0,10,0,0">
                <Button x:Name="SetupInstallButton" Content="Install what is missing" Width="200" FontWeight="Bold"/>
                <Button x:Name="SetupUpdateButton" Content="Check for updates" Width="160"/>
                <Button x:Name="SetupUpdateAllButton" Content="Update all" Width="120" Visibility="Collapsed"/>
                <Button x:Name="SetupRefreshButton" Content="Check again" Width="130"/>
              </StackPanel>
            </StackPanel>
            <ScrollViewer VerticalScrollBarVisibility="Auto">
              <StackPanel x:Name="SetupList"/>
            </ScrollViewer>
          </DockPanel>
        </Border>

      </Grid>
    </Grid>

    <!-- footer -->
    <Border Grid.Row="2" Background="#1D1D30" CornerRadius="6" Padding="10" Margin="0,12,0,0">
      <DockPanel>
        <ProgressBar x:Name="Progress" Width="260" Height="16" DockPanel.Dock="Left" Margin="4"/>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Right">
          <Button x:Name="OpenFolderButton" Content="Open output folder" Width="150"/>
          <Button x:Name="RunButton" Content="Run" Width="120" FontWeight="Bold"/>
        </StackPanel>
        <TextBlock x:Name="StatusText" VerticalAlignment="Center" Margin="14,0" Foreground="#9AA3B8" Text="Ready."/>
      </DockPanel>
    </Border>
  </Grid>
</Window>
'@

try {
    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    [System.Windows.MessageBox]::Show(
        "Could not build the window:`n$($_.Exception.Message)",
        'PDF Toolbox', 'OK', 'Error') | Out-Null
    return
}

$ui = @{}
foreach ($name in @('SubTitle', 'VersionText', 'TaskList', 'FilesHeading', 'AddFilesButton', 'AddFolderButton',
        'RemoveButton', 'ClearButton', 'MoveUpButton', 'MoveDownButton', 'FileList', 'DropHint', 'OptionsHeading',
        'PreviewFileButton',
        'PanelOcr', 'OcrMode', 'OcrOptimize', 'OcrJobs', 'OcrDeskew',
        'PanelMerge', 'MergeName', 'MergeArrangeButton',
        'PanelPages', 'PagesLabel', 'PagesSpec', 'PickPagesButton', 'RotateLabel', 'RotateDegrees', 'PagesHint',
        'PanelInsert', 'InsertSource', 'InsertBrowseButton', 'InsertPosition', 'InsertBefore',
        'PanelConvert', 'ConvertKind', 'ConvertHint', 'PanelCompress', 'CompressLevel',
        'PanelSearch', 'SearchTerm', 'PanelPassword', 'PasswordBox', 'PanelNone',
        'OutputFolder', 'OutputBrowseButton', 'LogBox', 'Progress', 'StatusText', 'RunButton', 'OpenFolderButton',
        'AboutPanel', 'AboutLine1', 'AboutLine2', 'AboutContact', 'AboutMailLink', 'AboutMailText',
        'AboutUpdateButton', 'AboutUpdateStatus', 'AboutDownloadLine', 'AboutDownloadLink',
        'SetupPanel', 'SetupList', 'SetupSummary', 'SetupRefreshButton', 'SetupInstallButton',
        'SetupUpdateButton', 'SetupUpdateAllButton')) {
    $ui[$name] = $window.FindName($name)
}

# ======================================================================
# SMALL UI HELPERS
# ======================================================================
$script:Files = New-Object 'System.Collections.ObjectModel.ObservableCollection[object]'
$script:CurrentTask = $null
$script:LastOutput = $null
$ui.FileList.ItemsSource = $script:Files

# Lets the window repaint during long work without a second thread.
function Update-Ui {
    $frame = New-Object System.Windows.Threading.DispatcherFrame
    [System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke(
        [System.Windows.Threading.DispatcherPriority]::Background,
        [System.Windows.Threading.DispatcherOperationCallback] { param($f) $f.Continue = $false; return $null },
        $frame) | Out-Null
    [System.Windows.Threading.Dispatcher]::PushFrame($frame)
}

function Write-Log {
    param([string]$Text)
    $ui.LogBox.AppendText($Text + "`r`n")
    $ui.LogBox.ScrollToEnd()
    Update-Ui
}

function Clear-Log { $ui.LogBox.Clear() }

function Set-Status {
    param([string]$Text, [switch]$Refresh)
    $ui.StatusText.Text = $Text
    # Only pump the message queue during long work, never from an event handler.
    if ($Refresh) { Update-Ui }
}

function Set-Progress {
    param([int]$Done, [int]$Total)
    $ui.Progress.Maximum = [Math]::Max(1, $Total)
    $ui.Progress.Value = $Done
    Update-Ui
}

function Show-Message {
    param([string]$Text, [string]$Title = 'PDF Toolbox')
    [System.Windows.MessageBox]::Show($window, $Text, $Title, 'OK', 'Information') | Out-Null
}

function Get-SelectedOrAllFiles {
    if ($ui.FileList.SelectedItems.Count -gt 0) { return @($ui.FileList.SelectedItems) }
    return @($script:Files)
}

function Add-FilePaths {
    param([string[]]$Paths, [string[]]$Extensions)
    $added = 0
    foreach ($path in $Paths) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        $item = Get-Item -LiteralPath $path
        $candidates = if ($item.PSIsContainer) {
            @(Get-ChildItem -LiteralPath $path -File | Where-Object { $Extensions -contains $_.Extension.ToLower() -and $_.Name -notlike '~$*' } | Sort-Object Name)
        } else { @($item) }
        foreach ($file in $candidates) {
            if ($Extensions -notcontains $file.Extension.ToLower()) { continue }
            if (@($script:Files | Where-Object { $_.Path -eq $file.FullName }).Count -gt 0) { continue }
            $pages = ''
            if ($file.Extension -ieq '.pdf') {
                $info = Get-PdfInfo $file.FullName
                $pages = if ($null -eq $info) { '?' } elseif ($info.Pages -lt 0) { 'locked' } else { "$($info.Pages)" }
            }
            $script:Files.Add([pscustomobject]@{
                    Name   = $file.Name
                    Pages  = $pages
                    Size   = (Format-FileSize $file.Length)
                    Folder = $file.DirectoryName
                    Path   = $file.FullName
                })
            $added++
        }
    }
    if ($added -gt 0 -and -not $ui.OutputFolder.Text) {
        $ui.OutputFolder.Text = Join-Path $script:Files[0].Folder (Get-DefaultOutputName)
    }
    Set-Status "$($script:Files.Count) file(s) in the list."
}

function Get-DefaultOutputName {
    switch ($script:CurrentTask) {
        'ocr' { 'OCR PDFs' }
        'compress' { 'Compressed' }
        'toword' { 'Word' }
        'toexcel' { 'Excel' }
        'totext' { 'Text' }
        'toimages' { 'Images' }
        'officetopdf' { 'PDFs' }
        default { 'PDF Toolbox output' }
    }
}

function Get-TaskExtensions {
    switch ($script:CurrentTask) {
        'officetopdf' { $script:OfficeExtensions }
        'imagestopdf' { $script:ImageExtensions }
        default { @('.pdf') }
    }
}

# ======================================================================
# TASKS
# ======================================================================
$script:Tasks = @(
    @{ Key = 'ocr'; Title = 'OCR (make searchable)'; Panel = 'PanelOcr'; Files = 'many'; Heading = 'PDFs to OCR' },
    @{ Key = 'merge'; Title = 'Merge PDFs'; Panel = 'PanelMerge'; Files = 'many'; Heading = 'PDFs to merge, in order' },
    @{ Key = 'extract'; Title = 'Extract pages'; Panel = 'PanelPages'; Files = 'one'; Heading = 'PDF to take pages from' },
    @{ Key = 'delete'; Title = 'Delete pages'; Panel = 'PanelPages'; Files = 'one'; Heading = 'PDF to remove pages from' },
    @{ Key = 'rotate'; Title = 'Rotate pages'; Panel = 'PanelPages'; Files = 'one'; Heading = 'PDF to rotate' },
    @{ Key = 'split'; Title = 'Split into single pages'; Panel = 'PanelNone'; Files = 'one'; Heading = 'PDF to split' },
    @{ Key = 'rearrange'; Title = 'Rearrange pages'; Panel = 'PanelNone'; Files = 'one'; Heading = 'PDF to rearrange' },
    @{ Key = 'insert'; Title = 'Insert pages or an image'; Panel = 'PanelInsert'; Files = 'one'; Heading = 'PDF to insert into' },
    @{ Key = 'compress'; Title = 'Compress'; Panel = 'PanelCompress'; Files = 'many'; Heading = 'PDFs to compress' },
    @{ Key = 'convert'; Title = 'Convert'; Panel = 'PanelConvert'; Files = 'many'; Heading = 'Files to convert' },
    @{ Key = 'search'; Title = 'Search text'; Panel = 'PanelSearch'; Files = 'many'; Heading = 'PDFs to search' },
    @{ Key = 'count'; Title = 'Count pages'; Panel = 'PanelNone'; Files = 'many'; Heading = 'PDFs to count' },
    @{ Key = 'password'; Title = 'Remove a password'; Panel = 'PanelPassword'; Files = 'one'; Heading = 'Password-protected PDF' },
    @{ Key = 'setup'; Title = 'Check setup'; Panel = 'PanelNone'; Files = 'none'; Heading = 'Nothing needed' },
    @{ Key = 'about'; Title = 'About this tool'; Panel = 'PanelNone'; Files = 'none'; Heading = 'Nothing needed' }
)

# winget and pip change the PATH for new processes, not for this one.
function Update-SessionPath {
    $machine = [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $user = [Environment]::GetEnvironmentVariable('Path', 'User')
    $env:Path = (@($machine, $user) | Where-Object { $_ }) -join ';'
    foreach ($extra in @('C:\Program Files\Tesseract-OCR')) {
        if ((Test-Path -LiteralPath $extra) -and ($env:Path -notlike "*$extra*")) { $env:Path += ";$extra" }
    }
    $gs = Get-ChildItem 'C:\Program Files\gs' -Recurse -Filter gswin64c.exe -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($gs -and ($env:Path -notlike "*$($gs.DirectoryName)*")) { $env:Path += ";$($gs.DirectoryName)" }
    if (Get-Command py -ErrorAction SilentlyContinue) { $script:PyExe = 'py' }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { $script:PyExe = 'python' }
}

# Runs the installer in its own console window. It stays interactive there, so
# it can ask about admin rights or a blocked proxy, which a button cannot.
function Start-Installer {
    $installer = Join-Path $PSScriptRoot 'InstallPDFToolbox.ps1'
    if (-not (Test-Path -LiteralPath $installer)) {
        Show-Message "InstallPDFToolbox.ps1 is not in this folder:`n$PSScriptRoot`n`nKeep all the toolbox files together."
        return
    }
    $ui.SetupInstallButton.IsEnabled = $false
    $ui.SetupRefreshButton.IsEnabled = $false
    $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::Goldenrod
    $ui.SetupSummary.Text = 'The installer is running in a separate window. Answer its questions there, then come back.'
    Update-Ui
    try {
        $process = Start-Process -FilePath 'powershell.exe' `
            -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"{0}"' -f $installer)) -PassThru
        $null = $process.Handle
        while (-not $process.HasExited) {
            Start-Sleep -Milliseconds 400
            Update-Ui
        }
    } catch {
        $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::IndianRed
        $ui.SetupSummary.Text = "Could not start the installer: $($_.Exception.Message)"
        $ui.SetupInstallButton.IsEnabled = $true
        $ui.SetupRefreshButton.IsEnabled = $true
        return
    }
    Update-SessionPath
    $ui.SetupInstallButton.IsEnabled = $true
    $ui.SetupRefreshButton.IsEnabled = $true
    Show-SetupPage
}

$script:UpdateInfo = @{}

# Runs a command out of process and returns its output, without freezing the
# window. Used for pip and winget, which are both slow on a company network.
function Invoke-Tool {
    param([string]$FilePath, [string[]]$Arguments, [int]$TimeoutSeconds = 180)
    $id = [guid]::NewGuid().ToString('N')
    $outFile = Join-Path $script:WorkDir "$id.out.txt"
    $errFile = Join-Path $script:WorkDir "$id.err.txt"
    try {
        $process = Start-Process -FilePath $FilePath -ArgumentList $Arguments -PassThru -NoNewWindow `
            -RedirectStandardOutput $outFile -RedirectStandardError $errFile
        $null = $process.Handle
        $waited = 0
        while (-not $process.HasExited) {
            Start-Sleep -Milliseconds 300
            $waited += 0.3
            Update-Ui
            if ($waited -ge $TimeoutSeconds) { try { $process.Kill() } catch { }; break }
        }
        $text = ''
        if (Test-Path -LiteralPath $outFile) { $text = (Get-Content -LiteralPath $outFile -Raw -ErrorAction SilentlyContinue) }
        $errorText = ''
        if (Test-Path -LiteralPath $errFile) { $errorText = (Get-Content -LiteralPath $errFile -Raw -ErrorAction SilentlyContinue) }
        return [pscustomobject]@{ Code = $process.ExitCode; Output = "$text"; Error = "$errorText" }
    } catch {
        return [pscustomobject]@{ Code = -1; Output = ''; Error = $_.Exception.Message }
    } finally {
        Remove-Item -LiteralPath $outFile, $errFile -Force -ErrorAction SilentlyContinue
    }
}

# The packages the toolbox itself uses. pip reports every outdated package on
# the machine, so anything not in this list is none of our business.
$script:PipComponents = @('pypdf', 'cryptography', 'ocrmypdf', 'pdfplumber', 'openpyxl', 'pillow')

function Get-PipOutdated {
    $found = @{}
    if (-not $script:PyExe) { return $found }
    $result = Invoke-Tool -FilePath $script:PyExe -Arguments @('-m', 'pip', 'list', '--outdated', '--format=json')
    if ($result.Code -ne 0 -or -not $result.Output.Trim()) { return $found }
    try {
        foreach ($row in ($result.Output | ConvertFrom-Json)) {
            $name = $row.name.ToLower()
            if ($script:PipComponents -notcontains $name) { continue }
            $found[$name] = [pscustomobject]@{ Current = $row.version; Latest = $row.latest_version }
        }
    } catch { }
    return $found
}

# winget tells us per package. Its table output is localised and awkward, so
# only the presence of the id in an upgrade listing is trusted.
function Get-WingetOutdated {
    param([string]$Id)
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { return $null }
    $result = Invoke-Tool -FilePath 'winget' -Arguments @('upgrade', '--id', $Id, '--exact', '--include-unknown',
        '--accept-source-agreements', '--disable-interactivity')
    if ($result.Code -ne 0) { return $null }
    if ($result.Output -notmatch [regex]::Escape($Id)) { return $null }
    $version = ''
    foreach ($line in ($result.Output -split "`r?`n")) {
        if ($line -match [regex]::Escape($Id)) {
            $bits = @($line -split '\s{2,}' | Where-Object { $_.Trim() })
            if ($bits.Count -ge 3) { $version = $bits[-2] }
            break
        }
    }
    return [pscustomobject]@{ Current = ''; Latest = $version }
}

function Get-UpdateInfo {
    $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::Goldenrod
    $ui.SetupSummary.Text = 'Checking for updates. This goes out to the internet, so it can take a minute...'
    Update-Ui
    $script:UpdateInfo = @{}
    foreach ($pair in (Get-PipOutdated).GetEnumerator()) { $script:UpdateInfo[$pair.Key] = $pair.Value }
    foreach ($item in @(
            @{ Key = 'ghostscript'; Id = 'Artifex.GhostScript' },
            @{ Key = 'tesseract'; Id = 'UB-Mannheim.TesseractOCR' })) {
        $found = Get-WingetOutdated -Id $item.Id
        if ($found) { $script:UpdateInfo[$item.Key] = $found }
    }
    $script:UpdatesChecked = $true
    Show-SetupPage
    $count = $script:UpdateInfo.Keys.Count
    $ui.SetupSummary.Foreground = $(if ($count) { [System.Windows.Media.Brushes]::Goldenrod } else { [System.Windows.Media.Brushes]::MediumSeaGreen })
    $ui.SetupSummary.Text = $(if ($count) {
            "$count of the toolbox's components have a newer version. Update them individually, or use Update all. Ghostscript and Pillow are the two worth keeping current."
        } else { 'Every toolbox component is on its latest version. Other Python packages on this machine are not checked.' })
}

# Does the update itself. pip needs nothing from the user; winget may raise a
# Windows admin prompt, which has to be accepted for the upgrade to go through.
function Update-Component {
    param([string]$Label, [string]$PipName, [string]$WingetId, [switch]$SkipRefresh)
    $ui.SetupInstallButton.IsEnabled = $false
    $ui.SetupUpdateButton.IsEnabled = $false
    $ui.SetupRefreshButton.IsEnabled = $false
    $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::Goldenrod
    $ui.SetupSummary.Text = "Updating $Label..."
    Update-Ui
    if ($PipName) {
        $result = Invoke-Tool -FilePath $script:PyExe -Arguments @('-m', 'pip', 'install', '--user', '--upgrade', $PipName) -TimeoutSeconds 300
    } else {
        $ui.SetupSummary.Text = "Updating $Label. Windows may ask for permission: accept it to let the upgrade finish."
        Update-Ui
        $result = Invoke-Tool -FilePath 'winget' -Arguments @('upgrade', '--id', $WingetId, '--exact',
            '--accept-source-agreements', '--accept-package-agreements') -TimeoutSeconds 600
    }
    Update-SessionPath
    $key = $(if ($PipName) { $PipName.ToLower() } else { $Label.ToLower() })
    if ($result.Code -eq 0) { $script:UpdateInfo.Remove($key) }
    if ($SkipRefresh) { return ($result.Code -eq 0) }
    Show-SetupPage
    $ui.SetupInstallButton.IsEnabled = $true
    $ui.SetupUpdateButton.IsEnabled = $true
    $ui.SetupRefreshButton.IsEnabled = $true
    if ($result.Code -eq 0) {
        $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::MediumSeaGreen
        $ui.SetupSummary.Text = "$Label updated. Close and reopen the toolbox so everything picks up the new version."
    } else {
        $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::IndianRed
        $reason = ("$($result.Error)".Trim() -split "`r?`n" | Where-Object { $_.Trim() } | Select-Object -Last 1)
        if (-not $reason) { $reason = "exit code $($result.Code)" }
        $ui.SetupSummary.Text = "$Label did not update: $reason"
    }
    return ($result.Code -eq 0)
}

# Updates everything that has a newer version, one after another.
function Update-AllComponents {
    $pending = @($script:ComponentList | Where-Object {
            $key = $(if ($_.Pip) { $_.Pip.ToLower() } else { $_.Name.ToLower() })
            $script:UpdateInfo.ContainsKey($key)
        })
    if ($pending.Count -eq 0) { return }
    $done = 0
    $failed = @()
    foreach ($item in $pending) {
        $ok = Update-Component -Label $item.Name -PipName $item.Pip -WingetId $item.Winget -SkipRefresh
        if ($ok) { $done++ } else { $failed += $item.Name }
    }
    Show-SetupPage
    $ui.SetupInstallButton.IsEnabled = $true
    $ui.SetupUpdateButton.IsEnabled = $true
    $ui.SetupRefreshButton.IsEnabled = $true
    if ($failed.Count -eq 0) {
        $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::MediumSeaGreen
        $ui.SetupSummary.Text = "$done component(s) updated. Close and reopen the toolbox so everything picks up the new versions."
    } else {
        $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::IndianRed
        $ui.SetupSummary.Text = "$done updated. These did not: $($failed -join ', '). Try them one at a time to see why."
    }
}

function New-CheckRow {
    param([string]$Name, [bool]$Ok, [string]$Detail, [string]$PipName, [string]$WingetId, $Update)
    $row = New-Object System.Windows.Controls.DockPanel
    $row.Margin = New-Object System.Windows.Thickness 0, 3, 0, 3
    $status = New-Object System.Windows.Controls.TextBlock
    $status.Text = $(if ($Ok) { 'OK' } else { 'MISSING' })
    $status.Width = 90
    $status.FontWeight = 'Bold'
    $status.Foreground = $(if ($Ok) { [System.Windows.Media.Brushes]::MediumSeaGreen } else { [System.Windows.Media.Brushes]::IndianRed })
    [void]$row.Children.Add($status)
    $label = New-Object System.Windows.Controls.TextBlock
    $label.Text = $Name
    $label.Width = 150
    $label.Foreground = [System.Windows.Media.Brushes]::White
    [void]$row.Children.Add($label)
    # An Update button, only when a newer version was actually found.
    if ($Ok -and $Update -and ($PipName -or $WingetId)) {
        $button = New-Object System.Windows.Controls.Button
        $button.Content = 'Update'
        $button.Width = 80
        $button.Margin = New-Object System.Windows.Thickness 0, 0, 10, 0
        $button.Tag = [pscustomobject]@{ Label = $Name; Pip = $PipName; Winget = $WingetId }
        $button.Add_Click({
                param($sender, $eventArgs)
                $info = $sender.Tag
                Update-Component -Label $info.Label -PipName $info.Pip -WingetId $info.Winget
            })
        [void]$row.Children.Add($button)

        $newer = New-Object System.Windows.Controls.TextBlock
        $newer.Width = 180
        $newer.Foreground = [System.Windows.Media.Brushes]::Goldenrod
        $newer.Text = $(if ($Update.Current) { "$($Update.Current) -> $($Update.Latest)" } else { "newer version: $($Update.Latest)" })
        [void]$row.Children.Add($newer)
    }

    # NOTE: do not call this $detail. PowerShell variable names are
    # case-insensitive, so it would collide with the [string]$Detail
    # parameter and turn the control into a string.
    $detailBlock = New-Object System.Windows.Controls.TextBlock
    $detailBlock.Text = $Detail
    $detailBlock.TextWrapping = 'Wrap'
    $detailBlock.Foreground = [System.Windows.Media.Brushes]::Gray
    [void]$row.Children.Add($detailBlock)
    return $row
}

# Every probe here is wrapped: a check that fails must show as MISSING,
# never stop the page from drawing.
function Test-Quietly {
    param([scriptblock]$Probe)
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { return [bool](& $Probe) } catch { return $false }
    finally { $ErrorActionPreference = $previous }
}

function Show-SetupPage {
    $ui.SetupList.Children.Clear()
    $ui.SetupSummary.Text = 'Checking...'
    $previous = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $pyOk = [bool]$script:PyExe
    $pyText = if ($pyOk) { "$(& $script:PyExe --version 2>&1)" } else { 'needed for everything except Office conversion' }
    $tess = Get-Command tesseract -ErrorAction SilentlyContinue
    $gs = Get-GhostscriptExe
    $script:ComponentList = @(
        @{ Name = 'pypdf'; Pip = 'pypdf' },
        @{ Name = 'cryptography'; Pip = 'cryptography' },
        @{ Name = 'OCRmyPDF'; Pip = 'ocrmypdf' },
        @{ Name = 'pdfplumber'; Pip = 'pdfplumber' },
        @{ Name = 'openpyxl'; Pip = 'openpyxl' },
        @{ Name = 'Pillow'; Pip = 'pillow' },
        @{ Name = 'Tesseract'; Winget = 'UB-Mannheim.TesseractOCR' },
        @{ Name = 'Ghostscript'; Winget = 'Artifex.GhostScript' }
    )
    $checks = @(
        @{ Name = 'Python'; Ok = $pyOk; Detail = $pyText },
        @{ Name = 'pypdf'; Ok = (Test-Quietly { Test-PyModule 'pypdf' }); Detail = 'merge, split, extract, delete, rotate, insert, count'; Pip = 'pypdf' },
        @{ Name = 'cryptography'; Ok = (Test-Quietly { Test-PyModule 'cryptography' }); Detail = 'opens password-protected PDFs'; Pip = 'cryptography' },
        @{ Name = 'OCRmyPDF'; Ok = (Test-Quietly { [bool](Get-OcrCommand) }); Detail = 'makes scanned PDFs searchable. Update only when you can re-test a batch'; Pip = 'ocrmypdf' },
        @{ Name = 'Tesseract'; Ok = [bool]$tess; Detail = $(if ($tess) { $tess.Source } else { 'the OCR engine itself' }); Winget = 'UB-Mannheim.TesseractOCR' },
        @{ Name = 'Ghostscript'; Ok = [bool]$gs; Detail = $(if ($gs) { "$gs  (worth keeping current: it reads untrusted PDFs)" } else { 'compression, PDF to images, page previews' }); Winget = 'Artifex.GhostScript' },
        @{ Name = 'pdfplumber'; Ok = (Test-Quietly { Test-PyModule 'pdfplumber' }); Detail = 'finds tables for PDF to Excel'; Pip = 'pdfplumber' },
        @{ Name = 'openpyxl'; Ok = (Test-Quietly { Test-PyModule 'openpyxl' }); Detail = 'writes .xlsx files'; Pip = 'openpyxl' },
        @{ Name = 'Pillow'; Ok = (Test-Quietly { Test-PyModule 'PIL' }); Detail = 'images to PDF, and images inserted into a PDF. Worth keeping current'; Pip = 'pillow' },
        @{ Name = 'Word'; Ok = (Test-Quietly { Test-Path 'Registry::HKEY_CLASSES_ROOT\Word.Application' }); Detail = 'Office to PDF, and PDF to Word' },
        @{ Name = 'Excel'; Ok = (Test-Quietly { Test-Path 'Registry::HKEY_CLASSES_ROOT\Excel.Application' }); Detail = 'Excel to PDF' },
        @{ Name = 'PowerPoint'; Ok = (Test-Quietly { Test-Path 'Registry::HKEY_CLASSES_ROOT\PowerPoint.Application' }); Detail = 'PowerPoint to PDF' }
    )
    foreach ($check in $checks) {
        $key = $(if ($check.Pip) { $check.Pip.ToLower() } elseif ($check.Winget) { $check.Name.ToLower() } else { '' })
        $update = $(if ($key -and $script:UpdateInfo.ContainsKey($key)) { $script:UpdateInfo[$key] } else { $null })
        [void]$ui.SetupList.Children.Add((New-CheckRow -Name $check.Name -Ok $check.Ok -Detail $check.Detail `
                    -PipName $check.Pip -WingetId $check.Winget -Update $update))
    }
    $ErrorActionPreference = $previous
    $missing = @($checks | Where-Object { -not $_.Ok } | ForEach-Object { $_.Name })
    $ui.SetupUpdateAllButton.Visibility = $(if ($script:UpdateInfo.Keys.Count -gt 0) { 'Visible' } else { 'Collapsed' })
    $office = @('Word', 'Excel', 'PowerPoint')
    $installable = @($missing | Where-Object { $office -notcontains $_ })
    if ($missing.Count -eq 0) {
        $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::MediumSeaGreen
        $ui.SetupSummary.Text = 'Everything is installed. Nothing to do.'
        $ui.SetupInstallButton.Visibility = 'Collapsed'
    } else {
        $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::Goldenrod
        $text = "Missing: $($missing -join ', ')."
        if (-not $script:UpdatesChecked) { $text += '  Press Check for updates to see whether anything has a newer version.' }
        if ($installable.Count -gt 0) {
            $text += '  Press Install what is missing. It opens the installer in its own window, skips anything already there, and this page refreshes when it finishes.'
            $ui.SetupInstallButton.Visibility = 'Visible'
        } else {
            $ui.SetupInstallButton.Visibility = 'Collapsed'
        }
        if (@($missing | Where-Object { $office -contains $_ }).Count -gt 0) {
            $text += '  Word, Excel and PowerPoint cannot be installed from here: those come from IT.'
        }
        $ui.SetupSummary.Text = $text
    }
    Set-Status 'Check setup.'
}

function Get-Task {
    param([string]$Key)
    return ($script:Tasks | Where-Object { $_.Key -eq $Key } | Select-Object -First 1)
}

function Select-Task {
    param([string]$Key)
    $task = Get-Task $Key
    if (-not $task) { return }
    $script:CurrentTask = $Key
    foreach ($panel in @('PanelOcr', 'PanelMerge', 'PanelPages', 'PanelInsert', 'PanelConvert',
            'PanelCompress', 'PanelSearch', 'PanelPassword', 'PanelNone')) {
        $ui[$panel].Visibility = 'Collapsed'
    }
    $ui[$task.Panel].Visibility = 'Visible'
    $ui.FilesHeading.Text = $task.Heading
    $ui.OptionsHeading.Text = $task.Title

    # page-based tasks share one panel, so adjust its labels
    $isRotate = ($Key -eq 'rotate')
    $ui.RotateLabel.Visibility = $(if ($isRotate) { 'Visible' } else { 'Collapsed' })
    $ui.RotateDegrees.Visibility = $(if ($isRotate) { 'Visible' } else { 'Collapsed' })
    switch ($Key) {
        'extract' { $ui.PagesLabel.Content = 'Pages to extract' }
        'delete' { $ui.PagesLabel.Content = 'Pages to remove' }
        'rotate' { $ui.PagesLabel.Content = 'Pages to rotate' }
    }
    if ($Key -eq 'rotate' -and -not $ui.PagesSpec.Text) { $ui.PagesSpec.Text = 'all' }

    $ui.MoveUpButton.Visibility = $(if ($Key -eq 'merge') { 'Visible' } else { 'Collapsed' })
    $ui.MoveDownButton.Visibility = $ui.MoveUpButton.Visibility

    if ($script:Files.Count -gt 0) {
        $ui.OutputFolder.Text = Join-Path $script:Files[0].Folder (Get-DefaultOutputName)
    }
    if ($Key -eq 'setup') {
        $ui.SetupPanel.Visibility = 'Visible'
        $ui.RunButton.Visibility = 'Collapsed'
        $ui.OpenFolderButton.Visibility = 'Collapsed'
        $ui.Progress.Visibility = 'Collapsed'
        try { Show-SetupPage }
        catch {
            $ui.SetupSummary.Foreground = [System.Windows.Media.Brushes]::IndianRed
            $ui.SetupSummary.Text = "The check could not finish: $($_.Exception.Message)"
        }
        return
    }
    $ui.SetupPanel.Visibility = 'Collapsed'
    if ($Key -eq 'about') {
        $ui.AboutPanel.Visibility = 'Visible'
        $ui.RunButton.Visibility = 'Collapsed'
        $ui.OpenFolderButton.Visibility = 'Collapsed'
        $ui.Progress.Visibility = 'Collapsed'
        Set-Status "PDF Toolbox v$script:Version, built by $script:Author."
        return
    }
    $ui.AboutPanel.Visibility = 'Collapsed'
    $ui.RunButton.Visibility = 'Visible'
    $ui.OpenFolderButton.Visibility = 'Visible'
    $ui.Progress.Visibility = 'Visible'
    Set-Status "$($task.Title). Add files, set the options, then press Run."
}

# ======================================================================
# RUNNERS
# ======================================================================
function Test-Ready {
    param([string]$Need)
    if ($Need -eq 'none') { return $true }
    if ($script:Files.Count -eq 0) { Show-Message 'Add at least one file first.'; return $false }
    if ($Need -eq 'one' -and $script:Files.Count -gt 1 -and $ui.FileList.SelectedItems.Count -ne 1) {
        Show-Message 'This task works on one file. Click the one you want in the list.'
        return $false
    }
    return $true
}

function Get-SingleFile {
    if ($ui.FileList.SelectedItems.Count -eq 1) { return $ui.FileList.SelectedItems[0] }
    return $script:Files[0]
}

function New-OutputFolder {
    $folder = $ui.OutputFolder.Text
    if (-not $folder) { $folder = Join-Path $script:Files[0].Folder (Get-DefaultOutputName) }
    New-Item -ItemType Directory -Force -Path $folder | Out-Null
    $script:LastOutput = $folder
    return $folder
}

function Invoke-HelperLogged {
    param([string[]]$Arguments)
    $result = Invoke-Helper -Arguments $Arguments -Quiet
    foreach ($line in $result.Lines) { Write-Log "   $line" }
    return $result
}

function Start-OcrTask {
    $ocr = Get-OcrCommand
    if (-not $ocr) { Show-Message 'OCRmyPDF was not found. Run Install PDFToolbox.bat.'; return }
    $files = @($script:Files)
    $outDir = New-OutputFolder
    $flags = New-Object System.Collections.Generic.List[string]
    switch ($ui.OcrMode.SelectedIndex) { 1 { $flags.Add('--skip-text') } 2 { $flags.Add('--force-ocr') } }
    if ($ui.OcrDeskew.IsChecked) { $flags.Add('--rotate-pages'); $flags.Add('--deskew') }
    $flags.Add('--optimize'); $flags.Add([string]$ui.OcrOptimize.SelectedIndex)
    $maxJobs = [int]$ui.OcrJobs.SelectedItem
    $cores = [Environment]::ProcessorCount
    $flags.Add('--jobs'); $flags.Add([string]([Math]::Max(1, [Math]::Floor($cores / $maxJobs))))

    $logFile = Join-Path $outDir 'OCR log.txt'
    $tmpDir = Expand-LongPath (Join-Path $script:WorkDir 'ocr')
    New-Item -ItemType Directory -Force -Path $tmpDir | Out-Null
    Write-RunLog $logFile ''
    Write-RunLog $logFile ('=== OCR run started {0} | {1} file(s) | options: {2} ===' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $files.Count, ($flags -join ' '))
    Write-RunLog $logFile ("    $script:Stamp | run by $env:USERNAME on $env:COMPUTERNAME")

    Write-Log "Running $maxJobs file(s) at a time."
    $jobs = New-Object System.Collections.ArrayList
    $stats = @{ OK = 0; Warn = 0; Failed = 0 }
    $started = Get-Date
    $index = 0
    foreach ($file in $files) {
        $index++
        while (@($jobs | Where-Object { -not $_.Done -and -not $_.Proc.HasExited }).Count -ge $maxJobs) {
            Start-Sleep -Milliseconds 400
            Update-OcrJobsUi -Jobs $jobs -LogFile $logFile -Stats $stats -Total $files.Count
        }
        Update-OcrJobsUi -Jobs $jobs -LogFile $logFile -Stats $stats -Total $files.Count
        $outPath = Join-Path $outDir ([System.IO.Path]::GetFileName($file.Path))
        if ($outPath -ieq $file.Path) {
            $outPath = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' OCR.pdf')
        }
        $id = [guid]::NewGuid().ToString('N')
        $errFile = Join-Path $tmpDir "$id.err.txt"
        $outFile = Join-Path $tmpDir "$id.out.txt"
        try {
            $argList = @($ocr.Pre) + @($flags) + @(('"{0}"' -f $file.Path), ('"{0}"' -f $outPath))
            $proc = Start-Process -FilePath $ocr.Exe -ArgumentList $argList -PassThru -NoNewWindow `
                -RedirectStandardError $errFile -RedirectStandardOutput $outFile
            $null = $proc.Handle
            [void]$jobs.Add([pscustomobject]@{ Proc = $proc; Name = $file.Name; Err = $errFile; Out = $outFile
                    In = $file.Path; OutPath = $outPath; Done = $false
                })
            Write-Log "[$index/$($files.Count)] Started: $($file.Name)"
        } catch {
            $stats.Failed++
            Write-Log "[$index/$($files.Count)] COULD NOT START: $($file.Name) ($($_.Exception.Message))"
            Write-RunLog $logFile "FAILED  $($file.Name) (could not start: $($_.Exception.Message))"
        }
    }
    while (@($jobs | Where-Object { -not $_.Done }).Count -gt 0) {
        Start-Sleep -Milliseconds 400
        Update-OcrJobsUi -Jobs $jobs -LogFile $logFile -Stats $stats -Total $files.Count
    }
    $elapsed = (Get-Date) - $started
    Write-RunLog $logFile ('=== Finished {0} | {1} OK | {2} to check | {3} failed ===' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $stats.OK, $stats.Warn, $stats.Failed)
    Write-Log ''
    Write-Log ('Finished in {0:hh\:mm\:ss}.  OK: {1}   To check: {2}   Failed: {3}' -f $elapsed, $stats.OK, $stats.Warn, $stats.Failed)
    if ($stats.Warn -gt 0) {
        Write-Log 'The "to check" files are complete and searchable. The PDF checker found a'
        Write-Log 'damaged image inside the original scan. Open those pages and look at them.'
    }
    Write-Log "Log: $logFile"
}

function Update-OcrJobsUi {
    param($Jobs, [string]$LogFile, [hashtable]$Stats, [int]$Total)
    foreach ($j in $Jobs) {
        if ($j.Done -or -not $j.Proc.HasExited) { continue }
        $j.Done = $true
        try {
            $code = $j.Proc.ExitCode
            if ($code -eq 0) {
                $Stats.OK++
                Write-Log "   Done:    $($j.Name)"
                Write-RunLog $LogFile "OK      $($j.Name)"
            } elseif ($code -eq 4) {
                $check = Test-OcrOutput -InPath $j.In -OutPath $j.OutPath
                if ($check.Usable) {
                    $Stats.Warn++
                    Write-Log "   Done:    $($j.Name)  (check this one)"
                    Write-RunLog $LogFile "WARN    $($j.Name) (damaged image in the original scan; $($check.Reason))"
                } else {
                    $Stats.Failed++
                    Write-Log "   FAILED:  $($j.Name) ($($check.Reason))"
                    Write-RunLog $LogFile "FAILED  $($j.Name) ($($check.Reason))"
                }
            } else {
                $Stats.Failed++
                $reason = Get-OcrExitText $code
                Write-Log "   FAILED:  $($j.Name) ($reason)"
                Write-RunLog $LogFile "FAILED  $($j.Name) ($reason)"
                try {
                    if (Test-Path -LiteralPath $j.Err) {
                        Get-Content -LiteralPath $j.Err -Tail 5 -ErrorAction SilentlyContinue | ForEach-Object {
                            if ($_.Trim()) { Write-RunLog $LogFile "        $_" }
                        }
                    }
                } catch { }
            }
        } catch {
            $Stats.Failed++
            Write-Log "   PROBLEM: $($j.Name) ($($_.Exception.Message))"
        }
        try { Remove-Item -LiteralPath $j.Err, $j.Out -Force -ErrorAction SilentlyContinue } catch { }
    }
    Set-Progress ($Stats.OK + $Stats.Warn + $Stats.Failed) $Total
}

function Start-SimpleTask {
    $files = @($script:Files)
    $task = $script:CurrentTask
    $done = 0
    $failed = 0

    switch ($task) {
        'merge' {
            if ($files.Count -lt 2) { Show-Message 'Add at least two PDFs to merge.'; return }
            $name = $ui.MergeName.Text.Trim()
            if (-not $name) { $name = 'Combined.pdf' }
            if ($name -notmatch '\.pdf$') { $name += '.pdf' }
            $outPath = Join-Path (New-OutputFolder) $name
            $list = New-ListFile @($files | ForEach-Object { $_.Path })
            Write-Log "Merging $($files.Count) files into $name"
            $r = Invoke-HelperLogged @('merge', $outPath, $list)
            Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
            if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            Set-Progress 1 1
        }
        'count' {
            $list = New-ListFile @($files | ForEach-Object { $_.Path })
            $r = Invoke-Helper -Arguments @('count', $list) -Quiet
            Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
            $total = 0
            foreach ($line in $r.Lines) {
                if ($line -notmatch "`t") { continue }
                $parts = $line -split "`t"
                if ($parts[1] -match '^\d+$') { $total += [int]$parts[1] }
                Write-Log "   $($parts[0])  -  $($parts[1]) page(s)"
            }
            Write-Log "   TOTAL: $total page(s) over $($files.Count) file(s)"
            $done = $files.Count
            Set-Progress 1 1
        }
        'search' {
            $term = $ui.SearchTerm.Text.Trim()
            if (-not $term) { Show-Message 'Type the text to search for.'; return }
            $list = New-ListFile @($files | ForEach-Object { $_.Path })
            $r = Invoke-Helper -Arguments @('search', $list, $term) -Quiet
            Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
            $hits = 0
            $noText = @()
            foreach ($line in $r.Lines) {
                if ($line -notmatch "`t") { continue }
                $parts = $line -split "`t"
                if ($parts[1] -match '^\d+$') { $hits++; Write-Log "   $($parts[0])  -  page $($parts[1])" }
                elseif ($parts[1] -eq 'NOTEXT') { $noText += $parts[0] }
            }
            Write-Log ''
            Write-Log "   $hits match(es)."
            if ($noText.Count -gt 0) {
                Write-Log "   $($noText.Count) file(s) have no searchable text and were not really searched:"
                foreach ($n in $noText) { Write-Log "     $n" }
                Write-Log '   OCR those first, then search again.'
            }
            $done = $files.Count
            Set-Progress 1 1
        }
        'compress' {
            if (-not (Get-GhostscriptExe)) { Show-Message 'Ghostscript was not found. Run Install PDFToolbox.bat.'; return }
            $outDir = New-OutputFolder
            $preset = @('/printer', '/ebook', '/screen')[[Math]::Max(0, $ui.CompressLevel.SelectedIndex)]
            $before = 0; $after = 0
            foreach ($file in $files) {
                $out = Join-Path $outDir ([System.IO.Path]::GetFileName($file.Path))
                if ($out -ieq $file.Path) { $out = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' compressed.pdf') }
                $r = Invoke-Ghostscript @('-sDEVICE=pdfwrite', '-dCompatibilityLevel=1.4', "-dPDFSETTINGS=$preset", "-sOutputFile=$out", $file.Path)
                if ($r.Code -eq 0 -and (Test-Path -LiteralPath $out)) {
                    $sizeBefore = (Get-Item -LiteralPath $file.Path).Length
                    $sizeAfter = (Get-Item -LiteralPath $out).Length
                    $before += $sizeBefore; $after += $sizeAfter; $done++
                    $pct = if ($sizeBefore -gt 0) { [Math]::Round(100 - ($sizeAfter / $sizeBefore * 100)) } else { 0 }
                    Write-Log ("   {0}: {1} -> {2} ({3}% smaller)" -f $file.Name, (Format-FileSize $sizeBefore), (Format-FileSize $sizeAfter), $pct)
                } else {
                    $failed++
                    Write-Log "   FAILED: $($file.Name)"
                }
                Set-Progress ($done + $failed) $files.Count
            }
            if ($done -gt 0) {
                $pct = if ($before -gt 0) { [Math]::Round(100 - ($after / $before * 100)) } else { 0 }
                Write-Log ("   Total: {0} -> {1} ({2}% smaller)" -f (Format-FileSize $before), (Format-FileSize $after), $pct)
            }
        }
        'extract' {
            $file = Get-SingleFile
            $spec = $ui.PagesSpec.Text.Trim()
            if (-not $spec) { Show-Message 'Enter the pages, or pick them from previews.'; return }
            $outPath = Join-Path (New-OutputFolder) ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' extract.pdf')
            $r = Invoke-HelperLogged @('extract', $file.Path, $outPath, $spec)
            if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            Set-Progress 1 1
        }
        'delete' {
            $file = Get-SingleFile
            $spec = $ui.PagesSpec.Text.Trim()
            if (-not $spec) { Show-Message 'Enter the pages, or pick them from previews.'; return }
            $outPath = Join-Path (New-OutputFolder) ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' trimmed.pdf')
            $r = Invoke-HelperLogged @('delete', $file.Path, $outPath, $spec)
            if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            Set-Progress 1 1
        }
        'rotate' {
            $file = Get-SingleFile
            $spec = $ui.PagesSpec.Text.Trim()
            if (-not $spec) { $spec = 'all' }
            $degrees = @(90, 180, 270)[[Math]::Max(0, $ui.RotateDegrees.SelectedIndex)]
            $outPath = Join-Path (New-OutputFolder) ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' rotated.pdf')
            $r = Invoke-HelperLogged @('rotate', $file.Path, $outPath, $spec, [string]$degrees)
            if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            Set-Progress 1 1
        }
        'split' {
            $file = Get-SingleFile
            $outDir = Join-Path (New-OutputFolder) ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' pages')
            $r = Invoke-HelperLogged @('split', $file.Path, $outDir)
            if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            Set-Progress 1 1
        }
        'rearrange' {
            $file = Get-SingleFile
            $order = Get-PageOrderVisually -PdfPath $file.Path
            if (-not $order) { Write-Log 'Nothing arranged.'; return }
            $outPath = Join-Path (New-OutputFolder) ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' reordered.pdf')
            $r = Invoke-HelperLogged @('reorder', $file.Path, $outPath, ($order -join ','))
            if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            Set-Progress 1 1
        }
        'insert' {
            $file = Get-SingleFile
            $source = $ui.InsertSource.Text
            if (-not $source -or -not (Test-Path -LiteralPath $source)) { Show-Message 'Choose the PDF or image to insert.'; return }
            $position = switch ($ui.InsertPosition.SelectedIndex) {
                0 { 'start' }
                1 { 'end' }
                default {
                    $before = $ui.InsertBefore.Text.Trim()
                    if ($before -notmatch '^\d+$') { Show-Message 'Type the page number to insert before.'; return }
                    $before
                }
            }
            $outPath = Join-Path (New-OutputFolder) ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' updated.pdf')
            $r = Invoke-HelperLogged @('insert', $file.Path, $outPath, $source, 'all', $position)
            if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            Set-Progress 1 1
        }
        'password' {
            $file = Get-SingleFile
            $env:PDFTB_PASSWORD = $ui.PasswordBox.Password
            try {
                $outPath = Join-Path (New-OutputFolder) ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + ' unlocked.pdf')
                $r = Invoke-HelperLogged @('unlock', $file.Path, $outPath)
                if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
            } finally { Clear-PdfPassword }
            Set-Progress 1 1
        }
        'convert' {
            $outDir = New-OutputFolder
            switch ($ui.ConvertKind.SelectedIndex) {
                0 {
                    # Office to PDF
                    $stats = @{ OK = 0; Failed = 0 }
                    $used = @{}
                    $all = @($files | ForEach-Object { Get-Item -LiteralPath $_.Path })
                    Convert-OfficeFiles -App 'Word' -Files @($all | Where-Object { @('.doc', '.docx') -contains $_.Extension.ToLower() }) -OutDir $outDir -Stats $stats -Used $used
                    Convert-OfficeFiles -App 'Excel' -Files @($all | Where-Object { @('.xls', '.xlsx', '.xlsm') -contains $_.Extension.ToLower() }) -OutDir $outDir -Stats $stats -Used $used
                    Convert-OfficeFiles -App 'PowerPoint' -Files @($all | Where-Object { @('.ppt', '.pptx') -contains $_.Extension.ToLower() }) -OutDir $outDir -Stats $stats -Used $used
                    $done = $stats.OK; $failed = $stats.Failed
                    Set-Progress 1 1
                }
                1 {
                    # PDF to Word
                    try { $word = New-Object -ComObject Word.Application; $word.Visible = $false; $word.DisplayAlerts = 0 }
                    catch { Show-Message 'Word is not installed or could not start.'; return }
                    try {
                        foreach ($file in $files) {
                            $out = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + '.docx')
                            $doc = $null
                            try {
                                $doc = $word.Documents.Open($file.Path, $false, $true)
                                $doc.SaveAs2($out, 16)
                                $doc.Close(0)
                                $done++
                                Write-Log "   Converted: $($file.Name)"
                            } catch {
                                $failed++
                                Write-Log "   FAILED: $($file.Name) ($($_.Exception.Message))"
                                if ($doc) { try { $doc.Close(0) } catch { } }
                            } finally {
                                if ($doc) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
                            }
                            Set-Progress ($done + $failed) $files.Count
                        }
                    } finally {
                        try { $word.Quit() } catch { }
                        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($word)
                        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
                    }
                    Write-Log 'Check the result against the PDF. Layout is rebuilt, not copied.'
                }
                2 {
                    # PDF tables to Excel
                    foreach ($file in $files) {
                        $base = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($file.Path))
                        Write-Log "   $($file.Name)"
                        $r = Invoke-HelperLogged @('tables', $file.Path, $base)
                        if ($r.Code -eq 0) { $done++ } else { $failed++ }
                        Set-Progress ($done + $failed) $files.Count
                    }
                    Write-Log 'Table detection is a best guess. Tie the totals back to the PDF.'
                }
                3 {
                    # PDF to text
                    foreach ($file in $files) {
                        $out = Join-Path $outDir ([System.IO.Path]::GetFileNameWithoutExtension($file.Path) + '.txt')
                        Write-Log "   $($file.Name)"
                        $r = Invoke-HelperLogged @('text', $file.Path, $out)
                        if ($r.Code -eq 0) { $done++ } else { $failed++ }
                        Set-Progress ($done + $failed) $files.Count
                    }
                }
                4 {
                    # PDF to images
                    if (-not (Get-GhostscriptExe)) { Show-Message 'Ghostscript was not found.'; return }
                    foreach ($file in $files) {
                        $stem = [System.IO.Path]::GetFileNameWithoutExtension($file.Path)
                        $dir = Join-Path $outDir "$stem images"
                        New-Item -ItemType Directory -Force -Path $dir | Out-Null
                        $pattern = Join-Path $dir "$stem p%03d.png"
                        $r = Invoke-Ghostscript @('-sDEVICE=png16m', '-r150', "-sOutputFile=$pattern", $file.Path)
                        $made = @(Get-ChildItem -LiteralPath $dir -Filter '*.png' -File -ErrorAction SilentlyContinue)
                        if ($r.Code -eq 0 -and $made.Count -gt 0) { $done++; Write-Log "   $($file.Name): $($made.Count) image(s)" }
                        else { $failed++; Write-Log "   FAILED: $($file.Name)" }
                        Set-Progress ($done + $failed) $files.Count
                    }
                }
                5 {
                    # images to one PDF
                    $images = @($files | Where-Object { $script:ImageExtensions -contains ([System.IO.Path]::GetExtension($_.Path)).ToLower() })
                    if ($images.Count -eq 0) { Show-Message 'Add image files for this conversion.'; return }
                    $outPath = Join-Path $outDir 'Scanned.pdf'
                    $list = New-ListFile @($images | ForEach-Object { $_.Path })
                    $r = Invoke-HelperLogged @('imgtopdf', $list, $outPath)
                    Remove-Item -LiteralPath $list -ErrorAction SilentlyContinue
                    if ($r.Code -eq 0) { $done = 1 } else { $failed = 1 }
                    Set-Progress 1 1
                }
            }
        }
    }

    if ($done -gt 0 -or $failed -gt 0) {
        Write-Log ''
        Write-Log "Done: $done   Failed: $failed"
    }
}

# ======================================================================
# EVENTS
# ======================================================================
$ui.VersionText.Text = "v$script:Version"
$ui.AboutLine1.Text = "Version $script:Version, released $script:ReleaseDate"
$ui.AboutLine2.Text = "Built by $script:Author"
$ui.AboutMailText.Text = $script:AuthorEmail
# Subject is filled in with the version, so a report always carries it.
$ui.AboutMailLink.NavigateUri = New-Object System.Uri (
    'mailto:{0}?subject={1}' -f $script:AuthorEmail,
    [System.Uri]::EscapeDataString("PDF Toolbox v$script:Version - "))
$ui.AboutMailLink.Add_RequestNavigate({
        param($sender, $e)
        try { Start-Process $e.Uri.AbsoluteUri }
        catch { Show-Message "Could not open your email app. The address is $script:AuthorEmail" }
        $e.Handled = $true
    })
$ui.AboutUpdateButton.Add_Click({
        $ui.AboutUpdateButton.IsEnabled = $false
        $ui.AboutDownloadLine.Visibility = 'Collapsed'
        $ui.AboutUpdateStatus.Foreground = [System.Windows.Media.Brushes]::Goldenrod
        $ui.AboutUpdateStatus.Text = 'Checking...'
        Update-Ui
        $info = Get-LatestVersionInfo
        if (-not $info.Ok) {
            $ui.AboutUpdateStatus.Foreground = [System.Windows.Media.Brushes]::IndianRed
            $ui.AboutUpdateStatus.Text = "Could not check: $($info.Message)"
        } elseif ($info.Newer) {
            $ui.AboutUpdateStatus.Foreground = [System.Windows.Media.Brushes]::Goldenrod
            $notes = $(if ($info.Notes) { "  $($info.Notes)" } else { '' })
            $ui.AboutUpdateStatus.Text = "Version $($info.Latest) is available. You are on $script:Version.$notes"
            try { $ui.AboutDownloadLink.NavigateUri = New-Object System.Uri $info.Download } catch { }
            $ui.AboutDownloadLine.Visibility = 'Visible'
        } else {
            $ui.AboutUpdateStatus.Foreground = [System.Windows.Media.Brushes]::MediumSeaGreen
            $ui.AboutUpdateStatus.Text = "You are on the latest version ($script:Version)."
        }
        $ui.AboutUpdateButton.IsEnabled = $true
    })

$ui.AboutDownloadLink.Add_RequestNavigate({
        param($sender, $e)
        try { Start-Process $e.Uri.AbsoluteUri } catch { Show-Message "Could not open the browser. Go to $script:UpdateDownloadUrl" }
        $e.Handled = $true
    })

$ui.AboutContact.Text = "Please say which version you are on (this is $script:Version, shown at the top of this page and in the title bar) and what you were doing at the time."

$ui.SubTitle.Text = "Built by $script:Author"
$window.Title = "PDF Toolbox v$script:Version"

foreach ($task in $script:Tasks) { [void]$ui.TaskList.Items.Add($task.Title) }
$ui.OcrMode.Items.Add('Fully scanned (fastest)') | Out-Null
$ui.OcrMode.Items.Add('Mixed: skip pages with text') | Out-Null
$ui.OcrMode.Items.Add('Force OCR: replace text layer') | Out-Null
$ui.OcrMode.SelectedIndex = 0
foreach ($label in @('None', 'Lossless', 'Balanced')) { $ui.OcrOptimize.Items.Add($label) | Out-Null }
$ui.OcrOptimize.SelectedIndex = 1
foreach ($n in 1..8) { $ui.OcrJobs.Items.Add($n) | Out-Null }
$ui.OcrJobs.SelectedIndex = $(if ([Environment]::ProcessorCount -le 4) { 1 } else { 2 })
foreach ($label in @('90 degrees clockwise', '180 degrees', '90 degrees anticlockwise')) { $ui.RotateDegrees.Items.Add($label) | Out-Null }
$ui.RotateDegrees.SelectedIndex = 0
foreach ($label in @('At the start', 'At the end', 'Before a page')) { $ui.InsertPosition.Items.Add($label) | Out-Null }
$ui.InsertPosition.SelectedIndex = 1
foreach ($label in @('Word / Excel / PowerPoint  ->  PDF', 'PDF  ->  Word', 'PDF  ->  Excel (tables)',
        'PDF  ->  plain text', 'PDF  ->  images', 'Images  ->  one PDF')) { $ui.ConvertKind.Items.Add($label) | Out-Null }
$ui.ConvertKind.SelectedIndex = 0
foreach ($label in @('Light (300 dpi)', 'Balanced (150 dpi)', 'Maximum (72 dpi)')) { $ui.CompressLevel.Items.Add($label) | Out-Null }
$ui.CompressLevel.SelectedIndex = 1

$ui.TaskList.Add_SelectionChanged({
        $i = $ui.TaskList.SelectedIndex
        if ($i -ge 0) { Select-Task $script:Tasks[$i].Key }
    })

$ui.ConvertKind.Add_SelectionChanged({
        $hints = @('Needs Word, Excel or PowerPoint installed.',
            'Word rebuilds the layout. Check the result.',
            'Only works on PDFs with real text. OCR scans first.',
            'Extracts the text layer only.',
            'One PNG per page, 150 dpi.',
            'Add image files to the list for this one.')
        $ui.ConvertHint.Text = $hints[[Math]::Max(0, $ui.ConvertKind.SelectedIndex)]
    })

$ui.AddFilesButton.Add_Click({
        $picked = Show-FilePicker -Title 'Add files' -Extensions (Get-TaskExtensions) -Multiple
        if ($picked) { Add-FilePaths -Paths @($picked) -Extensions (Get-TaskExtensions) }
    })

$ui.AddFolderButton.Add_Click({
        $picked = Show-FolderPicker -Title 'Add every file in a folder'
        if ($picked) { Add-FilePaths -Paths @($picked) -Extensions (Get-TaskExtensions) }
    })

$ui.RemoveButton.Add_Click({
        foreach ($item in @($ui.FileList.SelectedItems)) { $script:Files.Remove($item) }
        Set-Status "$($script:Files.Count) file(s) in the list."
    })

$ui.ClearButton.Add_Click({ $script:Files.Clear(); Set-Status 'List cleared.' })

$ui.MoveUpButton.Add_Click({
        $i = $ui.FileList.SelectedIndex
        if ($i -gt 0) { $script:Files.Move($i, $i - 1); $ui.FileList.SelectedIndex = $i - 1 }
    })

$ui.MoveDownButton.Add_Click({
        $i = $ui.FileList.SelectedIndex
        if ($i -ge 0 -and $i -lt $script:Files.Count - 1) { $script:Files.Move($i, $i + 1); $ui.FileList.SelectedIndex = $i + 1 }
    })

# Drag a row to reorder it, or drag files in from Explorer to add them.
# This scriptblock is NOT closured on purpose: it keeps this script's scope,
# so Add-FilePaths and $ui still resolve when it runs.
Enable-DragReorder -Control $ui.FileList -Collection $script:Files -OnExternalFiles {
    param($paths)
    Add-FilePaths -Paths @($paths) -Extensions (Get-TaskExtensions)
}

function Show-SelectedFilePreview {
    if ($script:Files.Count -eq 0) { Show-Message 'Add a file first.'; return }
    $row = Get-SingleFile
    if ([System.IO.Path]::GetExtension($row.Path).ToLower() -ne '.pdf') {
        Show-Message 'Preview only works on PDFs. Convert the file first.'
        return
    }
    $total = $(if ($row.Pages -match '^\d+$') { [int]$row.Pages } else { 0 })
    Show-PagePreview -PdfPath $row.Path -Page 1 -Title $row.Name -TotalPages $total
}

$ui.PreviewFileButton.Add_Click({ Show-SelectedFilePreview })
$ui.FileList.Add_MouseDoubleClick({ Show-SelectedFilePreview })

$ui.PickPagesButton.Add_Click({
        if ($script:Files.Count -eq 0) { Show-Message 'Add a PDF first.'; return }
        $file = Get-SingleFile
        $spec = Select-PagesVisually -PdfPath $file.Path -Instruction 'Tick the pages you want. Ctrl+click for more than one, Shift+click for a run.'
        if ($spec) { $ui.PagesSpec.Text = $spec }
    })

$ui.MergeArrangeButton.Add_Click({
        if ($script:Files.Count -lt 2) { Show-Message 'Add at least two PDFs.'; return }
        $items = @($script:Files | ForEach-Object { Get-Item -LiteralPath $_.Path })
        $arranged = Get-FileOrderVisually -Files $items
        if (-not $arranged) { return }
        $ordered = @()
        foreach ($file in $arranged) {
            $match = @($script:Files | Where-Object { $_.Path -eq $file.FullName })
            if ($match.Count -gt 0) { $ordered += $match[0] }
        }
        $script:Files.Clear()
        foreach ($row in $ordered) { $script:Files.Add($row) }
    })

$ui.InsertBrowseButton.Add_Click({
        $picked = Show-FilePicker -Title 'Choose the PDF or image to insert' -Extensions (@('.pdf') + $script:ImageExtensions)
        if ($picked) { $ui.InsertSource.Text = $picked }
    })

$ui.OutputBrowseButton.Add_Click({
        $picked = Show-FolderPicker -Title 'Choose where to save the results'
        if ($picked) { $ui.OutputFolder.Text = $picked }
    })

$ui.OpenFolderButton.Add_Click({
        $folder = if ($script:LastOutput) { $script:LastOutput } else { $ui.OutputFolder.Text }
        if ($folder -and (Test-Path -LiteralPath $folder)) { Invoke-Item -LiteralPath $folder }
        else { Show-Message 'Nothing has been saved yet.' }
    })

$ui.RunButton.Add_Click({
        $task = Get-Task $script:CurrentTask
        if (-not $task) { Show-Message 'Choose a task on the left.'; return }
        if (-not (Test-Ready $task.Files)) { return }
        $ui.RunButton.IsEnabled = $false
        Clear-Log
        Set-Progress 0 1
        Set-Status "$($task.Title): working..." -Refresh
        $started = Get-Date
        try {
            if ($script:CurrentTask -eq 'ocr') { Start-OcrTask } else { Start-SimpleTask }
            Set-Status ("Finished in {0:mm\:ss}." -f ((Get-Date) - $started))
        } catch {
            Write-Log ''
            Write-Log "Something went wrong: $($_.Exception.Message)"
            Set-Status 'Stopped with an error.'
        } finally {
            Clear-PdfPassword
            $ui.RunButton.IsEnabled = $true
        }
    })

$ui.SetupRefreshButton.Add_Click({ Update-SessionPath; Show-SetupPage })
$ui.SetupInstallButton.Add_Click({ Start-Installer })
$ui.SetupUpdateButton.Add_Click({ Get-UpdateInfo })
$ui.SetupUpdateAllButton.Add_Click({ Update-AllComponents })

$ui.TaskList.SelectedIndex = 0
Write-Log "PDF Toolbox v$script:Version. Pick a task on the left, add files, then press Run."
[void]$window.ShowDialog()
