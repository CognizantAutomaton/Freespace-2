#region attempt to retrieve FS2 root folder from registry
[string]$InstallLocation = [string]::Empty

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 273620") {
    # steam
    $InstallLocation = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 273620").InstallLocation
} else {
    # legacy
    [Array]$hkpaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FreeSpace2",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\FreeSpace"
    )

    for ([int]$n = 0; ($InstallLocation.Length -eq 0) -and ($n -lt $hkpaths.Count); $n++) {
        [string]$hkpath = $hkpaths[$n]

        if (Test-Path $hkpath) {
            (Get-ItemProperty -Path $hkpath).UninstallString.Split(" ") | Where-Object {
                $_.Length -gt 0
            } | ForEach-Object {
                [System.IO.Path]::GetDirectoryName($_.TrimStart("-f").TrimStart("-c"))
            } | Select-Object -Skip 1 | ForEach-Object {
                if (Test-Path -LiteralPath $_) {
                    $InstallLocation = $_
                }
            }
        }
    }
}
#endregion

#region dependencies
Add-Type -AssemblyName PresentationCore, PresentationFramework
Add-Type -AssemblyName System.Windows.Forms
Add-Type -TypeDefinition @"
using System;
using System.Windows.Data;

namespace MyConverter {
    [ValueConversion(typeof(object), typeof(string))]
    public class ConvertLongToFilesize : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            double num = double.Parse(value.ToString());
            string[] units = new string[] { "B", "KiB", "MiB", "GiB", "TiB", "PiB", "YiB" };
            int unitsIdx = 0;

            while (num > 1024) {
                num /= 1024;
                unitsIdx++;
            }

            return Math.Round(num, 2) + " " + units[unitsIdx];
        }

        public object ConvertBack(object value, Type targetType,
            object parameter, System.Globalization.CultureInfo culture) {
            // don't intend this to ever be called
            return null;
        }
    }

    [ValueConversion(typeof(UInt32), typeof(string))]
    public class ConvertTimestampToDate : IValueConverter {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture) {
            DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            dtDateTime = dtDateTime.AddSeconds((UInt32)value).ToLocalTime();
            return dtDateTime;
        }

        public object ConvertBack(object value, Type targetType,
            object parameter, System.Globalization.CultureInfo culture) {
            // don't intend this to ever be called
            return null;
        }
    }
}
"@ -ReferencedAssemblies PresentationFramework
Add-Type @"
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

// observable collection viewmodel
public class FileViewModel : INotifyPropertyChanged {
    public event PropertyChangedEventHandler PropertyChanged;
    private UInt32 _size;
    private UInt32 _timestamp;
    private string _filename;
    private string _path;

    public string Path {
        get { return _path; }
        set {
            _path = value;
            OnPropertyChanged();
        }
    }
    public UInt32 Offset { get; set; }
    public string Filename {
        get { return _filename; }
        set {
            _filename = value;
            OnPropertyChanged();
        }
    }
    public UInt32 Size {
        get { return _size; }
        set {
            _size = value;
            OnPropertyChanged();
        }
    }
    public UInt32 Timestamp {
        get {
            //DateTime dtDateTime = new DateTime(1970, 1, 1, 0, 0, 0, 0, System.DateTimeKind.Utc);
            //dtDateTime = dtDateTime.AddSeconds(_timestamp).ToLocalTime();
            return _timestamp;
        }
        set {
            //UInt32 unixTimestamp = value; //(UInt32)(value.Subtract(new DateTime(1970, 1, 1))).TotalSeconds;
            _timestamp = value;
            OnPropertyChanged();
        }
    }
    public int SeekType { get; set; }
    public string ExternalVP { get; set; }

    public void OnPropertyChanged([CallerMemberName]string caller = null) {
        var handler = PropertyChanged;
        if (handler != null) {
            handler(this, new PropertyChangedEventArgs(caller));
        }
    }
}
"@
#endregion

#region classes
enum SeekTypes {
    Present = 0
    FileSystem = 1
    External = 2
}

class WAVPlayer {
    $Runspace = $null
    $Session = $null
    $Handle = $null
    $Sync = $null

    WAVPlayer($Runspace, $Sync) {
        $this.Runspace = $Runspace
        $this.Sync = $Sync
    }

    [void]Play() {
        # add a script to run in the other thread
        $this.Session = [PowerShell]::Create().AddScript({
            $Sync.Window.Dispatcher.Invoke([Action]{
                $script:item = $Sync.Gui.gridFiles.SelectedItem
            })

            if ($item.SeekType -eq 0) {
                $buffer = [byte[]]::CreateInstance([byte], $item.Size)
                $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($VP_Filename, "Open")))
                [void]$reader.BaseStream.Seek($item.Offset, "Begin")
                [void]$reader.Read($buffer, 0, $item.Size)
                $reader.Close()

                [System.IO.Stream]$stream = New-Object System.IO.MemoryStream($buffer, 0, $buffer.Count)
                $Sync.Player = New-Object System.Media.SoundPlayer($stream)
                $Sync.Player.Play()
                $stream.Close()
            }
        }, $true)
        $this.Runspace.SessionStateProxy.SetVariable("VP_Filename", $global:VP_Filename)

        # invoke the runspace session created above
        $this.Session.Runspace = $this.Runspace
        $this.Handle = $this.Session.BeginInvoke()
    }

    [void]Stop() {
        if ($null -ne $this.Session) {
            if ($null -ne $this.Sync.Player) {
                $this.Sync.Player.Stop()
            }
            $this.Session.EndInvoke($this.Handle)
        }
    }
}
#endregion

#region runspace
# init synchronized hashtable
$Sync = [HashTable]::Synchronized(@{})

# init runspace
$Runspace = [RunspaceFactory]::CreateRunspace()
$Runspace.ApartmentState = [Threading.ApartmentState]::STA
$Runspace.ThreadOptions = "ReuseThread"
$Runspace.Open()

# provide the other thread with the synchronized hashtable (variable shared across threads)
$Runspace.SessionStateProxy.SetVariable("Sync", $Sync)

$player = New-Object WAVPlayer($Runspace, $Sync)
#endregion

#region GUI
# utilizing a converter requires that you pull the rather ugly assembly name
$conv = New-Object MyConverter.ConvertLongToFilesize
$ConverterAssemblyName = $conv.GetType().Assembly.FullName.Split(',')[0]

[Xml]$WpfXml = @"
<Window x:Name="VPViewer" x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        xmlns:Converter="clr-namespace:MyConverter;assembly=ConverterAssemblyName"
        mc:Ignorable="d"
        Title="VP Compiler" WindowStartupLocation="CenterScreen" Visibility="Visible" Height="600" Width="700">
    <Window.Resources>
        <Converter:ConvertLongToFilesize x:Key="Convert2Filesize"/>
        <Converter:ConvertTimestampToDate x:Key="ConvertTimestampToDate"/>
    </Window.Resources>
    <DockPanel>
        <Menu DockPanel.Dock="Top">
            <MenuItem Header="_File">
                <MenuItem
                    x:Name="mnuNew"
                    Header="_New"/>
                <MenuItem
                    x:Name="mnuOpen"
                    Header="_Open"/>
                <MenuItem
                    x:Name="mnuSave"
                    Header="_Save"/>
                <MenuItem
                    x:Name="mnuSaveAs"
                    Header="Save _As"/>
                <Separator />
                <MenuItem
                    x:Name="mnuExtractAll"
                    Header="_Extract All"/>
            </MenuItem>
        </Menu>
        <Grid>
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="100*" MinWidth="100"/>
                <ColumnDefinition Width="5"/>
                <ColumnDefinition Width="350*" MinWidth="250"/>
            </Grid.ColumnDefinitions>
            <TreeView
                Grid.Column="0"
                x:Name="treeVPP"
                HorizontalAlignment="Stretch">
                <TreeView.ContextMenu>
                    <ContextMenu>
                        <MenuItem
                            x:Name="mnuAddFolder"
                            Header="_Add Folder"/>
                        <MenuItem
                            x:Name="mnuRemoveFolder"
                            Header="_Remove Folder"/>
                        <MenuItem
                            x:Name="mnuRenameFolder"
                            Header="Re_name Folder"/>
                        <Separator/>
                        <MenuItem
                            x:Name="mnuAddFiles"
                            Header="Add _Files"/>
                    </ContextMenu>
                </TreeView.ContextMenu>
            </TreeView>
            <GridSplitter
                Grid.Column="1"
                Width="5"
                HorizontalAlignment="Stretch"/>
            <DataGrid
                Grid.Column="2"
                x:Name="gridFiles"
                Height="Auto"
                HorizontalAlignment="Stretch"
                CanUserReorderColumns="False"
                CanUserResizeColumns="True"
                CanUserResizeRows="False"
                CanUserSortColumns="True"
                CanUserAddRows="False"
                IsReadOnly="True"
                AutoGenerateColumns="False"
                SelectionMode="Extended"
                SelectionUnit="FullRow"
                VerticalAlignment="Stretch">
                <DataGrid.Resources>
                    <Style TargetType="{x:Type DataGridCell}">
                        <Setter Property="BorderThickness" Value="0" />
                        <Setter Property="FocusVisualStyle" Value="{x:Null}" />
                    </Style>
                </DataGrid.Resources>
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Filename" Binding="{Binding Filename}" Width="*"/>
                    <DataGridTextColumn Header="Size" Binding="{Binding Size, Converter={StaticResource Convert2Filesize}}" Width="*"/>
                    <DataGridTextColumn Header="Path" Binding="{Binding Path}" Width="*"/>
                    <DataGridTextColumn Header="Date and Time" Binding="{Binding Timestamp, Converter={StaticResource ConvertTimestampToDate}}" Width="*"/>
                </DataGrid.Columns>
                <DataGrid.ContextMenu>
                    <ContextMenu>
                        <MenuItem x:Name="mnuExtract" Header="Extract"/>
                        <MenuItem x:Name="mnuOpenFile" Header="Open"/>
                        <MenuItem x:Name="mnuRemoveFiles" Header="Remove"/>
                    </ContextMenu>
                </DataGrid.ContextMenu>
            </DataGrid>
        </Grid>
    </DockPanel>
</Window>
"@ -replace "ConverterAssemblyName", $ConverterAssemblyName

# these attributes can disturb powershell's ability to load XAML, so remove them
$WpfXml.Window.RemoveAttribute('x:Class')
$WpfXml.Window.RemoveAttribute('mc:Ignorable')

# add namespaces for later use if needed
$WpfNs = New-Object -TypeName Xml.XmlNamespaceManager -ArgumentList $WpfXml.NameTable
$WpfNs.AddNamespace('x', $WpfXml.DocumentElement.x)
$WpfNs.AddNamespace('d', $WpfXml.DocumentElement.d)
$WpfNs.AddNamespace('mc', $WpfXml.DocumentElement.mc)

$Sync.Gui = @{}

# Read XAML markup
try {
    $Sync.Window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $WpfXml))
} catch {
    Write-Host $_ -ForegroundColor Red
    Exit
}

#===================================================
# Retrieve a list of all GUI elements
#===================================================
$WpfXml.SelectNodes('//*[@x:Name]', $WpfNs) | ForEach-Object {
    $Sync.Gui.Add($_.Name, $Sync.Window.FindName($_.Name))
}

# bind observable collection to datagrid
$Sync.ViewFiles = New-Object System.Collections.ObjectModel.ObservableCollection[FileViewModel]
$Sync.Gui.gridFiles.ItemsSource = $Sync.ViewFiles
#endregion

#region variables and functions
$global:VP_Filename = [string]::Empty
$global:Flag_Changed = $false

[ScriptBlock]$ConvertTimestamp = {
    param([UInt32]$timestamp)

    [DateTime]$dtDateTime = New-Object DateTime(1970, 1, 1, 0, 0, 0, 0, [System.DateTimeKind]::Utc)
    $dtDateTime = $dtDateTime.AddSeconds($timestamp).ToLocalTime()

    return $dtDateTime
}

[ScriptBlock]$SaveCommand = {
    param([string]$filename)

    [string]$tmp_filename = Join-Path ([IO.Path]::GetDirectoryName($filename)) "$((New-Guid).Guid).tmp"

    [UInt32]$position = 16
    [UInt32]$diroffset = $position + ($Sync.ViewFiles | ForEach-Object { $_.Size } | Measure-Object -Sum).Sum
    [UInt32]$index_position = $diroffset
    [UInt32]$totalbytes = $diroffset
    [UInt32]$direntries = $Sync.ViewFiles.Count

    # create a stack
    $stack = New-Object 'System.Collections.Generic.Stack[System.Windows.Controls.TreeViewItem]'
    $stack.Push($Sync.Gui.treeVPP.Items[0])

    # count remaining direntries
    while ($stack.Count -gt 0) {
        $dir = $stack.Pop()
        $direntries += 2

        for ([int]$n = $dir.Items.Count - 1; $n -ge 0; $n--) {
            $stack.Push($dir.Items[$n])
        }
    }

    $stack.Clear()

    # compute total VP bytes to write
    $totalbytes += $direntries * 44
    $writer = $null
    $stream = $null
    $CurrentReader = $null


    # instantiate VP file bytes
    [byte[]]@([byte]0) | Set-Content -LiteralPath $tmp_filename -AsByteStream

    try {
        $stream = New-Object IO.FileStream($tmp_filename, "Open", "ReadWrite")
        $writer = New-Object IO.StreamWriter($stream)
        $CurrentReader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($global:VP_Filename, "Open")))

        [void]$writer.BaseStream.SetLength($totalbytes)
        [void]$writer.BaseStream.Seek(0, "Begin")

        # write header
        [void]$writer.BaseStream.Write([System.Text.Encoding]::ASCII.GetBytes("VPVP"), 0, 4)
        [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes(2), 0, 4)
        [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($diroffset), 0, 4)
        [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($direntries), 0, 4)

        $stack.Push($Sync.Gui.treeVPP.Items[0])
        # a list of directories flagged as visited
        $visited = New-Object System.Collections.ArrayList

        # write files and file index
        while ($stack.Count -gt 0) {
            $dir = $stack.Peek()

            if ($dir -notin $visited) {
                [void]$writer.BaseStream.Seek($index_position, "Begin")

                # flag directory as visited
                $visited.Add($dir)

                # write directory index info
                [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($position), 0, 4)
                [void]$writer.BaseStream.Seek($index_position + 8, "Begin")
                [void]$writer.BaseStream.Write([System.Text.Encoding]::ASCII.GetBytes($dir.Header), 0, $dir.Header.Length)
                $index_position += 44
            }

            for ([int]$n = $dir.Items.Count - 1; $n -ge 0; $n--) {
                if ($dir.Items[$n] -notin $visited) {
                    $stack.Push($dir.Items[$n])
                }
            }

            if ($stack.Peek() -eq $dir) {
                # retrieve all files under current directory tag
                $files = $Sync.ViewFiles | Where-Object { $_.Path -eq $dir.Tag.ToString() } | Sort-Object -Property Filename
                [byte[]]$buffer = $null

                foreach ($file in $files) {
                    [void]$writer.BaseStream.Seek($index_position, "Begin")

                    switch ($file.SeekType) {
                        ([int][SeekTypes]::Present) {
                            # handle files already present
                            $buffer = [byte[]]::CreateInstance([byte], $file.Size)
                            [void]$CurrentReader.BaseStream.Seek($file.Offset, "Begin")
                            [void]$CurrentReader.Read($buffer, 0, $file.Size)
                        }

                        ([int][SeekTypes]::FileSystem) {
                            # handle files added from filesystem
                            if (Test-Path -LiteralPath $file.Filename) {
                                $buffer = [System.IO.File]::ReadAllBytes($file.Filename)
                                $file.Filename = [IO.Path]::GetFileName($file.Filename)
                            }
                        }

                        ([int][SeekTypes]::External) {
                            # handle files from an external VP
                            if (Test-Path -LiteralPath $file.ExternalVP) {
                                $reader = $null

                                try {
                                    $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($file.ExternalVP, "Open")))
                                    [void]$reader.BaseStream.Seek($file.Offset, "Begin")
                                    $buffer = [byte[]]::CreateInstance([byte], $file.Size)
                                    [void]$reader.Read($buffer, 0, $file.Size)
                                } catch {
                                    $buffer = [byte[]]::CreateInstance([byte], $file.Size)
                                    [Array]::Clear($buffer, 0, $file.Size)
                                }

                                if ($null -ne $reader) {
                                    $reader.Close()
                                }
                            } elseif ($file.ExternalVP.Trim().Length -eq 0) {
                                [System.Windows.MessageBox]::Show("You did not set the 'ExternalVP' field for '$($file.Filename)'")
                            }
                        }
                    }

                    # write file index info
                    [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($position), 0, 4)
                    [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($file.Size), 0, 4)
                    [void]$writer.BaseStream.Write([System.Text.Encoding]::ASCII.GetBytes($file.Filename), 0, $file.Filename.Length)

                    [void]$writer.BaseStream.Seek($index_position + 40, "Begin")
                    [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($file.Timestamp), 0, 4)
                    $index_position += 44

                    # write file bytes
                    [void]$writer.BaseStream.Seek($position, "Begin")
                    [void]$writer.BaseStream.Write($buffer, 0, $buffer.Count)

                    $position += $file.Size
                }

                # write backdir index info
                [void]$writer.BaseStream.Seek($index_position, "Begin")
                [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($position), 0, 4)
                [void]$writer.BaseStream.Seek($index_position + 8, "Begin")
                [void]$writer.BaseStream.Write([System.Text.Encoding]::ASCII.GetBytes(".."), 0, 2)
                $index_position += 44

                [void]$stack.Pop()
            }
        }
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
        Write-Host "expression: $($_.InvocationInfo.Line.Trim())" -ForegroundColor Red
        Write-Host "on line $($_.InvocationInfo.ScriptLineNumber), character $($_.InvocationInfo.OffsetInLine)" -ForegroundColor Red
    }

    if ($null -ne $writer) {
        $writer.Close()
    }
    if ($null -ne $stream) {
        $stream.Close()
    }
    if ($null -ne $CurrentReader) {
        $CurrentReader.Close()
    }

    try {
        if (Test-Path -LiteralPath $filename) {
            Remove-Item -LiteralPath $filename -Force
        }

        Rename-Item -LiteralPath $tmp_filename -NewName $filename
        $player.Stop()
        $global:VP_Filename = $filename
        $ReadVPP.Invoke($filename)
    } catch {
        Write-Host $_ -ForegroundColor Red
    }
}

[ScriptBlock]$SaveAsCommand = {
    $sfd = New-Object Microsoft.Win32.SaveFileDialog -Property @{
        Filter = "Volition Packages (*.vp)|*.vp"
        Filename = $global:VP_Filename
    }
    [string]$result = "Cancel"

    if ($sfd.ShowDialog()) {
        $SaveCommand.Invoke($sfd.FileName)
        $result = "OK"
    }

    return $result
}

[ScriptBlock]$CreateNewVPP = {
    $global:VP_Filename = [string]::Empty
    $Sync.ViewFiles.Clear()

    $Sync.Gui.treeVPP.Items.Clear()
    $Sync.Gui.treeVPP.Items.Add((New-Object System.Windows.Controls.TreeViewItem -Property @{
        Header = "data"
        Tag = "data"
    }))

    $global:Flag_Changed = $false
}

[ScriptBlock]$SaveOrSaveAs = {
    [string]$result = "OK"

    if ($global:VP_Filename.Length -gt 0) {
        $SaveCommand.Invoke($global:VP_Filename)
    } else {
        $result = $SaveAsCommand.Invoke()[0]
    }

    return $result
}

[ScriptBlock]$ConfirmChanges = {
    if ($global:Flag_Changed) {
        $response = [Windows.MessageBox]::Show("Save changes?", "Save changes", "YesNoCancel", "Question")

        switch ($response) {
            "Yes" {
                $result = $SaveOrSaveAs.Invoke()[0]
                if ($result -eq "OK") {
                    $CreateNewVPP.Invoke()
                }
            }

            "No" {
                $CreateNewVPP.Invoke()
            }
        }
    } else {
        $CreateNewVPP.Invoke()
    }
}

[ScriptBlock]$WriteDataAndTimestamp = {
    param(
        [string]$FilePath,
        [FileViewModel]$item
    )

    [byte[]]$buffer = $null

    switch ($item.SeekType) {
        ([int][SeekTypes]::Present) {
            $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($global:VP_Filename, "Open")))
            [void]$reader.BaseStream.Seek($item.Offset, "Begin")
            $buffer = [byte[]]::CreateInstance([byte], $item.Size)
            [void]$reader.Read($buffer, 0, $item.Size)
            $reader.Close()
        }
    }

    if (($item.SeekType -eq ([int][SeekTypes]::Present)) -and ($null -ne $buffer)) {
        [System.IO.File]::WriteAllBytes($FilePath, $buffer)

        # write created/modified date
        $dtDateTime = $ConvertTimestamp.Invoke($item.Timestamp)[0]
        $FileData = Get-Item -LiteralPath $FilePath
        $FileData.LastWriteTime = $dtDateTime
        $FileData.CreationTime = $dtDateTime
    }
}

[ScriptBlock]$OpenFile = {
    if ($Sync.Gui.gridFiles.SelectedItems.Count -eq 1) {
        $item = $Sync.Gui.gridFiles.SelectedItem
        $FilePath = Join-Path $env:temp $item.Filename
        $WriteDataAndTimestamp.Invoke($FilePath, $item)
        Invoke-Item -LiteralPath $FilePath
    }
}

[ScriptBlock]$ReadVPP = {
    param([string]$FilePath)

    [System.IO.BinaryReader]$reader = $null
    [byte[]]$buffer = [byte[]]::CreateInstance([byte], 16)

    $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($FilePath, "Open")))

    [System.Windows.Controls.TreeViewItem]$node = $null
    $path = New-Object System.Collections.ArrayList

    # header
    [void]$reader.Read($buffer, 0, 16)
    [string]$header = [System.Text.Encoding]::ASCII.GetString($buffer, 0, 4) #VPVP
    [UInt32]$version = [System.BitConverter]::ToUInt32($buffer, 4) #2
    [UInt32]$diroffset = [System.BitConverter]::ToUInt32($buffer, 8)
    [UInt32]$direntries = [System.BitConverter]::ToUInt32($buffer, 12)

    if (($header -eq "VPVP") -and ($version -eq 2)) {
        $Sync.Gui.treeVPP.Items.Clear()
        $Sync.ViewFiles.Clear()
        $global:VP_Filename = $FilePath

        for ([int]$n = 0; $n -lt $direntries; $n++) {
            [void]$reader.BaseStream.Seek($diroffset, "Begin")

            $tbl = @{
                Offset = $reader.ReadUInt32() # 4 bytes
                Size = $reader.ReadUInt32() # 4 bytes
                Filename = [System.Text.Encoding]::ASCII.GetString($reader.ReadBytes(32)).Split([byte]0)[0] # 32 bytes
                Timestamp = $reader.ReadUInt32() # 4 bytes
            }

            # entry is a directory
            if (($tbl.Size -eq 0) -and ($tbl.Timestamp -eq 0)) {
                $tbl.Filename = $tbl.Filename.TrimEnd("?")

                if ($Sync.Gui.treeVPP.Items.Count -eq 0) {
                    # should always be the "data" root dir
                    $node = New-Object System.Windows.Controls.TreeViewItem -Property @{
                        Header = $tbl.Filename
                        Tag = $tbl.Filename
                    }

                    $Sync.Gui.treeVPP.Items.Add($node)

                    $path.Add($node.Header)
                } else {
                    # entry is a backdir
                    if ($tbl.Filename -eq "..") {
                        $path.RemoveAt($path.Count - 1)

                        if ($node.Parent -isnot [System.Windows.Controls.TreeView]) {
                            <#
                            $Sync.Gui.treeVPP.Items.Add((New-Object System.Windows.Controls.TreeViewItem -Property @{
                                Header = "All files"
                                Tag = [string]::Empty
                            }))
                        } else {
                            #>
                            $node = $node.Parent
                        }
                    } else {
                        $newNode = New-Object System.Windows.Controls.TreeViewItem -Property @{
                            Header = $tbl.Filename
                        }

                        $node.Items.Add($newNode)
                        $node.IsExpanded = $true

                        $path.Add($newNode.Header)
                        $newNode.Tag = [string]::Join("\", @($path))

                        $node = $newNode
                    }
                }
            } else {
                [void]$Sync.ViewFiles.Add((New-Object FileViewModel -Property @{
                    Offset = $tbl.Offset
                    Size = $tbl.Size
                    Filename = $tbl.Filename
                    Timestamp = $tbl.Timestamp
                    Path = [string]::Join("\", @($path))
                    SeekType = [SeekTypes]::Present
                }))
            }

            # 4 + 4 + 32 + 4 = 44
            $diroffset += 44
        }
    } else {
        throw "'$FilePath' is not a valid VP"
    }

    $reader.Close()
}

function Invoke-InputBox {
    param(
        [string]$WindowTitle,
        [string]$LabelText,
        [string]$DefaultValue
    )
    [Xml]$InputBoxXml = @"
<Window x:Name="InputBox" x:Class="WpfApp1.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfApp1"
        mc:Ignorable="d"
        Title="$WindowTitle" Visibility="Visible" ResizeMode="CanMinimize" Height="85" Width="450" WindowStartupLocation="CenterOwner">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="23"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Grid
            Grid.Row="0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="65"/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <TextBlock
                Grid.Column="0"
                Text="$LabelText"/>
            <TextBox
                Grid.Column="1"
                x:Name="txtFolderName"
                Text="$DefaultValue"/>
        </Grid>
        <Grid
            Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition/>
                <ColumnDefinition/>
                <ColumnDefinition/>
                <ColumnDefinition/>
            </Grid.ColumnDefinitions>
            <Button
                Grid.Column="1"
                x:Name="btnOK"
                Content="OK"/>
            <Button
                Grid.Column="2"
                x:Name="btnCancel"
                Content="Cancel"/>
        </Grid>
    </Grid>
</Window>
"@

    $InputBoxXml.Window.RemoveAttribute('x:Class')
    $InputBoxXml.Window.RemoveAttribute('mc:Ignorable')

    $InputBoxNs = New-Object -TypeName Xml.XmlNamespaceManager -ArgumentList $InputBoxXml.NameTable
    $InputBoxNs.AddNamespace('x', $InputBoxXml.DocumentElement.x)
    $InputBoxNs.AddNamespace('d', $InputBoxXml.DocumentElement.d)
    $InputBoxNs.AddNamespace('mc', $InputBoxXml.DocumentElement.mc)

    $Sync.InputBoxGui = @{}

    try {
        $Sync.InputBox = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $InputBoxXml))
    } catch {
        Write-Host $_ -ForegroundColor Red
        Exit
    }

    $InputBoxXml.SelectNodes('//*[@x:Name]', $InputBoxNs) | ForEach-Object {
        $Sync.InputBoxGui.Add($_.Name, $Sync.InputBox.FindName($_.Name))
    }

    $Sync.InputBox.Owner = $Sync.Window

    $Sync.InputBox.add_Loaded({
        $Sync.InputBoxGui.txtFolderName.Text = [string]::Empty
        $Sync.InputBoxGui.txtFolderName.Focus()
    })

    $btnPressOK = {
        if ($Sync.InputBoxGui.txtFolderName.Text.Trim().Length -gt 0) {
            $Sync.InputBoxReturnValue = $Sync.InputBoxGui.txtFolderName.Text.Trim()
            $Sync.InputBox.Close()
        } else {
            [Windows.MessageBox]::Show("Folder name cannot be empty.")
        }
    }

    $Sync.InputBox.add_PreviewKeyDown({
        if ($_.Key -eq "Enter") {
            $btnPressOK.Invoke()
        }
    })

    $Sync.InputBoxGui.btnOK.add_Click($btnPressOK)

    $Sync.InputBoxGui.btnCancel.add_Click({
        $Sync.InputBoxReturnValue = [string]::Empty
        $Sync.InputBox.Close()
    })

    [void]$Sync.InputBox.ShowDialog()
}
#endregion

#region Form element event handlers
$Sync.Gui.mnuNew.add_Click($ConfirmChanges)

$Sync.Gui.mnuOpen.add_Click({
    $ofd = New-Object Microsoft.Win32.OpenFileDialog -Property @{
        Filter = "Volition Package files (*.vp)|*.vp"
    }

    if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
        $ofd.InitialDirectory = $InstallLocation
    }

    if ($global:VP_Filename.Length -gt 0) {
        $ofd.FileName = $global:VP_Filename
    }

    if ($ofd.ShowDialog()) {
        $player.Stop()

        try {
            $ReadVPP.Invoke($ofd.FileName)
        } catch {
            [void][Windows.MessageBox]::Show($_.Exception.Message)
        }
    }
})

$Sync.Gui.mnuSave.add_Click({
    [void]$SaveOrSaveAs.Invoke()
})

$Sync.Gui.mnuSaveAs.add_Click($SaveAsCommand)

$Sync.Gui.mnuExtractAll.add_Click({
    if ($Sync.ViewFiles.Count -gt 0) {
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            Description = "Browse to FS root folder"
        }

        if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
            $fbd.SelectedPath = $InstallLocation
        }

        if ($fbd.ShowDialog() -eq "OK") {
            $BasePath = $fbd.SelectedPath
            $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($global:VP_Filename, "Open")))
            [byte[]]$buffer = $null

            foreach ($file in $Sync.ViewFiles) {
                [string]$FileDir = Join-Path $BasePath $file.Path
                [string]$FilePath = Join-Path $FileDir $file.Filename

                if (-not (Test-Path -LiteralPath $FileDir)) {
                    mkdir $FileDir
                }

                $buffer = $null

                switch ($file.SeekType) {
                    0 {
                        [void]$reader.BaseStream.Seek($file.Offset, "Begin")
                        $buffer = [byte[]]::CreateInstance([byte], $file.Size)
                        [void]$reader.Read($buffer, 0, $file.Size)
                    }
                }

                if (($file.SeekType -eq 0) -and ($null -ne $buffer)) {
                    [System.IO.File]::WriteAllBytes($FilePath, $buffer)

                    # write created/modified date
                    $dtDateTime = $ConvertTimestamp.Invoke($file.Timestamp)[0]
                    $FileData = Get-Item -LiteralPath $FilePath
                    $FileData.LastWriteTime = $dtDateTime
                    $FileData.CreationTime = $dtDateTime
                }
            }

            $reader.Close()
        }
    }
})

$Sync.Gui.treeVPP.add_SelectedItemChanged({
    if ($Sync.ViewFiles.Count -gt 0) {
        [System.Windows.Data.CollectionViewSource]::GetDefaultView($Sync.Gui.gridFiles.ItemsSource).Filter = {
            param($obj)
            if ($null -ne $Sync.Gui.treeVPP.SelectedItem) {
                return ($obj.Path -like "$($Sync.Gui.treeVPP.SelectedItem.Tag.ToString())*")
            } else {
                return $null
            }
        }
    }
})

$Sync.Gui.treeVPP.add_ContextMenuOpening({
    foreach ($item in $Sync.Gui.treeVPP.ContextMenu.Items) {
        $item.Visibility = "Collapsed"
    }

    if ($null -ne $Sync.Gui.treeVPP.SelectedItem) {
        foreach ($item in $Sync.Gui.treeVPP.ContextMenu.Items) {
            if ($Sync.Gui.treeVPP.SelectedItem -eq $Sync.Gui.treeVPP.Items[0]) {
                # don't display "remove folder" or "rename folder"
                # context menu items for root node (data)
                if (-not @("mnuRemoveFolder","mnuRenameFolder").Contains($item.Name)) {
                    $item.Visibility = "Visible"
                }
            } else {
                $item.Visibility = "Visible"
            }
        }
    } else {
        $Sync.Gui.treeVPP.ContextMenu.IsOpen = $false
        $_.Handled = $true
    }
})

$Sync.Gui.treeVPP.add_MouseRightButtonDown({
    if ($null -ne $_.Source) {
        if ($_.Source -is [System.Windows.Controls.TreeViewItem]) {
            $_.Source.IsSelected = $true
            $_.Handled = $true
        }
    }
})

$Sync.Gui.mnuAddFolder.add_Click({
    Invoke-InputBox -WindowTitle "Add folder" -LabelText "Label:"

    if (($null -ne $Sync.InputBoxReturnValue) -and ($Sync.InputBoxReturnValue.Length -gt 0)) {
        $item = $Sync.Gui.treeVPP.SelectedItem

        $structure = New-Object System.Collections.ArrayList
        $structure.Add($Sync.InputBoxReturnValue)

        while ($item -is [System.Windows.Controls.TreeViewItem]) {
            $structure.Add($item.Header)
            $item = $item.Parent
        }

        $structure = @($structure)
        [Array]::Reverse($structure)

        $Sync.Gui.treeVPP.SelectedItem.Items.Add((New-Object System.Windows.Controls.TreeViewItem -Property @{
            Header = $Sync.InputBoxReturnValue
            Tag = [string]::Join("\", $structure)
        }))

        $Sync.Gui.treeVPP.SelectedItem.IsExpanded = $true

        $global:Flag_Changed = $true
    }
})

$Sync.Gui.mnuRemoveFolder.add_Click({
    $item = $Sync.Gui.treeVPP.SelectedItem
    $Sync.ViewFiles = $Sync.ViewFiles | Where-Object { $_.Path -ne $item.Tag.ToString() }
    $Sync.Gui.gridFiles.ItemsSource = $Sync.ViewFiles
    $item.Parent.Items.Remove($item)
    $Sync.Gui.treeVPP.Items[0].IsSelected = $true
})

$Sync.Gui.mnuRenameFolder.add_Click({
    [string]$OldTag = $Sync.Gui.treeVPP.SelectedItem.Tag.ToString()
    Invoke-InputBox -WindowTitle "Rename folder" -LabelText "Label:" -DefaultValue $Sync.Gui.treeVPP.SelectedItem.Header

    if (($null -ne $Sync.InputBoxReturnValue) -and ($Sync.InputBoxReturnValue.Length -gt 0)) {
        $Sync.Gui.treeVPP.SelectedItem.Header = $Sync.InputBoxReturnValue

        $item = $Sync.Gui.treeVPP.SelectedItem
        $structure = New-Object System.Collections.ArrayList
        while ($item -is [System.Windows.Controls.TreeViewItem]) {
            $structure.Add($item.Header)
            $item = $item.Parent
        }

        $structure = @($structure)
        [Array]::Reverse($structure)
        [string]$NewTag = [string]::Join("\", $structure)
        $Sync.Gui.treeVPP.SelectedItem.Tag = $NewTag

        # update file paths in this folder
        $files = $Sync.ViewFiles | Where-Object { $_.Path -eq $OldTag }

        foreach ($file in $files) {
            $file.Path = $NewTag
        }
    }
})

[ScriptBlock]$AddFilesFunc = {
    $ofd = New-Object Microsoft.Win32.OpenFileDialog -Property @{
        Filter = "All files (*.*)|*.*"
        Multiselect = $true
    }

    if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
        $ofd.InitialDirectory = $InstallLocation
    }

    if ($ofd.ShowDialog()) {
        $item = $Sync.Gui.treeVPP.SelectedItem
        $path = if ($null -ne $item) { $item.Tag.ToString() } else { "data" }

        foreach ($filename in $ofd.FileNames) {
            $FileItem = Get-Item -LiteralPath $filename

            [void]$Sync.ViewFiles.Add((New-Object FileViewModel -Property @{
                Offset = 0
                Size = $FileItem.Length
                Filename = $filename
                Timestamp = (New-TimeSpan -Start (Get-Date "01/01/1970") -End ($FileItem.LastWriteTime)).TotalSeconds
                Path = $path
                SeekType = [SeekTypes]::FileSystem
            }))

            $global:Flag_Changed = $true
        }
    }
}

$Sync.Gui.mnuAddFiles.add_Click($AddFilesFunc)

$Sync.Gui.gridFiles.add_PreviewKeyDown({
    $item = $_
    $grid = $Sync.Gui.gridFiles

    switch ($_.Key) {
        "Home" {
            if ($grid.Items.Count -gt 0) {
                $grid.SelectedItem = $grid.Items[0]
                $grid.ScrollIntoView($grid.Items[0])
                $item.Handled = $true
            }
        }

        "End" {
            if ($grid.Items.Count -gt 0) {
                $grid.SelectedItem = $grid.Items[$grid.Items.Count - 1]
                $grid.ScrollIntoView($grid.Items[$grid.Items.Count - 1])
                $item.Handled = $true
            }
        }
    }
})

$Sync.Gui.gridFiles.add_ContextMenuOpening({
    foreach ($item in $Sync.Gui.gridFiles.ContextMenu.Items) {
        $item.Visibility = "Collapsed"
    }

    if (($Sync.Gui.gridFiles.Items.Count -gt 0) -and ($null -ne $Sync.Gui.gridFiles.SelectedItem)) {
        foreach ($item in $Sync.Gui.gridFiles.ContextMenu.Items) {
            $item.Visibility = "Visible"
        }
    } else {
        $Sync.Gui.gridFiles.ContextMenu.IsOpen = $false
        $_.Handled = $true
    }
})

$Sync.Gui.mnuExtract.add_Click({
    if ($Sync.Gui.gridFiles.SelectedItems.Count -gt 1) {
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            Description = "Browse to folder"
        }

        if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
            $fbd.SelectedPath = $InstallLocation
        }

        if ($fbd.ShowDialog() -eq "OK") {
            foreach ($item in $Sync.Gui.gridFiles.SelectedItems) {
                [string]$FilePath = Join-Path $fbd.SelectedPath $item.Filename
                $WriteDataAndTimestamp.Invoke($FilePath, $item)
            }
        }
    } elseif ($Sync.Gui.gridFiles.SelectedItems.Count -eq 1) {
        $item = $Sync.Gui.gridFiles.SelectedItem

        $extension = [System.IO.Path]::GetExtension($item.Filename)

        $sfd = New-Object Microsoft.Win32.SaveFileDialog -Property @{
            Filter = if ($extension.Length -eq 0) {"All files (*.*)|*.*"} else {"$($extension.ToUpper().TrimStart(".")) files (*$extension)|*$extension|All files (*.*)|*.*"}
            FileName = $item.Filename
        }

        if ($sfd.ShowDialog()) {
            $WriteDataAndTimestamp.Invoke($sfd.FileName, $item)
        }
    }
})

$Sync.Gui.gridFiles.add_SelectionChanged({
    if ($Sync.Gui.gridFiles.SelectedItems.Count -eq 1) {
        $item = $Sync.Gui.gridFiles.SelectedItem

        $player.Stop()

        if ($item.FileName.ToLower().EndsWith(".wav")) {
            $player.Play()
        }
    }
})

$Sync.Gui.mnuOpenFile.add_Click($OpenFile)

$Sync.Gui.gridFiles.add_MouseDoubleClick($OpenFile)

$Sync.Gui.mnuRemoveFiles.add_Click({
    if ($Sync.Gui.gridFiles.SelectedItems.Count -gt 1) {
        $items = @($Sync.Gui.gridFiles.SelectedItems)

        foreach ($item in $items) {
            $Sync.ViewFiles.Remove($item)
        }
    } elseif ($Sync.Gui.gridFiles.SelectedItems.Count -eq 1) {
        $Sync.ViewFiles.Remove($Sync.Gui.gridFiles.SelectedItem)
    }
})
#endregion

#region Window events
$Sync.Window.add_Loaded({
    $CreateNewVPP.Invoke()

    # bring app to front in VSCode
    if ($host.Name -eq "Visual Studio Code Host") {
        Invoke-Expression @'
    $z="$global:PSScriptRoot\WinAPI.ps1"
    if(Test-Path $z){
        ipmo $z
        $h=(gps -Id $PID).MainWindowHandle
        Show-Window $h 2
        Show-Window $h 1
        rmo WinAPI
    }
'@
    }
})

$Sync.Window.add_Closing({
    $player.Stop()
    $Runspace.Close()
})

$Sync.Window.add_Closed({

})
#endregion

# display the form
[void]$Sync.Window.ShowDialog()
