Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms # For FolderBrowserDialog

#region Global variables for last used paths
$Global:LastUsedContentPath_Tab1 = $null
$Global:LastUsedLinkPath_Tab1 = $null
$Global:LastUsedContentPath_Tab2 = $null
$Global:LastUsedLinkPath_Tab2 = $null
#endregion

#region GUI Structure
[xml]$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:shell="clr-namespace:System.Windows.Shell;assembly=PresentationFramework"
        Title="Make Junctions" Height="450" Width="600"
        WindowStartupLocation="CenterScreen"
        AllowsTransparency="True" WindowStyle="None" Background="Transparent"
        Name="MainWindow" ResizeMode="CanResize">

    <shell:WindowChrome.WindowChrome>
        <shell:WindowChrome ResizeBorderThickness="8"
                            CaptionHeight="30"
                            CornerRadius="0"
                            GlassFrameThickness="0"
                            UseAeroCaptionButtons="False"/>
    </shell:WindowChrome.WindowChrome>

    <Border BorderBrush="#FF007ACC" BorderThickness="1" Background="#FF2D2D30" CornerRadius="0">
        <Grid>
            <Grid.RowDefinitions>
                <RowDefinition Height="Auto"/><RowDefinition Height="*"/>   </Grid.RowDefinitions>

            <Border x:Name="CustomTitleBar" Grid.Row="0" Background="#FF252526" Height="30" CornerRadius="0,0,0,0">
                <Grid>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>
                    <TextBlock Text="Make Junctions" VerticalAlignment="Center" HorizontalAlignment="Left" Margin="10,0,0,0" Foreground="White" FontWeight="SemiBold"/>
                    <Button x:Name="CloseButton" Grid.Column="1" Content="✕" Width="40" Height="30"
                            shell:WindowChrome.IsHitTestVisibleInChrome="True">
                        <Button.Style>
                            <Style TargetType="Button">
                                <Setter Property="Background" Value="#FF252526"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter Property="BorderThickness" Value="0"/>
                                <Setter Property="Template">
                                    <Setter.Value>
                                        <ControlTemplate TargetType="Button">
                                            <Border Background="{TemplateBinding Background}">
                                                <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                            </Border>
                                        </ControlTemplate>
                                    </Setter.Value>
                                </Setter>
                                <Style.Triggers>
                                    <Trigger Property="IsMouseOver" Value="True">
                                        <Setter Property="Background" Value="#FFE81123"/>
                                    </Trigger>
                                </Style.Triggers>
                            </Style>
                        </Button.Style>
                    </Button>
                </Grid>
            </Border>

            <TabControl Grid.Row="1" Background="#FF252526" BorderThickness="0" Margin="0,1,0,0" Name="MainTabControl">
                <TabControl.Resources>
                    <Style TargetType="TabItem">
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="TabItem">
                                    <Grid Name="Panel">
                                        <ContentPresenter x:Name="ContentSite"
                                            VerticalAlignment="Center"
                                            HorizontalAlignment="Center"
                                            ContentSource="Header"
                                            Margin="10,2"/>
                                    </Grid>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsSelected" Value="True">
                                            <Setter TargetName="Panel" Property="Background" Value="#FF007ACC" />
                                            <Setter Property="Foreground" Value="White"/>
                                        </Trigger>
                                        <Trigger Property="IsSelected" Value="False">
                                            <Setter TargetName="Panel" Property="Background" Value="#FF3E3E42" />
                                            <Setter Property="Foreground" Value="#FF999999"/>
                                        </Trigger>
                                        <MultiTrigger>
                                            <MultiTrigger.Conditions>
                                                <Condition Property="IsMouseOver" Value="True" />
                                                <Condition Property="IsSelected" Value="False" />
                                            </MultiTrigger.Conditions>
                                            <Setter TargetName="Panel" Property="Background" Value="#FF4F4F53"/>
                                            <Setter Property="Foreground" Value="White"/>
                                        </MultiTrigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                        <Setter Property="Padding" Value="10,5"/>
                        <Setter Property="Foreground" Value="#FF999999"/>
                        <Setter Property="FontWeight" Value="SemiBold"/>
                    </Style>
                    <Style x:Key="BrowseButtonStyle" TargetType="Button">
                        <Setter Property="Background" Value="#FF3E3E42"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="BorderBrush" Value="#FF007ACC"/>
                        <Setter Property="BorderThickness" Value="1"/>
                        <Setter Property="Padding" Value="8,3"/>
                        <Setter Property="Margin" Value="5,5,0,5"/>
                        <Setter Property="VerticalContentAlignment" Value="Center"/>
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}"
                                            BorderBrush="{TemplateBinding BorderBrush}"
                                            BorderThickness="{TemplateBinding BorderThickness}"
                                            CornerRadius="2">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter Property="Background" Value="#FF4F4F53"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter Property="Background" Value="#FF007ACC"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                    <Style x:Key="ActionButtonStyle" TargetType="Button">
                        <Setter Property="Background" Value="#FF007ACC"/>
                        <Setter Property="Foreground" Value="White"/>
                        <Setter Property="BorderThickness" Value="0"/>
                        <Setter Property="FontWeight" Value="Bold"/>
                        <Setter Property="Padding" Value="10,7"/> 
                        <Setter Property="Margin" Value="5,15,5,5"/>
                        <Setter Property="Height" Value="60"/> 
                        <Setter Property="Template">
                            <Setter.Value>
                                <ControlTemplate TargetType="Button">
                                    <Border Background="{TemplateBinding Background}" CornerRadius="2">
                                        <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                    </Border>
                                    <ControlTemplate.Triggers>
                                        <Trigger Property="IsMouseOver" Value="True">
                                            <Setter Property="Background" Value="#FF005A9E"/>
                                        </Trigger>
                                        <Trigger Property="IsPressed" Value="True">
                                            <Setter Property="Background" Value="#FF004C87"/>
                                        </Trigger>
                                    </ControlTemplate.Triggers>
                                </ControlTemplate>
                            </Setter.Value>
                        </Setter>
                    </Style>
                </TabControl.Resources>
                <TabItem Header="Make In Folder" Name="TabMakeInFolder">
                    <Grid Margin="10"> 
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="*"/>    
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/> 
                            <ColumnDefinition Width="*"/>    
                            <ColumnDefinition Width="Auto"/> 
                        </Grid.ColumnDefinitions>

                        <Label Content="Content Path:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Foreground="White" Margin="0,0,10,0"/>
                        <TextBox Name="TextBoxContentPath_Tab1" Grid.Row="0" Grid.Column="1" Margin="5" Padding="5" Background="#FF3E3E42" Foreground="White" BorderBrush="#FF007ACC" VerticalContentAlignment="Center" AllowDrop="True"/>
                        <Button Name="ButtonBrowseContentPath_Tab1" Grid.Row="0" Grid.Column="2" Content="Browse..." Style="{StaticResource BrowseButtonStyle}"/>

                        <Label Content="Link Path:" Grid.Row="1" Grid.Column="0" VerticalAlignment="Center" Foreground="White" Margin="0,0,10,0"/>
                        <TextBox Name="TextBoxLinkPath_Tab1" Grid.Row="1" Grid.Column="1" Margin="5" Padding="5" Background="#FF3E3E42" Foreground="White" BorderBrush="#FF007ACC" VerticalContentAlignment="Center" AllowDrop="True"/>
                        <Button Name="ButtonBrowseLinkPath_Tab1" Grid.Row="1" Grid.Column="2" Content="Browse..." Style="{StaticResource BrowseButtonStyle}"/>
                        
                        <Button Name="ButtonMakeJunction_Tab1" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="3" Content="Make Junction" Style="{StaticResource ActionButtonStyle}"/>
                    </Grid>
                </TabItem>
                <TabItem Header="Replace Folder" Name="TabReplaceFolder">
                    <Grid Margin="10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="*"/>    
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/> 
                            <ColumnDefinition Width="*"/>    
                            <ColumnDefinition Width="Auto"/> 
                        </Grid.ColumnDefinitions>

                        <Label Content="Content Path:" Grid.Row="0" Grid.Column="0" VerticalAlignment="Center" Foreground="White" Margin="0,0,10,0"/>
                        <TextBox Name="TextBoxContentPath_Tab2" Grid.Row="0" Grid.Column="1" Margin="5" Padding="5" Background="#FF3E3E42" Foreground="White" BorderBrush="#FF007ACC" VerticalContentAlignment="Center" AllowDrop="True"/>
                        <Button Name="ButtonBrowseContentPath_Tab2" Grid.Row="0" Grid.Column="2" Content="Browse..." Style="{StaticResource BrowseButtonStyle}"/>

                        <Label Content="Link Path:" Grid.Row="1" Grid.Column="0" VerticalAlignment="Center" Foreground="White" Margin="0,0,10,0"/>
                        <TextBox Name="TextBoxLinkPath_Tab2" Grid.Row="1" Grid.Column="1" Margin="5" Padding="5" Background="#FF3E3E42" Foreground="White" BorderBrush="#FF007ACC" VerticalContentAlignment="Center" AllowDrop="True" ToolTip="The folder at this exact path will be replaced by a junction."/>
                        <Button Name="ButtonBrowseLinkPath_Tab2" Grid.Row="1" Grid.Column="2" Content="Browse..." Style="{StaticResource BrowseButtonStyle}"/>
                        
                        <Button Name="ButtonMakeJunction_Tab2" Grid.Row="2" Grid.Column="0" Grid.ColumnSpan="3" Content="Replace Folder with Junction" Style="{StaticResource ActionButtonStyle}"/>
                    </Grid>
                </TabItem>
                <TabItem Header="From Text" Name="TabFromText">
                    <Grid Margin="10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/> 
                            <RowDefinition Height="*"/>    
                            <RowDefinition Height="Auto"/> 
                        </Grid.RowDefinitions>

                        <TextBlock Text="Add a list of junctions to be made here. Format is Content Path and Link Path, separated by a tab. Any junctions made in the other tabs will be added here for future reference."
                                   Grid.Row="0"
                                   TextWrapping="Wrap"
                                   Foreground="#FFCCCCCC" 
                                   Margin="0,0,0,10"/>

                        <TextBox Name="TextBoxLogAndInput"
                                 Grid.Row="1"
                                 AcceptsReturn="True"
                                 AcceptsTab="True" 
                                 VerticalScrollBarVisibility="Auto"
                                 HorizontalScrollBarVisibility="Auto"
                                 TextWrapping="NoWrap"
                                 Background="#FF1E1E1E" 
                                 Foreground="White"
                                 BorderBrush="#FF007ACC"
                                 BorderThickness="1"
                                 Padding="5"
                                 FontFamily="Consolas"
                                 IsReadOnly="False"/>
                        
                        <Button Name="ButtonMakeJunctionsFromText"
                                Grid.Row="2"
                                Content="Make Junctions from List" 
                                Style="{StaticResource ActionButtonStyle}"
                                Margin="0,10,0,0"/>
                    </Grid>
                </TabItem>
            </TabControl>

            <Grid x:Name="ResizeGripVisual" Grid.Row="1" HorizontalAlignment="Right" VerticalAlignment="Bottom" 
                  Width="18" Height="18" Margin="0,0,0,0" IsHitTestVisible="False">
                <Path Stroke="#888888" StrokeThickness="1.5">
                    <Path.Data>
                        <LineGeometry StartPoint="5,13" EndPoint="13,5"/>
                    </Path.Data>
                </Path>
                <Path Stroke="#888888" StrokeThickness="1.5">
                    <Path.Data>
                        <LineGeometry StartPoint="9,13" EndPoint="13,9"/>
                    </Path.Data>
                </Path>
            </Grid>
        </Grid>
    </Border>
</Window>
"@

$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try {
    $Global:Window=[Windows.Markup.XamlReader]::Load( $reader )
} catch {
    Write-Error "Error loading XAML: $($_.Exception.Message)"
    Write-Host "--- XAML Content ---"; Write-Host $xaml; Write-Host "--- End XAML Content ---"
    exit 1
}

#endregion

#region Global Variables (Controls)
# Tab 1
$TextBoxContentPath_Tab1 = $Global:Window.FindName("TextBoxContentPath_Tab1")
$TextBoxLinkPath_Tab1 = $Global:Window.FindName("TextBoxLinkPath_Tab1")
$ButtonBrowseContentPath_Tab1 = $Global:Window.FindName("ButtonBrowseContentPath_Tab1")
$ButtonBrowseLinkPath_Tab1 = $Global:Window.FindName("ButtonBrowseLinkPath_Tab1")
$ButtonMakeJunction_Tab1 = $Global:Window.FindName("ButtonMakeJunction_Tab1")

# Tab 2
$TextBoxContentPath_Tab2 = $Global:Window.FindName("TextBoxContentPath_Tab2")
$TextBoxLinkPath_Tab2 = $Global:Window.FindName("TextBoxLinkPath_Tab2")
$ButtonBrowseContentPath_Tab2 = $Global:Window.FindName("ButtonBrowseContentPath_Tab2")
$ButtonBrowseLinkPath_Tab2 = $Global:Window.FindName("ButtonBrowseLinkPath_Tab2")
$ButtonMakeJunction_Tab2 = $Global:Window.FindName("ButtonMakeJunction_Tab2")

# Tab 3
$TextBoxLogAndInput = $Global:Window.FindName("TextBoxLogAndInput")
$ButtonMakeJunctionsFromText = $Global:Window.FindName("ButtonMakeJunctionsFromText")

# Common
$CustomTitleBar = $Global:Window.FindName("CustomTitleBar")
$CloseButton = $Global:Window.FindName("CloseButton")
#endregion

#region Helper Functions
Function Show-FolderBrowserDialog {
    param([string]$InitialPathHint)
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select a folder"; $folderBrowserDialog.ShowNewFolderButton = $true
    if (-not [string]::IsNullOrEmpty($InitialPathHint) -and (Test-Path -Path $InitialPathHint -PathType Container)) {
        $folderBrowserDialog.SelectedPath = $InitialPathHint
    } elseif ($PSScriptRoot -and (Test-Path -Path $PSScriptRoot -PathType Container)) {
        $folderBrowserDialog.SelectedPath = $PSScriptRoot
    } elseif ($PWD -and (Test-Path -Path $PWD.Path -PathType Container)) {
        $folderBrowserDialog.SelectedPath = $PWD.Path
    }
    if ($folderBrowserDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowserDialog.SelectedPath
    }
    return $null
}

Function Test-IsReparsePoint {
    param([string]$Path)
    if (Test-Path -Path $Path) { 
        try {
            $item = Get-Item -Path $Path -Force -ErrorAction Stop
            return $item.Attributes.HasFlag([System.IO.FileAttributes]::ReparsePoint)
        } catch { return $false }
    }
    return $false
}

Function Show-MessageBox {
    param(
        [string]$Message,
        [string]$Title = "Make Junctions",
        [System.Windows.MessageBoxButton]$Button = [System.Windows.MessageBoxButton]::OK,
        [System.Windows.MessageBoxImage]$Image = [System.Windows.MessageBoxImage]::None
    )
    return [System.Windows.MessageBox]::Show($Global:Window, $Message, $Title, $Button, $Image)
}
#endregion

#region Event Handlers

$ScriptBlock_TextBox_PreviewDragEnterOver = {
    param($sender, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $e.Effects = [System.Windows.DragDropEffects]::Copy
    } else {
        $e.Effects = [System.Windows.DragDropEffects]::None
    }
    $e.Handled = $true 
}

# --- Drop Handlers for Tab 1 ---
$ScriptBlock_Drop_ContentPath_Tab1 = {
    param($sender, $e) 
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $filePaths = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($filePaths -and $filePaths.Count -gt 0) {
            $firstPath = $filePaths[0]
            if (Test-Path -Path $firstPath -PathType Container) {
                $TextBoxContentPath_Tab1.Text = $firstPath 
                $Global:LastUsedContentPath_Tab1 = $firstPath
            } else { Show-MessageBox -Message "Please drop a single folder." -Title "Invalid Drop" -Image Warning }
        }
    }
    $e.Handled = $true
}
$ScriptBlock_Drop_LinkPath_Tab1 = {
    param($sender, $e) 
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $filePaths = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($filePaths -and $filePaths.Count -gt 0) {
            $firstPath = $filePaths[0]
            if (Test-Path -Path $firstPath -PathType Container) {
                $TextBoxLinkPath_Tab1.Text = $firstPath 
                $Global:LastUsedLinkPath_Tab1 = $firstPath
            } else { Show-MessageBox -Message "Please drop a single folder." -Title "Invalid Drop" -Image Warning }
        }
    }
    $e.Handled = $true
}

# --- Drop Handlers for Tab 2 ---
$ScriptBlock_Drop_ContentPath_Tab2 = {
    param($sender, $e) 
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $filePaths = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($filePaths -and $filePaths.Count -gt 0) {
            $firstPath = $filePaths[0]
            if (Test-Path -Path $firstPath -PathType Container) {
                $TextBoxContentPath_Tab2.Text = $firstPath 
                $Global:LastUsedContentPath_Tab2 = $firstPath
            } else { Show-MessageBox -Message "Please drop a single folder." -Title "Invalid Drop" -Image Warning }
        }
    }
    $e.Handled = $true
}
$ScriptBlock_Drop_LinkPath_Tab2 = {
    param($sender, $e) 
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) {
        $filePaths = $e.Data.GetData([System.Windows.DataFormats]::FileDrop)
        if ($filePaths -and $filePaths.Count -gt 0) {
            $firstPath = $filePaths[0]
            if (Test-Path -Path $firstPath -PathType Container) {
                $TextBoxLinkPath_Tab2.Text = $firstPath 
                $Global:LastUsedLinkPath_Tab2 = $firstPath
            } else { Show-MessageBox -Message "Please drop a single folder." -Title "Invalid Drop" -Image Warning }
        }
    }
    $e.Handled = $true
}

# --- Attach Drag-Drop Event Handlers ---
# Tab 1
$TextBoxContentPath_Tab1.Add_PreviewDragEnter($ScriptBlock_TextBox_PreviewDragEnterOver)
$TextBoxContentPath_Tab1.Add_PreviewDragOver($ScriptBlock_TextBox_PreviewDragEnterOver) 
$TextBoxContentPath_Tab1.Add_Drop($ScriptBlock_Drop_ContentPath_Tab1)
$TextBoxLinkPath_Tab1.Add_PreviewDragEnter($ScriptBlock_TextBox_PreviewDragEnterOver)
$TextBoxLinkPath_Tab1.Add_PreviewDragOver($ScriptBlock_TextBox_PreviewDragEnterOver) 
$TextBoxLinkPath_Tab1.Add_Drop($ScriptBlock_Drop_LinkPath_Tab1)
# Tab 2
$TextBoxContentPath_Tab2.Add_PreviewDragEnter($ScriptBlock_TextBox_PreviewDragEnterOver)
$TextBoxContentPath_Tab2.Add_PreviewDragOver($ScriptBlock_TextBox_PreviewDragEnterOver) 
$TextBoxContentPath_Tab2.Add_Drop($ScriptBlock_Drop_ContentPath_Tab2)
$TextBoxLinkPath_Tab2.Add_PreviewDragEnter($ScriptBlock_TextBox_PreviewDragEnterOver)
$TextBoxLinkPath_Tab2.Add_PreviewDragOver($ScriptBlock_TextBox_PreviewDragEnterOver) 
$TextBoxLinkPath_Tab2.Add_Drop($ScriptBlock_Drop_LinkPath_Tab2)

# --- Common UI Event Handlers ---
$CustomTitleBar.Add_MouseLeftButtonDown({
    param($sender, $e)
    $clickedElement = $e.OriginalSource
    $isClickOnCloseButtonOrChild = $false
    while ($clickedElement -ne $null) {
        if ($clickedElement -eq $CloseButton) { $isClickOnCloseButtonOrChild = $true; break }
        if ($clickedElement -eq $CustomTitleBar) { break }
        if ($clickedElement -is [System.Windows.DependencyObject]) {
            $clickedElement = [System.Windows.Media.VisualTreeHelper]::GetParent($clickedElement)
        } else { break }
    }
    if (-not $isClickOnCloseButtonOrChild -and $e.ButtonState -eq 'Pressed') {
        $Global:Window.DragMove()
    }
})
$CloseButton.Add_Click({ $Global:Window.Close() })

# --- Tab 1 Event Handlers ---
$ButtonBrowseContentPath_Tab1.Add_Click({
    $currentPathInTextBox = $TextBoxContentPath_Tab1.Text.Trim()
    $pathHint = $null
    if (-not [string]::IsNullOrEmpty($currentPathInTextBox) -and (Test-Path -Path $currentPathInTextBox -PathType Container)) {
        $pathHint = $currentPathInTextBox
    } elseif (-not [string]::IsNullOrEmpty($Global:LastUsedContentPath_Tab1)) {
        $pathHint = $Global:LastUsedContentPath_Tab1
    } 
    $selectedPath = Show-FolderBrowserDialog -InitialPathHint $pathHint
    if (-not [string]::IsNullOrEmpty($selectedPath)) {
        $TextBoxContentPath_Tab1.Text = $selectedPath
        $Global:LastUsedContentPath_Tab1 = $selectedPath
    }
})
$ButtonBrowseLinkPath_Tab1.Add_Click({
    $currentPathInTextBox = $TextBoxLinkPath_Tab1.Text.Trim()
    $pathHint = $null
    if (-not [string]::IsNullOrEmpty($currentPathInTextBox) -and (Test-Path -Path $currentPathInTextBox -PathType Container)) {
        $pathHint = $currentPathInTextBox
    } elseif (-not [string]::IsNullOrEmpty($Global:LastUsedLinkPath_Tab1)) {
        $pathHint = $Global:LastUsedLinkPath_Tab1
    }
    $selectedPath = Show-FolderBrowserDialog -InitialPathHint $pathHint
    if (-not [string]::IsNullOrEmpty($selectedPath)) {
        $TextBoxLinkPath_Tab1.Text = $selectedPath
        $Global:LastUsedLinkPath_Tab1 = $selectedPath
    }
})
$ButtonMakeJunction_Tab1.Add_Click({
    $contentPath = $TextBoxContentPath_Tab1.Text.Trim()
    $linkParentPath = $TextBoxLinkPath_Tab1.Text.Trim()

    if ([string]::IsNullOrWhiteSpace($contentPath) -or [string]::IsNullOrWhiteSpace($linkParentPath)) {
        Show-MessageBox -Message "Both 'Content Path' and 'Link Path' must be specified." -Title "Input Required" -Image Error; return
    }
    if ($contentPath.StartsWith("\\") -or $linkParentPath.StartsWith("\\")) {
        Show-MessageBox -Message "Network paths (UNC paths like \\server\share) are not supported for the Content Path or Link Path in this version. Please use local paths." -Title "Network Path Not Supported" -Image Warning; return
    }
    if (-not (Test-Path -Path $contentPath -PathType Container)) {
        Show-MessageBox -Message "The 'Content Path' does not exist or is not a folder.`n`nPath: $contentPath" -Title "Invalid Content Path" -Image Error; return
    }
    if (-not (Test-Path -Path $linkParentPath -PathType Container)) {
        Show-MessageBox -Message "The 'Link Path' (parent directory for the junction) does not exist or is not a folder.`n`nPath: $linkParentPath" -Title "Invalid Link Path" -Image Error; return
    }

    $contentFolderName = (Get-Item -Path $contentPath).Name
    $fullJunctionPath = Join-Path -Path $linkParentPath -ChildPath $contentFolderName
    $actionToTake = "CreateNew"
    $pathExists = Test-Path -Path $fullJunctionPath
    $isReparsePoint = $false; if ($pathExists) { $isReparsePoint = Test-IsReparsePoint -Path $fullJunctionPath }

    if ($pathExists) {
        if ($isReparsePoint) {
            $message = "A junction or symbolic link already exists at:`n{0}`n`nDo you want to delete it and create a new junction to:`n{1}?" -f $fullJunctionPath, $contentPath
            if ((Show-MessageBox -Message $message -Title "Existing Link Found" -Button YesNo -Image Question) -ne "Yes") { return }
            $actionToTake = "ReplaceExistingLink"
        } else { 
            if (Test-Path -Path $fullJunctionPath -PathType Container) { 
                if ((Get-ChildItem -Path $fullJunctionPath -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                    $message = "An empty folder exists at:`n{0}`n`nDo you want to delete it and replace it with a junction to:`n{1}?" -f $fullJunctionPath, $contentPath
                    if ((Show-MessageBox -Message $message -Title "Empty Folder Found" -Button YesNo -Image Question) -ne "Yes") { return }
                    $actionToTake = "ReplaceEmptyFolder"
                } else {
                    $message = "A folder with content already exists at:`n{0}`n`nWARNING: Replacing it will require deleting this folder and ALL its contents. This action CANNOT BE UNDONE.`n`nAre you absolutely sure you want to delete this folder and its content, then create a junction to:`n{1}?" -f $fullJunctionPath, $contentPath
                    if ((Show-MessageBox -Message $message -Title "Folder with Content Found" -Button YesNoCancel -Image Exclamation) -ne "Yes") { return }
                    $actionToTake = "ReplaceNonEmptyFolder"
                }
            } else { 
                 Show-MessageBox -Message "A file (not a folder or junction) already exists at:`n$fullJunctionPath`n`nPlease remove or rename this file and try again." -Title "File Exists at Junction Path" -Image Error; return
            }
        }
    }
    try {
        $operationMessage = ""; if ($actionToTake -in ("ReplaceExistingLink", "ReplaceEmptyFolder", "ReplaceNonEmptyFolder")) {
            Write-Host "Attempting to remove existing item at: $fullJunctionPath (Action: $actionToTake)"
            Remove-Item -Path $fullJunctionPath -Recurse -Force -ErrorAction Stop; $operationMessage = "Replaced existing item and "
        }
        Write-Host "Attempting to create junction: '$fullJunctionPath' -> '$contentPath' using New-Item."
        New-Item -ItemType Junction -Path $fullJunctionPath -Value $contentPath -Force -ErrorAction Stop | Out-Null
        
        $logEntry = "$contentPath`t$fullJunctionPath"
        $TextBoxLogAndInput.AppendText("$logEntry`r`n")

        $successMessage = "$($operationMessage)Junction successfully created!`n`nLink: {0}`nTarget: {1}" -f $fullJunctionPath, $contentPath
        Show-MessageBox -Message $successMessage -Title "Success" -Image Information
    } catch {
        $errorMessage = "An error occurred: $($_.Exception.Message)"; if ($_.Exception.InnerException) { $errorMessage += "`nInner Exception: $($_.Exception.InnerException.Message)" }
        if ($_.Exception.Message -match "symbolic link|not have sufficient privilege|privilege not held|0x80070522") {
             $errorMessage += "`n`nThis typically means PowerShell does not have sufficient privileges. Try running the script as an Administrator, or ensure Developer Mode is enabled on your system (for Windows 10/11)."
        }
        Show-MessageBox -Message $errorMessage -Title "Operation Failed" -Image Error; Write-Error $_
    }
})

# --- Tab 2 Event Handlers ---
$ButtonBrowseContentPath_Tab2.Add_Click({
    $currentPathInTextBox = $TextBoxContentPath_Tab2.Text.Trim()
    $pathHint = $null
    if (-not [string]::IsNullOrEmpty($currentPathInTextBox) -and (Test-Path -Path $currentPathInTextBox -PathType Container)) {
        $pathHint = $currentPathInTextBox
    } elseif (-not [string]::IsNullOrEmpty($Global:LastUsedContentPath_Tab2)) {
        $pathHint = $Global:LastUsedContentPath_Tab2
    }
    $selectedPath = Show-FolderBrowserDialog -InitialPathHint $pathHint
    if (-not [string]::IsNullOrEmpty($selectedPath)) {
        $TextBoxContentPath_Tab2.Text = $selectedPath
        $Global:LastUsedContentPath_Tab2 = $selectedPath
    }
})
$ButtonBrowseLinkPath_Tab2.Add_Click({
    $currentPathInTextBox = $TextBoxLinkPath_Tab2.Text.Trim()
    $pathHint = $null
    if (-not [string]::IsNullOrEmpty($currentPathInTextBox) -and (Test-Path -Path $currentPathInTextBox -PathType Container)) {
        $pathHint = $currentPathInTextBox
    } elseif (-not [string]::IsNullOrEmpty($Global:LastUsedLinkPath_Tab2)) {
        $pathHint = $Global:LastUsedLinkPath_Tab2
    }
    $selectedPath = Show-FolderBrowserDialog -InitialPathHint $pathHint
    if (-not [string]::IsNullOrEmpty($selectedPath)) {
        $TextBoxLinkPath_Tab2.Text = $selectedPath
        $Global:LastUsedLinkPath_Tab2 = $selectedPath
    }
})
$ButtonMakeJunction_Tab2.Add_Click({
    $contentPath = $TextBoxContentPath_Tab2.Text.Trim()
    $fullJunctionPath = $TextBoxLinkPath_Tab2.Text.Trim() 

    if ([string]::IsNullOrWhiteSpace($contentPath) -or [string]::IsNullOrWhiteSpace($fullJunctionPath)) {
        Show-MessageBox -Message "Both 'Content Path' and 'Link Path' must be specified." -Title "Input Required" -Image Error; return
    }
    if ($contentPath.StartsWith("\\") -or $fullJunctionPath.StartsWith("\\")) {
        Show-MessageBox -Message "Network paths (UNC paths like \\server\share) are not supported for the Content Path or Link Path in this version. Please use local paths." -Title "Network Path Not Supported" -Image Warning; return
    }
    if (-not (Test-Path -Path $contentPath -PathType Container)) {
        Show-MessageBox -Message "The 'Content Path' does not exist or is not a folder.`n`nPath: $contentPath" -Title "Invalid Content Path" -Image Error; return
    }
    $linkParentDir = Split-Path -Path $fullJunctionPath -Parent
    if (($linkParentDir -eq $null -and $fullJunctionPath -notmatch '^[a-zA-Z]:\\?$') -or (-not (Test-Path -Path $linkParentDir -PathType Container) -and $linkParentDir)) { 
        Show-MessageBox -Message "The parent directory for the 'Link Path' does not exist or is not a folder.`n`nAttempted Parent Directory: $(if ($linkParentDir) {$linkParentDir} else {'Root of a drive - invalid for junction name'})" -Title "Invalid Link Path Parent" -Image Error; return
    }
    if ($fullJunctionPath -eq $contentPath) { 
        Show-MessageBox -Message "Content Path and Link Path cannot be the same." -Title "Invalid Paths" -Image Error; return
    }

    $actionToTake = "CreateNew"
    $pathExists = Test-Path -Path $fullJunctionPath
    $isReparsePoint = $false; if ($pathExists) { $isReparsePoint = Test-IsReparsePoint -Path $fullJunctionPath }

    if ($pathExists) {
        if ($isReparsePoint) {
            $message = "A junction or symbolic link already exists at the specified Link Path:`n{0}`n`nDo you want to delete it and create a new junction to:`n{1}?" -f $fullJunctionPath, $contentPath
            if ((Show-MessageBox -Message $message -Title "Existing Link Found" -Button YesNo -Image Question) -ne "Yes") { return }
            $actionToTake = "ReplaceExistingLink"
        } else { 
            if (Test-Path -Path $fullJunctionPath -PathType Container) { 
                if ((Get-ChildItem -Path $fullJunctionPath -Force -ErrorAction SilentlyContinue).Count -eq 0) {
                    Write-Host "[Info] Empty folder found at Link Path '$fullJunctionPath' on Tab 2. Proceeding with replacement."
                    $actionToTake = "ReplaceEmptyFolder"
                } else { 
                    $message = "A folder with content already exists at the specified Link Path:`n{0}`n`nWARNING: Replacing it will require deleting this folder and ALL its contents. This action CANNOT BE UNDONE.`n`nAre you absolutely sure you want to delete this folder and its content, then create a junction to:`n{1}?" -f $fullJunctionPath, $contentPath
                    if ((Show-MessageBox -Message $message -Title "Folder with Content Found" -Button YesNoCancel -Image Exclamation) -ne "Yes") { return }
                    $actionToTake = "ReplaceNonEmptyFolder"
                }
            } else { 
                 Show-MessageBox -Message "A file (not a folder or junction) already exists at the specified Link Path:`n$fullJunctionPath`n`nPlease remove or rename this file and try again." -Title "File Exists at Link Path" -Image Error; return
            }
        }
    }
    try {
        $operationMessage = ""; if ($actionToTake -in ("ReplaceExistingLink", "ReplaceEmptyFolder", "ReplaceNonEmptyFolder")) {
            Write-Host "Attempting to remove existing item at: $fullJunctionPath (Action: $actionToTake)"
            Remove-Item -Path $fullJunctionPath -Recurse -Force -ErrorAction Stop; $operationMessage = "Replaced existing item and "
        }
        Write-Host "Attempting to create junction: '$fullJunctionPath' -> '$contentPath' using New-Item."
        New-Item -ItemType Junction -Path $fullJunctionPath -Value $contentPath -Force -ErrorAction Stop | Out-Null
        
        $logEntry = "$contentPath`t$fullJunctionPath"
        $TextBoxLogAndInput.AppendText("$logEntry`r`n")

        $successMessage = "$($operationMessage)Junction successfully created!`n`nLink: {0}`nTarget: {1}" -f $fullJunctionPath, $contentPath
        Show-MessageBox -Message $successMessage -Title "Success" -Image Information
    } catch {
        $errorMessage = "An error occurred: $($_.Exception.Message)"; if ($_.Exception.InnerException) { $errorMessage += "`nInner Exception: $($_.Exception.InnerException.Message)" }
        if ($_.Exception.Message -match "symbolic link|not have sufficient privilege|privilege not held|0x80070522") {
             $errorMessage += "`n`nThis typically means PowerShell does not have sufficient privileges. Try running the script as an Administrator, or ensure Developer Mode is enabled on your system (for Windows 10/11)."
        }
        Show-MessageBox -Message $errorMessage -Title "Operation Failed" -Image Error; Write-Error $_
    }
})

# --- Tab 3 Event Handlers ---
$ButtonMakeJunctionsFromText.Add_Click({
    $processingIssues = [System.Collections.Generic.List[string]]::new()
    $junctionsCreatedCount = 0
    $lines = $TextBoxLogAndInput.Text.Split([environment]::NewLine, [System.StringSplitOptions]::RemoveEmptyEntries)

    foreach ($line in $lines) {
        $trimmedLine = $line.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmedLine)) { continue }

        $parts = $trimmedLine.Split("`t") 
        if ($parts.Count -ne 2) {
            $processingIssues.Add("Skipped - Invalid format (must be 2 paths tab-separated): `"$trimmedLine`"")
            continue
        }

        $currentContentPath = $parts[0].Trim()
        $currentFullJunctionPath = $parts[1].Trim()

        # Validation for each pair
        if ([string]::IsNullOrWhiteSpace($currentContentPath) -or [string]::IsNullOrWhiteSpace($currentFullJunctionPath)) {
            $processingIssues.Add("Skipped - Empty content or link path after splitting: `"$trimmedLine`"")
            continue
        }
        # UNC Path Check for "From Text" tab
        if ($currentContentPath.StartsWith("\\") -or $currentFullJunctionPath.StartsWith("\\")) {
            $processingIssues.Add("Skipped - Network path not supported: Content=`"$currentContentPath`", Link=`"$currentFullJunctionPath`"")
            continue
        }
        if (-not (Test-Path -Path $currentContentPath -PathType Container)) {
            $processingIssues.Add("Skipped - Content path invalid: `"$currentContentPath`"")
            continue
        }
        
        $linkParentDir = Split-Path -Path $currentFullJunctionPath -Parent
        $parentCheckFailed = $false
        if ([string]::IsNullOrEmpty($linkParentDir)) {
            if (-not (($currentFullJunctionPath -match '^[a-zA-Z]:\\$') -and (Test-Path -Path $currentFullJunctionPath -PathType Container))) {
                $parentCheckFailed = $true
            }
        } elseif (-not (Test-Path -Path $linkParentDir -PathType Container)) {
            $parentCheckFailed = $true
        }

        if ($parentCheckFailed) {
            $processingIssues.Add("Skipped - Link path's parent directory is invalid or does not exist for '$currentFullJunctionPath' (Parent evaluated as: '$linkParentDir')")
            continue
        }

        if ($currentFullJunctionPath -eq $currentContentPath) {
            $processingIssues.Add("Skipped - Paths are identical: `"$currentFullJunctionPath`"")
            continue
        }

        $actionToTake_FromText = "CreateNew_FromText" 
        $pathExists_FromText = Test-Path -Path $currentFullJunctionPath
        $isReparsePoint_FromText = $false
        $isContainer_FromText = $false
        $isEmpty_FromText = $false

        if ($pathExists_FromText) {
            $isReparsePoint_FromText = Test-IsReparsePoint -Path $currentFullJunctionPath
            if (-not $isReparsePoint_FromText) {
                $isContainer_FromText = Test-Path -Path $currentFullJunctionPath -PathType Container
                if ($isContainer_FromText) {
                    try { $isEmpty_FromText = ((Get-ChildItem -Path $currentFullJunctionPath -Force -ErrorAction Stop).Count -eq 0) } catch { $isEmpty_FromText = $false } 
                }
            }
        }
        
        if ($pathExists_FromText) {
            if ($isContainer_FromText -and $isEmpty_FromText -and (-not $isReparsePoint_FromText)) { 
                $actionToTake_FromText = "ReplaceEmptyFolder_FromText"
            } else {
                if ($isReparsePoint_FromText) {
                    $processingIssues.Add("Skipped (existing link): `"$currentFullJunctionPath`"")
                } elseif ($isContainer_FromText -and -not $isEmpty_FromText) {
                    $processingIssues.Add("Skipped (not empty): `"$currentFullJunctionPath`"")
                } elseif (-not $isContainer_FromText -and -not $isReparsePoint_FromText) { 
                    $processingIssues.Add("Skipped (is a file): `"$currentFullJunctionPath`"")
                } else { 
                     $processingIssues.Add("Skipped (Link path exists - unhandled state): `"$currentFullJunctionPath`"") 
                }
                continue 
            }
        } 
        
        try {
            if ($actionToTake_FromText -eq "ReplaceEmptyFolder_FromText") {
                Write-Host "[FromText] Attempting to remove empty folder at: $currentFullJunctionPath"
                Remove-Item -Path $currentFullJunctionPath -Recurse -Force -ErrorAction Stop
            }
            
            Write-Host "[FromText] Attempting to create junction: '$currentFullJunctionPath' -> '$currentContentPath' using New-Item."
            New-Item -ItemType Junction -Path $currentFullJunctionPath -Value $currentContentPath -Force -ErrorAction Stop | Out-Null
            $junctionsCreatedCount++
        } catch {
            $errMessage = "Error processing '$currentFullJunctionPath': $($_.Exception.Message)"
            if ($_.Exception.Message -match "symbolic link|not have sufficient privilege|privilege not held|0x80070522") {
                $errMessage += " (Suggest Admin/Developer Mode)"
            }
            $processingIssues.Add($errMessage)
            Write-Error "Error during 'From Text' processing for $currentFullJunctionPath : $_"
        }
    } 

    # Show Summary Message
    $summaryMessage = ""
    if ($junctionsCreatedCount -gt 0 -and $processingIssues.Count -eq 0) {
        $summaryMessage = "Successfully created $junctionsCreatedCount junction(s) from the list."
    } elseif ($junctionsCreatedCount -eq 0 -and $processingIssues.Count -eq 0) {
        $summaryMessage = "No junctions were created. The list may have been empty or all entries were processed without requiring action (e.g., skipped or already correct)."
    } else {
        $summaryMessage = "Junction processing complete."
        if ($junctionsCreatedCount -gt 0) {
            $summaryMessage += " $junctionsCreatedCount junction(s) created."
        }
        if ($processingIssues.Count -gt 0) {
            $issuesString = ($processingIssues | ForEach-Object { "- $_" }) -join "`n"
            $summaryMessage += "`n`nIssues encountered:`n$issuesString"
        }
    }
    Show-MessageBox -Message $summaryMessage -Title "From Text Processing Complete" -Image Information

})


#endregion

#region Show GUI
$Global:Window.ShowDialog() | Out-Null
#endregion