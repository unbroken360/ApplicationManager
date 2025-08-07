#############################################################################
# Author  : Stefan Schuh
#
# Created : 2020/11/04
# Modified : 2025/08/07
# 0.1 - Created WPF Version
# 0.2 - Change Uninstall Collection naming; Add Browse Button for App selction
# 0.3 - Added IntuneWinAppUtil.exe to Files Folder and as Parameter in AppCreation. Bugfix in CleanUpIntuneOutputFolder. Added ToolTips
# 0.4 - Install AAD Modules: Write "SchUseStrongCrypto" Registry value if NuGet is not registered. Add AAD group prefix
# 0.5 - Use AAD group prefix also for Pilot Group; Add ErrorHandler for TextBoxIntuneOutputFolder;
# 0.6 - Add cscript as commandline in case of .vbs; Bugfix for subfolders; Add ServiceUI to commandline if located in source folder and "Allow User to interact" isset
# 0.7 - Bugfi Add Connect-MSIntuneGraph -TenantID $TenantName additional to -TenantName
# 0.8 - New ErrorAction Stop (879); SourcePatch added to Description in Intune; Integration of Winget Application creation with PSADT; Change ProgressBar beahavior; Add ProgressBar to Config save action; Split App Creation Functions, for better future use; Change Intune App creation Function based on Cmdlet changes;
# 0.9.1 - Rebuild Azure Authentication, Rebranding
# 0.9.2 - Bugfix: RunInstallAs32Bit interpreted as System.String
# 0.9.3 - Bugfix: Some Variables are interpreted as System.String in global state; Bugfix: Uninstall deployment
# 0.9.4 - Bugfix: Uninstall Program not set; Open Logfile on UNC Paths
# 0.9.5 - Changed Azure authentication to Application auth (https://learn.microsoft.com/en-us/samples/microsoftgraph/powershell-intune-samples/important/); Enhanced Azure Authentication Security
# 0.9.6 - Update Drive-Connect behaviour, connect to SiteCode Drive only if Application should be created in ConfigMgr
# 0.10.0
# - Full compatibility with the latest PowerShell App Deployment Toolkit version 4.x.
# - Implement Winget PowerShell Module to support all Winget Apps
# - MSI Detection method is automatically used if MSI detected
# - Extracted core logic into a reusable function: Create-ApplicationObjects.
# - Parameters control behavior instead of referencing UI controls directly ($CreateInConfigMgr, $CreateInIntune, etc.).
# - Enables reusing the logic in multiple places like ButtonCreateClick and ButtonCreateWinGet.
# - Avoided redundant code and UI-specific dependencies inside logic functions.
# - Improved maintainability and testability of the script.
# 0.11.0 - Find release notes on GitHub: https://github.com/unbroken360/ApplicationManager

#
# Purpose : This script imports Application to ConfigMgr and Intune
# https://www.benecke.cloud/powershell-how-to-build-a-gui-with-visual-studio/
#############################################################################


[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @"

<Window Name="Window"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        Title="Application Management" Height="700" Width="856" ResizeMode="NoResize">
    <Grid Margin="0,0,0,29" Background="#FF0035" Width="846">
        <Grid.RowDefinitions>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <TabControl Name="TabControlMain" Margin="0,88,0,-29" Width="846">
            <TabItem Header="Application Import" Name="TabCreateApp" Margin="-2,0,-2,0">
                <Grid Background="#FFE5E5E5" Height="575" Margin="0,0,-4,0">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="69*"/>
                        <ColumnDefinition Width="73*"/>
                        <ColumnDefinition Width="255*"/>
                    </Grid.ColumnDefinitions>

                    <Grid Background="White" Grid.ColumnSpan="3" Margin="0,0,0,10">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="429*"/>
                            <RowDefinition Height="29*"/>
                            <RowDefinition Height="23*"/>
                            <RowDefinition Height="56*"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="12*"/>
                            <ColumnDefinition Width="7*"/>
                            <ColumnDefinition Width="21*"/>
                            <ColumnDefinition Width="94*"/>
                            <ColumnDefinition Width="13*"/>
                            <ColumnDefinition Width="155*"/>
                            <ColumnDefinition Width="400*"/>
                            <ColumnDefinition Width="142*"/>
                        </Grid.ColumnDefinitions>
                        <TextBox Name="TextBoxDDPackages" Grid.Column="5" HorizontalAlignment="Left" Margin="10,66,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" IsEnabled="False"/>
                        <ComboBox Name="DDPackages" Grid.Column="5" HorizontalAlignment="Left" Margin="10,68,0,0" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="22"/>
                        <Label Name="labelDDPackages" Content="Applicationfolder:" HorizontalAlignment="Center" Margin="0,66,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="3" Grid.Column="1"/>
                        <TextBox Name="TextBoxMSIPackage" Grid.Column="5" HorizontalAlignment="Left" Margin="10,100,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26"/>
                        <Label Name="MSILabel" Content="MSI:" HorizontalAlignment="Center" Margin="0,100,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="3" Grid.Column="1"/>
                        <Button Name="ButtonMSI" Content="Browse" Grid.Column="6" HorizontalAlignment="Left" Margin="372,100,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.748,-0.269" Height="26" Width="70" Grid.ColumnSpan="2"/>
                        <Button Name="ButtonMSIClear" Content="Clear" Grid.Column="7" HorizontalAlignment="Left" Margin="47,100,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.748,-0.269" Height="26" Width="70"/>
                        <TextBox Name="TextBoxAppName" Grid.Column="5" HorizontalAlignment="Left" Margin="12,141,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <TextBox Name="TextBoxPublisher" Grid.Column="5" HorizontalAlignment="Left" Margin="12,172,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <TextBox Name="TextBoxVersion" Grid.Column="5" HorizontalAlignment="Left" Margin="12,203,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <TextBox Name="TextBoxInstallProgram" Grid.Column="5" HorizontalAlignment="Left" Margin="12,234,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <TextBox Name="TextBoxUnInstallProgram" Grid.Column="5" HorizontalAlignment="Left" Margin="12,266,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <TextBox Name="TextBoxSourcePath" Grid.Column="5" HorizontalAlignment="Left" Margin="12,298,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <TextBox Name="TextBoxADGroup" Grid.Column="5" HorizontalAlignment="Left" Margin="12,329,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <TextBox Name="TextBoxCollection" Grid.Column="5" HorizontalAlignment="Left" Margin="12,360,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Grid.ColumnSpan="2" Height="26" RenderTransformOrigin="0.483,0.524"/>
                        <Label Name="ApplicationNameLabel" Content="Application Name:" HorizontalAlignment="Left" Margin="1,141,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="3" Grid.Column="1"/>
                        <Label Name="PublisherLabel" Content="Publisher:" HorizontalAlignment="Left" Margin="1,172,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="3" Grid.Column="1"/>
                        <Label Name="VersionLabel" Content="Version:" HorizontalAlignment="Left" Margin="1,202,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="3" Grid.Column="1"/>
                        <Label Name="InstallationProgramLabel" Content="Install Program:" HorizontalAlignment="Left" Margin="10,233,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="4"/>
                        <Label Name="UninstallationProgramLabel" Content="Uninstall Program:" HorizontalAlignment="Left" Margin="10,267,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="4"/>
                        <Label Name="ContentSourceLabel" Content="Content Source Path:" HorizontalAlignment="Left" Margin="10,299,0,0" VerticalAlignment="Top" Width="121" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="4"/>
                        <Label Name="ADGroupLabel" Content="AD Group:" HorizontalAlignment="Left" Margin="1,329,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="3" Grid.Column="1"/>
                        <Label Name="CollectionLabel" Content="Collection:" HorizontalAlignment="Left" Margin="1,360,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="3" Grid.Column="1"/>
                        <Label Name="LabelDescription" Content="Create Application in Microsoft Endpoint Manager - ConfigMgr and Intune" HorizontalAlignment="Left" Margin="1,10,0,0" VerticalAlignment="Top" Grid.ColumnSpan="6" Width="797" FontStyle="Italic" Grid.Column="2"/>
                        <Label Name="LabelOutput" HorizontalAlignment="Left" Margin="1,41,0,0" VerticalAlignment="Top" Grid.ColumnSpan="6" Width="797" FontStyle="Italic" Background="Transparent" Foreground="#FFFF0303" Visibility="Hidden" Grid.Column="2">

                        </Label>
                        <Button Name="ButtonBrowseApp" Content="Browse" Grid.Column="6" HorizontalAlignment="Left" Margin="372,66,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.748,-0.269" Height="26" Width="145" Grid.ColumnSpan="2"/>

                    </Grid>
                </Grid>
            </TabItem>
            <TabItem Header="Winget" Name="TabCreateWinget" IsEnabled="True">
                <Grid Background="White" Height="617" Margin="0,0,-4,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="429*"/>
                        <RowDefinition Height="29*"/>
                        <RowDefinition Height="23*"/>
                        <RowDefinition Height="56*"/>
                    </Grid.RowDefinitions>
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="147*"/>
                        <ColumnDefinition Width="155*"/>
                        <ColumnDefinition Width="400*"/>
                        <ColumnDefinition Width="142*"/>
                    </Grid.ColumnDefinitions>
                    <TextBox Name="TextBoxWingetSearch" Grid.Column="1" HorizontalAlignment="Left" Margin="10,49,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="445" Grid.ColumnSpan="2" Height="26"/>
                    <Label Name="LabelWingetSearch" Content="Search:" HorizontalAlignment="Left" Margin="39,49,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" Grid.ColumnSpan="2"/>
                    <Button Name="ButtonWingetSearch" Content="Search" Grid.Column="2" HorizontalAlignment="Left" Margin="327,49,0,0" VerticalAlignment="Top" RenderTransformOrigin="-0.748,-0.269" Height="26" Width="70" AutomationProperties.AccessKey="" IsDefault="True"/>
                    <Label Name="LabelWinGetDescription" Content="Create Winget App in Microsoft Endpoint Manager - ConfigMgr and Intune" HorizontalAlignment="Left" Margin="20,10,0,0" VerticalAlignment="Top" Grid.ColumnSpan="4" Width="797" FontStyle="Italic"/>
                    <Label Name="LabelOutput1" HorizontalAlignment="Left" Margin="20,34,0,0" VerticalAlignment="Top" Grid.ColumnSpan="4" Width="797" FontStyle="Italic" Background="Transparent" Foreground="#FFFF0303"/>

                    <DataGrid Name="dataGridWinget" Grid.Column="1" Margin="10,87,3,202" Grid.ColumnSpan="2" AutoGenerateColumns="True" IsReadOnly="True">
                    </DataGrid>
                    <TextBox Name="TextBoxWingetPreview" Grid.Column="1" HorizontalAlignment="Left" Margin="13,321,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="542" Grid.ColumnSpan="2" Height="45" Grid.RowSpan="3"/>

                </Grid>

            </TabItem>
            <TabItem Header="Application Retirement" HorizontalAlignment="Center" Height="22" VerticalAlignment="Center" Width="139" IsEnabled="False">
                <Grid Background="#FFE5E5E5"/>
            </TabItem>
            <TabItem Header="Application ClearUp" HorizontalAlignment="Center" Height="22" VerticalAlignment="Center" Width="127" IsEnabled="False">
                <Grid Background="#FFE5E5E5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="24*"/>
                        <ColumnDefinition Width="373*"/>
                    </Grid.ColumnDefinitions>
                </Grid>
            </TabItem>
            <TabItem Header="Application Migration" HorizontalAlignment="Center" Height="22" VerticalAlignment="Center" Width="137" FontWeight="Normal" IsEnabled="False">
                <Grid Background="#FFE5E5E5"/>
            </TabItem>
            <TabItem Header="CONFIG" Name="TabConfig" HorizontalAlignment="Center" VerticalAlignment="Bottom" FontWeight="Bold" Margin="0,0,0,1">
                <Grid Background="White" Height="547">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>

                    <TabControl Name="tabControl1" TabStripPlacement="Left" Margin="0,0,0,37" >
                        <TabItem Header="Global">
                            <Grid Background="White" Margin="0,4,0,-33">
                                <ScrollViewer Margin="0,0,0,26">
                                    <Grid Height="458" Width="699">
                                        <TextBox Name="TextBoxlogfile" HorizontalAlignment="Left" Margin="156,44,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxPackagefolderpath" HorizontalAlignment="Left" Margin="156,74,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="461" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxTestpackagestring" HorizontalAlignment="Left" Margin="156,104,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="Labellogfile" Content="Logfile:" HorizontalAlignment="Left" Margin="10,44,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelPackagefolderpath" Content="Sourcepath:" HorizontalAlignment="Left" Margin="10,74,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelTestpackagestring" Content="Testpackagestring:" HorizontalAlignment="Left" Margin="10,104,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBlock Name="textBlockGlobalConfig" HorizontalAlignment="Left" Margin="14,10,0,0" Text="Configuration Data for gerneral use" TextWrapping="Wrap" VerticalAlignment="Top"/>
                                        <TextBox Name="TextBoxPackagedelimiter" HorizontalAlignment="Left" Margin="156,165,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxRetiredpackagefolderpath" HorizontalAlignment="Left" Margin="156,195,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="461" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="Labelshowsummaryfile" Content="Show Summaryfile:" HorizontalAlignment="Left" Margin="10,135,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelPackagedelimiter" Content="Packagedelimiter:" HorizontalAlignment="Left" Margin="10,165,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelRetiredpackagefolderpath" Content="Retired package folderpath:" HorizontalAlignment="Left" Margin="10,197,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="11"/>
                                        <TextBox Name="TextBoxMaxLogSizeInKB" HorizontalAlignment="Left" Margin="156,287,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelShowonlynewpackages" Content="Show only new packages:" HorizontalAlignment="Left" Margin="10,257,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="11"/>
                                        <Label Name="LabelMaxLogSizeInKB" Content="Max LogSize (KB):" HorizontalAlignment="Left" Margin="10,287,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxMailserver" HorizontalAlignment="Left" Margin="156,318,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxMailfrom" HorizontalAlignment="Left" Margin="156,348,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxMailrecipients" HorizontalAlignment="Left" Margin="156,378,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="Labelmailserver" Content="Mailserver:" HorizontalAlignment="Left" Margin="10,318,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelMailfrom" Content="Mail from:" HorizontalAlignment="Left" Margin="10,348,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelMailrecipients" Content="Mail recipients:" HorizontalAlignment="Left" Margin="10,378,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <ComboBox Name="comboBoxShowsummaryfile" HorizontalAlignment="Left" Margin="156,134,0,0" VerticalAlignment="Top" Width="120" Height="26">
                                            <ComboBoxItem Content="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                        <ComboBox Name="comboBoxShowonlynewpackages" HorizontalAlignment="Left" Margin="156,257,0,0" VerticalAlignment="Top" Width="120" Height="26">
                                            <ComboBoxItem Content="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                        <Button Name="buttonPackagefolderpath" Content="..." HorizontalAlignment="Left" Margin="624,74,0,0" VerticalAlignment="Top" Height="26" Width="38" Foreground="White" Background="#FFA29A9A"/>
                                        <Button Name="buttonRetiredpackagefolderpath" Content="..." HorizontalAlignment="Left" Margin="624,195,0,0" VerticalAlignment="Top" Height="26" Width="38" Foreground="White" Background="#FFA29A9A"/>
                                        <TextBox Name="TextBoxPSADTTemplate" HorizontalAlignment="Left" Margin="156,226,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="461" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14" ToolTip="Patch to PSADT (https://psappdeploytoolkit.com/) for Winget"/>
                                        <Label Name="LabelPSADTTemplate" Content="PSADT-Template:" HorizontalAlignment="Left" Margin="10,228,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="11"/>
                                        <Button Name="buttonPSADTTemplate" Content="..." HorizontalAlignment="Left" Margin="624,226,0,0" VerticalAlignment="Top" Height="26" Width="38" Foreground="White" Background="#FFA29A9A"/>
                                        <Button Name="buttonInstallWingetModule" Content="Install Winget integration" HorizontalAlignment="Left" Margin="156,410,0,0" VerticalAlignment="Top" Height="27" Width="200" Foreground="White" Background="#FFA29A9A"/>

                                    </Grid>
                                </ScrollViewer>
                            </Grid>
                        </TabItem>
                        <TabItem Header="ConfigMgr">
                            <Grid Background="White" Margin="0,0,0,-29">
                                <ScrollViewer Margin="0,0,0,26" Visibility="Visible">
                                    <Grid Height="482" Width="700">
                                        <TextBox Name="TextBoxSiteCode" HorizontalAlignment="Left" Margin="156,44,0,50" TextWrapping="Wrap" VerticalAlignment="Top" Width="59" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxDPGroup" HorizontalAlignment="Left" Margin="156,74,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxApplicationFolderName" HorizontalAlignment="Left" Margin="156,104,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelSiteCode" Content="SiteCode:" HorizontalAlignment="Left" Margin="10,44,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelStandardDPGroup" Content="Standard DP-Group:" HorizontalAlignment="Left" Margin="10,74,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelApplicationFolderName" Content="Application Folder:" HorizontalAlignment="Left" Margin="10,104,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBlock Name="textBlockConfigMgrConfig" HorizontalAlignment="Left" Margin="14,10,0,0" Text="Configuration Data for gerneral use" TextWrapping="Wrap" VerticalAlignment="Top"/>
                                        <TextBox Name="TextBoxApplicationtestCollectionname" HorizontalAlignment="Left" Margin="156,135,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelApplicationtestCollectionname" Content="Test-Collection:" HorizontalAlignment="Left" Margin="10,135,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxCollectionFolderName" HorizontalAlignment="Left" Margin="156,166,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxCollectionUninstallFolderName" HorizontalAlignment="Left" Margin="156,198,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxDeviceLimitingCollection" HorizontalAlignment="Left" Margin="156,230,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelCollectionFolderName" Content="Collection Folder:" HorizontalAlignment="Left" Margin="10,166,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelCollectionUninstallFolderName" Content="Uninstall-Collection Folder:" HorizontalAlignment="Left" Margin="10,198,0,0" VerticalAlignment="Top" Width="140" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="11"/>
                                        <Label Name="LabelDeviceLimitingCollection" Content="Device Limiting Collection:" HorizontalAlignment="Left" Margin="10,229,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="11"/>
                                        <TextBox Name="TextBoxUserLimitingCollection" HorizontalAlignment="Left" Margin="156,261,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelUserLimitingCollection" Content="User Limiting Collection:" HorizontalAlignment="Left" Margin="10,261,0,0" VerticalAlignment="Top" Width="140" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="11"/>
                                        <Label Name="LabelRunInstallAs32Bit" Content="Run Install As 32Bit:" HorizontalAlignment="Left" Margin="10,291,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelAllowFallbackSourceLocation" Content="Allow Fallback Source Location:" HorizontalAlignment="Left" Margin="10,352,0,0" VerticalAlignment="Top" Width="155" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="10"/>
                                        <Label Name="LabelAllowInteractionDefault" Content="Allow Interaction:" HorizontalAlignment="Left" Margin="10,383,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelDownloadOnSlowNetwork" Content="Download On Slow Network:" HorizontalAlignment="Left" Margin="9,321,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="10"/>
                                        <ComboBox Name="comboBoxRunInstallAs32Bit" HorizontalAlignment="Left" Margin="156,292,0,0" VerticalAlignment="Top" Width="120" Height="26">
                                            <ComboBoxItem Content="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                        <ComboBox Name="comboBoxAllowInteractionDefault" HorizontalAlignment="Left" Margin="156,386,0,0" VerticalAlignment="Top" Width="120" Height="26">
                                            <ComboBoxItem Content="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                        <ComboBox Name="comboBoxDownloadOnSlowNetwork" HorizontalAlignment="Left" Margin="156,322,0,0" VerticalAlignment="Top" Width="120" Height="26">
                                            <ComboBoxItem Content="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                        <ComboBox Name="comboBoxAllowFallbackSourceLocation" HorizontalAlignment="Left" Margin="156,353,0,0" VerticalAlignment="Top" Width="120" Height="26">
                                            <ComboBoxItem Content="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                    </Grid>
                                </ScrollViewer>
                            </Grid>
                        </TabItem>
                        <TabItem Header="Azure / Intune" HorizontalAlignment="Left" Width="103" Margin="-1,0,0,0">
                            <Grid Background="White" Margin="0,0,0,-29">
                                <ScrollViewer Margin="0,0,0,24">
                                    <Grid Height="483" Width="700">
                                        <TextBlock Name="textBlockCloudConfig" HorizontalAlignment="Left" Margin="14,10,0,0" Text="Configuration Data for Azure connectivity" TextWrapping="Wrap" VerticalAlignment="Top"/>
                                        <TextBox Name="TextBoxTenantName" HorizontalAlignment="Left" Margin="156,54,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelTenantName" Content="TenantName:" HorizontalAlignment="Left" Margin="10,54,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxAADUser" HorizontalAlignment="Left" Margin="156,84,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14" ToolTip="Example: Username@Tenant.com"/>
                                        <Label Name="LabelAADUser" Content="AAD-User:" HorizontalAlignment="Left" Margin="10,84,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxIntuneOutputFolder" HorizontalAlignment="Left" Margin="156,175,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="461" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14" ToolTip="Only local paths allowed, no UNC path"/>
                                        <Label Name="LabelIntuneOutputFolder" Content="Intune Output-Folder:" HorizontalAlignment="Left" Margin="10,175,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" ToolTip="Only local paths allowed, no UNC path"/>
                                        <Button Name="buttonIntuneOutputFolder" Content="..." HorizontalAlignment="Left" Margin="624,175,0,0" VerticalAlignment="Top" Height="26" Width="38" Foreground="White" Background="#FFA29A9A"/>
                                        <ComboBox Name="comboBoxCleanUpIntuneOutputFolder" HorizontalAlignment="Left" Margin="156,206,0,0" VerticalAlignment="Top" Width="120" Height="26" ToolTip="Remove .intunewin file after successful App creation">
                                            <ComboBoxItem Content="True" IsSelected="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                        <TextBox Name="TextBoxAADGroupNamePrefix" HorizontalAlignment="Left" Margin="156,272,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14" ToolTip="Prefix for AAD-Group (e.g. GRP-MEM- results in GRP-MEM-Microsoft Teams-Install / -Uninstall)"/>
                                        <Label Name="LabelAADGroupNamePrefix" Content="Pilot AAD-Group:" HorizontalAlignment="Left" Margin="10,239,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelCleanUpIntuneOutputFolder" Content="CleanUp Output Folder:" HorizontalAlignment="Left" Margin="10,207,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="11"/>
                                        <TextBox Name="TextBoxPilotAADGroup" HorizontalAlignment="Left" Margin="156,239,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelPilotAADGroup" Content="AAD GroupName Prefi" HorizontalAlignment="Left" Margin="10,273,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Button Name="buttonInstallIntuneModule" Content="Configure Intune integration" HorizontalAlignment="Center" Margin="0,355,0,0" VerticalAlignment="Top" Height="27" Width="250" Foreground="White" Background="#FFA29A9A"/>
                                        <TextBox Name="TextBoxAzApplicationName" HorizontalAlignment="Left" Margin="157,114,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14" ToolTip="For Azure authentication a Enterprise Application is required, specify the name for the Application witch will be created" Text="Intune_Cancom_AppManager"/>
                                        <Label Name="LabelAzApplicationName" Content="Azure AppName:" HorizontalAlignment="Left" Margin="11,114,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" ToolTip="Specify ClientSecret for EnterpriseApplication"/>
                                        <PasswordBox Name="TextBoxClientSecret" HorizontalAlignment="Left" Margin="157,144,0,0" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14" ToolTip="Example: Username@Tenant.com"/>
                                        <Label Name="LabelClientSecret" Content="ClientSecret:" HorizontalAlignment="Left" Margin="41,144,0,0" VerticalAlignment="Top" Width="77" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                    </Grid>
                                </ScrollViewer>
                            </Grid>
                        </TabItem>
                        <TabItem Header="ActiveDirectory" HorizontalAlignment="Center" Height="20" VerticalAlignment="Center" Width="103">
                            <Grid Background="White" Margin="0,0,0,-29">
                                <ScrollViewer Margin="0,0,0,24">
                                    <Grid Height="483" Width="700">
                                        <TextBlock Name="textBlockADConfig" HorizontalAlignment="Left" Margin="14,10,0,0" Text="Configuration Data for ActiveDirectory" TextWrapping="Wrap" VerticalAlignment="Top"/>
                                        <TextBox Name="TextBoxDomainNetbiosName" HorizontalAlignment="Left" Margin="156,79,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxDeviceOUPath" HorizontalAlignment="Left" Margin="156,109,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <TextBox Name="TextBoxDeviceUninstallOUPath" HorizontalAlignment="Left" Margin="156,139,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelDomainNetbiosName" Content="Netbios Name:" HorizontalAlignment="Left" Margin="10,79,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelDeviceOUPath" Content="Device OUPath:" HorizontalAlignment="Left" Margin="10,109,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <Label Name="LabelDeviceUninstallOUPath" Content="Device Uninstall OUPath:" HorizontalAlignment="Left" Margin="10,139,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxDeviceOURetiredPath" HorizontalAlignment="Left" Margin="156,170,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelDeviceOURetiredPath" Content="Device Retired OUPath:" HorizontalAlignment="Left" Margin="10,170,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxUserOUPath" HorizontalAlignment="Left" Margin="156,201,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelUserOUPath" Content="User OUPath:" HorizontalAlignment="Left" Margin="10,201,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <ComboBox Name="comboBoxCreateADGroup" HorizontalAlignment="Left" Margin="156,48,0,0" VerticalAlignment="Top" Width="120" Height="26">
                                            <ComboBoxItem Content="True"/>
                                            <ComboBoxItem Content="False"/>
                                        </ComboBox>
                                        <Label Name="LabelCreateADGroup" Content="Create AD-Group:" HorizontalAlignment="Left" Margin="10,49,0,0" VerticalAlignment="Top" Width="118" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxADGroupNamePrefix" HorizontalAlignment="Left" Margin="156,233,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelADGroupNamePrefix" Content="AD GroupName Prefi" HorizontalAlignment="Left" Margin="10,233,0,0" VerticalAlignment="Top" Width="146" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal"/>
                                        <TextBox Name="TextBoxADUninstallGroupNamePrefix" HorizontalAlignment="Left" Margin="156,264,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="506" Height="26" RenderTransformOrigin="0.483,0.524" FontWeight="Normal" FontSize="14"/>
                                        <Label Name="LabelADUninstallGroupNamePrefix" Content="AD Uninstall GroupName Prefi" HorizontalAlignment="Left" Margin="10,264,0,0" VerticalAlignment="Top" Width="141" AutomationProperties.HelpText="Application to be imported" FontStyle="Normal" Height="26" FontWeight="Normal" FontSize="9"/>
                                    </Grid>
                                </ScrollViewer>
                            </Grid>
                        </TabItem>
                    </TabControl>

                    <Button Name="buttonConfigSave" Content="Save" HorizontalAlignment="Left" Margin="737,454,0,0" VerticalAlignment="Top" Background="#FFA29A9A" Foreground="#FF0A0808" Height="30" Width="60"/>


                </Grid>
            </TabItem>

        </TabControl>
        <Rectangle Name="RectangleProgressBar" HorizontalAlignment="Left"  Width="660" Height="76" Margin="20,522,0,0" VerticalAlignment="Top" Fill="White" Visibility="Hidden"/>
        <Label Name="labelProgressbar" HorizontalAlignment="Left" Margin="20,545,0,0" VerticalAlignment="Top" Width="660" Height="30" FontSize="14" HorizontalContentAlignment="Center" FontStyle="Italic" Padding="0,5,5,5" Visibility="Hidden" Panel.ZIndex="1" Content="">


        </Label>
        <ProgressBar Name="ProgessBarGlobal" Height="5" Width="660" Value="8" IsIndeterminate="True"  Opacity="0.9" Padding="0,0,0,0" BorderBrush="Transparent"  Visibility="Hidden" Margin="20,535,166,0" VerticalAlignment="Top" Background="Transparent">
            <ProgressBar.Foreground>
                <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                    <GradientStop Color="#00FFCB00"/>
                    <GradientStop Color="#00FFCB00" Offset="1"/>
                    <GradientStop Color="#7FFFCB00" Offset="0.49"/>
                </LinearGradientBrush>
            </ProgressBar.Foreground>
        </ProgressBar>
        <Label Name="label1" Foreground="#fff" Content="Application Management" HorizontalAlignment="Left" Height="44" Margin="10,-2,0,0" VerticalAlignment="Top" Width="405" FontSize="24" FontWeight="Bold"/>
        <Label Name="Header_small" Content="for Microsoft Endpoint Manager" HorizontalAlignment="Left" Height="28" Margin="10,31,0,0" VerticalAlignment="Top" Width="405" FontWeight="Bold" Foreground="#FF474646"/>
        <Button Name="button" Content="Button" HorizontalAlignment="Left" Margin="948,209,0,0" VerticalAlignment="Top" Width="0"/>
        <CheckBox Name="checkBox" Content="CheckBox" HorizontalAlignment="Left" Margin="883,352,0,0" VerticalAlignment="Top"/>
        <Image Name="CompanyLogo" Margin="753,3,13,522" Height="80" Width="80" Stretch="Fill" Source="logo.png"/>

        <GroupBox Name="BoxCommonControls" Header="Deployment Options" Margin="10,510,10,-17" Visibility="Visible">
            <Grid Margin="10">
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="*"/>
                    <!-- Target -->
                    <ColumnDefinition Width="*"/>
                    <!-- Purpose -->
                    <ColumnDefinition Width="*"/>
                    <!-- Options -->
                    <ColumnDefinition Width="*"/>
                    <!-- Distribution -->
                    <ColumnDefinition Width="*"/>
                    <!-- Create In -->
                    <ColumnDefinition Width="Auto"/>
                    <!-- Buttons -->
                </Grid.ColumnDefinitions>

                <!-- Target Column -->
                <StackPanel Grid.Column="0" VerticalAlignment="Top">
                    <Label Content="Target:" FontWeight="Bold" Margin="0,0,0,5"/>
                    <RadioButton Name="RadioButtonDevice" Content="Device" GroupName="Target" Margin="0,2"/>
                    <RadioButton Name="RadioButtonUser" Content="User" GroupName="Target" Margin="0,2"/>
                </StackPanel>

                <!-- Purpose Column -->
                <StackPanel Grid.Column="1" VerticalAlignment="Top">
                    <Label Content="Purpose:" FontWeight="Bold" Margin="0,0,0,5"/>
                    <RadioButton Name="RadioButtonRequired" Content="Required" GroupName="Purpose" IsChecked="True" Margin="0,2"/>
                    <RadioButton Name="RadioButtonAvailable" Content="Available" GroupName="Purpose" Margin="0,2"/>
                </StackPanel>

                <!-- Options Column -->
                <StackPanel Grid.Column="2" VerticalAlignment="Top">
                    <Label Content="Options:" FontWeight="Bold" Margin="0,0,0,5"/>
                    <CheckBox Name="checkboxInteraction" Content="Allow User to interact" Margin="0,2"/>
                    <CheckBox Name="CheckBoxCreateCollection" Content="Create Collection" Margin="0,2"/>
                    <CheckBox Name="CheckBoxCreateADGroup" Content="Create AD Group" Margin="0,2"/>
                </StackPanel>

                <!-- Distribution Column -->
                <StackPanel Grid.Column="3" VerticalAlignment="Top">
                    <Label Content="Distribution:" FontWeight="Bold" Margin="0,0,0,5"/>
                    <CheckBox Name="CheckBoxDistributeContent" Content="Distribute Content" Margin="0,2"/>
                    <Label Name="LabelSelectDP" Content="Select DP-Groups..." FontSize="10" Foreground="#FF172AEA" Margin="15,0,0,0"/>
                    <CheckBox Name="CheckBoxCreateDeployment" Content="Create Deployment" Margin="0,5,0,0"/>
                </StackPanel>

                <!-- Create In Column -->
                <StackPanel Grid.Column="4" VerticalAlignment="Top">
                    <Label Content="Create In:" FontWeight="Bold" Margin="0,0,0,5"/>
                    <CheckBox Name="checkboxCreateInConfigMgr" Content="ConfigMgr" Margin="0,2"/>
                    <CheckBox Name="checkboxCreateInIntune" Content="Intune" Margin="0,2"/>
                </StackPanel>

                <!-- Buttons Column (Aligned top) -->
                <Grid Grid.Column="5" VerticalAlignment="Top" HorizontalAlignment="Right" Margin="0,24,10,0" Width="100" Height="36">
                    <!-- Create WG Button -->
                    <Button Name="ButtonCreateWinGet"
                            Content="Create WG"
                         
                            Background="#FFE8E4E4"
                            Foreground="#FF0A0808"
                            FontWeight="SemiBold"
                            BorderThickness="0.5"
                            Visibility="Collapsed" />

                    <!-- Create Button (Default visible) -->
                    <Button Name="ButtonCreate"
                            Content="CREATE"
                            Background="#FFE8E4E4"
                            Foreground="#FF0A0808"
                            FontWeight="SemiBold"
                            BorderThickness="0.5"
                            Visibility="Visible" >
                   
                    </Button>
                </Grid>
            </Grid>
        </GroupBox>
    </Grid>
</Window>






"@
#Read XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml) 
try { $Form = [Windows.Markup.XamlReader]::Load( $reader ) }
catch { Write-Host "Unable to load Windows.Markup.XamlReader"; exit }
# Store Form Objects In PowerShell
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) }


#Set working directory for exe or ps1
if ($env:p2eincfilepath) {
    $workingdir = $env:p2eincfilepath
}
else {
    $workingdir = $PSScriptRoot
}



# here's the base64 string
$base64 = "iVBORw0KGgoAAAANSUhEUgAAAgAAAAIACAYAAAD0eNT6AABshUlEQVR42u39d3jd13Xn+7/XOeiFYAXYexeLSIkQ1LssF7nJshyXxHbi9Mlkksx9fjP3dyeTTMtNmZm0yaQnY1mOJduSbMnqVKUIHpJi750EQRIEiN5xzr5/rEOTkkgJJIFTP6/ngQFSJE3u88XZa6+99togIiIiIiIiIrnPNAQi2SFQBzAWKAqEwkCoAAov+iXFRqi6xG+NgFUAJcP4v+mD0PHhfw87BwwBg4Z1GBZP/qcBoM2o14slogBARIY5uRcD4wJhXCCM8UmbSiPMABtz0S+dCJQCRclgoPii/1YGTLh0AMBYoHwYf5UeoOUjfs2Z5GTfD5wPBgB6gebk1x0Q2pMBQwvQZVibYZ1Aq1Hfr1ddRAGASL6s3qcGwqzzK3cjTPJJ22YmJ/Wq5EdFctIuB6Ymf3xeNPnf7KLPF0/00cv8FSLJj4+SSH58mCEgJD/i7/u953/cBXQmv24Fuj0ooBto92AhHE/+nq6AnYgQOQOcBc4piyCiAEAkoyf1QIgGQtQIEbDkBGwlEGYGGG9QCcxKTuaTgWnJr6PJ1Xpp8ueil5jULzXJZ8/w+Mf5wCC87yMOnEx+3Q00JrMHzf7zoQ2sA+yYQXeAFgh9/meFBNiQEUkAcQULIgoAREZ9sgcKAsGMEAUrBJsfCNONMB5sRnKSr0yu3M+v5qclJ/pLfc9Z3g/tB3/cncwedAENni2gNfnjZgjNENmJBwVHIQyBxY3IUDKwSCgoEFEAIHLV4tQWAUXJyb4YrAbCtAAVRhiXTN3PT076Y4Ap+B59oUZvVAwCbclgYH/y66MQznhgEDnuPxeaIfRDZCBZoDhk1Cc0fCLvVaAhELms2cB1QDWwEMJCoNp83/78hK/JPnUKgUnJj4UXggJrSU78jcnMwSFgT/KjBd9e6NDwiSgDIHK5Ff8Y4FZgscHq5MR/PqU/ifcW40nm6sJPKpzGTzXshnDKiKwHawDOGPXdGiZRACCSPxM8JIvtjBAJfjb+foMFwBp8334Sns4fpxHLGb34scVmfPvgELAD7E1gKyT6gWBEh/yz6gdEAYBI1gvUFQRCMSRKwMYGuBlYZITFYDcnJ/sC/Fx9VCOWB4+EFwoOAn0eFISDYPsN+34gHA/es6DPsH6jPq4hEwUAItmz0i8FJiSP4M0FFkFYBjYfmIl3xCvF/7u+B/L9cfFAoBdvanQa2A3sBd6F0ATWbkRajPpBDZcoABDJvEm/EBgPlBqsAuqAecBy/DheeXLiV+GrfJghvG6gCziMFxBuBzZAOGlEGsC6dapAFACIpH/Sn2h+9G4WXrg3Cbgdb7JTgaf4IxotuUIhmRnow48btuC1A2uB4xBOQ6TBsHMKBkQBgEjqJv5xRhgDVhNgdXLFfwMwHe+NP1ajJKOgBy8kbAG2Aq8CuyCcNaJnjfoBDZEoABAZkSVYHYEQSXbgGwtUGUwM8DEjrAJbmvz50uSHivgkFRJ4zUAXfq9BDOx5CA1AI9iZZNviuBFRV0JRACByFQFAaSBMC4QFBncHmGe+pz8rOdlrwpdMcP5UQTvwNrARwhGwfUZkv1HfqyESBQAiw3o3rR1jsAxYgqf3b8b390WyxUngDeAHYFsNawc6dJJAFACIfHDSj+LNeGoM7gLuSQYAarUr2ZwZOAFswS812mPYq2AHjPqg4REFAJLvE38JsMzgFuAB/Oa86zUykmMGgSbgZWA7hANG5F2wM2o2JAoAJJ8m/SK8an+awQrgJvxzFb6vX6RRkhwTuFA82AccBzYAb4OtM+yEjhOKAgDJ1Uk/alAaYK7BvXiavxpv1jMuOfHreZT8+ZbwGwoPAuv8w2IQWoEBIzKgbQJRACDZPOmfb7lbYrAIeDC58r8RndUXOa8XrxXYgdcKbAHqDU5BpEuZAVEAINm22q8ArguwwmA+3pp3GZ7mF5EPCnizoRMQ6sFewY8VnjUifQoERAGAZPLEX2C+sp+Jt+O9D5iNp/qrNUIiw9YO7AfeAt4B2wnhbICOKLEhDY8oAJBMCwAmGXwZeAhYDExGzXpEriUj0IFfSLQHeCHA01FiZzU0ogBA0j3hG1BgUBLgFwy+npz0J2jiFxlRA/h1xfsC/MDgn4LXDsSjxFQsKAoAJGUTfwQow8/sP5yc+KfjBX8iMnoS+DHCowH+Er+Q6ATQq0BAFADIaE/+1eZH9+7Aj/Mtw6/gLdDoiKTMIHAG3xp4Bj85cMKInDXq+zQ8ogBARnLin2BwHd6t705gIV7RX6zREUmbfnxr4DRQD7wU4K0osRYNjSgAkGud+CvMe/LXAp/F2/RWoR79Ipn1rQqdyUzAD4BYgD1RYl0aGlEAIFc88QMrzZv2PIiv/ifg+/x6bkQyT8ALA1uAXcD3DFsLdtqoH9DwiAIAucw7Rx2BYIFQboSpYEuAXwBuA8q14hfJKoMQ9oE9Abxl2L4AyTbDljDqNUKiAEB+GgAUBkI1hDXAl/CWvbNRcZ9INmcEBoEDwEv4fQPvGtZg1A9qeEQBgJzv4HcD8HH8Wt46oFIjI5Iz+oHdwFrgVcM2gZ1VJkABgOT35H+9eee+e4E1+Pl+EclNg8AR4KkA/xIltlVDogBA8mvSLwZuTE78t+EV/trjF8kfHcAG4B+NyOvAWaNedwwoAJAcnvgtOdEvN/h1/FjfWI2MSF7qxxsJvQO8YlgswKkIG3TroAIAyaGJv8AIRWArgK8kV/wLkpO/ngGRfH578KODJ4F1wN8CW4zIoK4fVgAg2T/5lyRX/DfjBX73AhM1MiLyPs34aYEfAdshHI2wsVfDogBAsnDVD8w0WIGn+m/Db+or1+iIyGV0AofxjoLPga0z7IyyAQoAJLsm/+nAvzK4D5iBt++NaHRE5CMEoAs4BvYPEJ42rAmsx6jXjYMKACRDJ/4ioMpgfoBPGPwiUK2REZGrdATCG2AvAm8EaIoSi2tYFABIZk3+EWAO8FmDrwGzUIW/iFy7XuBYgL8Evm9Ys2FDaiKkAEAyY+KfZLAIeAR4GN/r12srIiMl4EWC6/DbBl8P0BglptoABQCSxgBgtsG/Ax7F9/lFREbTCeBZCH8RYeNuDYcCAEn9xD/VYDVwV3LlPw2IamREZJQNcuHI4ONGZL1R36lhUQAgoz/xFwIzzBv6fAmYCoxBFf4ikjoB6MZvGvyHAE9FiZ3UsCgAkNGZ+CPJVf5K8zP9n8b3/TXxi0i6DAF7gX/E2wkfAutWgaACABkhCdZEAjYX+EbyAp85QClK+YtIJqxP4BSwHb9l8FWD40Ykob4BCgDkGlb9BiUQZoB9Efg5YK5eNxHJQAHYBjwF/ATssGGtCgIyV4GGIKNNBO4FexB4AKjR5C8iGbygXInXJa0CHvdAgG4NjQIAubLV/wT8TP/P4/38CzUqIpIFQUA18HEIYwLB4tS+ECXWoaHJzBdLMm/yrwB+zeBXktG0Jn8RyTaDwD4ILwH/JcLGcxoSBQBy6Ukfw8qAeyD8DnAdurZXRLJbAm8lvNGI/FYg7AuEnigxjYwCAElO/lFgAn6Bz7eAWrQ9IyK5lQ14I8B3jfAicCbCxiENiwKAvBWoMwilgTAzeXvfV4HlmvxFJAcN4bcL/hj4NkT2GtanngHpo4kmjRIkqg1uBW43uBNv7KPXRERydb6ZDfYpoBDCcwnCu8BZDY0yAHklTm0l8BmDX05O/GM1+YtInmQC2oD1Ab4LPBMl1qNhUQYg3yb/bwE3AsUaFRHJo3lnInC3QXmAjji1b0aJ6UIhZQByW4I1NQF72ODXgAXoiJ+I5K8+YHOAx43wgwgbz2hIFADklEAdEIogTAiEL4L9W/x8v8ZfRLQuglMQ/siwJ8BawAZUHKgAIFcCgPJAqIXwq8DteEtfERG54AzwFthfGrbeqO/XkIwu1QCMsji1VYHEPcDPAp/ViIiIXFIN8AUIkQSBOLVbosTaNSzKAGTr5F9jcD/wTeAGYIxGRUTkQ7UDbwLfCfBilFibhkQZgGyb/McY3At8BVgDVGhUREQ+UhXeF2UQ6IxT+7YuE1IGICskWGPAdLCHgS/hnf1KNdYiIsMWkpmArcAzEJ4EGiNsDBoaBQCZOvmDV/f/O7BPAjOBqEZGROSqxIHjEP4J+N9AU4SNGpUREtEQjNjkHwXGAV8GewSYoclfROSaRP291B7x91bGBeoK/B4VuVaqARg5Y4GHwX4dHfMTERnJeWox2L+G0Ac8B5zGawTkGiiKukaBuuJAWATh5/BjfrNRZkVEZKQlgGMeANhfG7bXqNeVwgoA0jLxAxQHwu0QfhO4CyjXyIiIjKpu4Ptg/8uwbUC/ugZeHa1Ur15JIKyB8IvAHZr8RURSohz4JITf9PdgijQkygCkTJzasYbdAeHrwAOa/EVE0pIJeAnsjw1716jv05BcGRUBXtXkzwMQfh1YpclfRCRtmYD7ILQGQkGc2nejxLo0LMoAjObk/2m8u9/NQKVGRUQkbQJ+IuBt4PEAr6t1sDIAI/yE1QEUJkhcD3wDuAko0ciIiKR9ETvZMwEUAT1xat+IEtNNggoARmryD2UBbjT4JeAWVHQiIpJJQcBY4GPmWdmuQN0mYECnAz6cTgEMI0gKsATCvwEe0eQvIpKRQUAJfhz7q96bRQvc4QyaXH71XxwIiyH8e+CLGhERkazwBNh/MWy3mgUpALiayb8kkLgV+AXg81r5i4hkjQHgX8D+u2H7dETw0rQFcOnJvzQQVgPfxKv+NfmLiGSPIn/vDl8PJGZrOC5NeyTvk2ANgfgssC8BdwKlGhURkawzBvgMcDrBmnPoKmFlAD5i8jdgOtjDwP3AJLRNIiKSrfPbNOBzwM8A05Pv8aIMwCVNxNP+3wBm4HdRi4hIdioCVoFNhFAF/BVwVsOiDMB7xKktA74M9ohnATT5i4jkSBAwK/ne/uXke72g9DZxaiNAGXCvwR8Ds1FmREQk1wwBRwP8DvBylFiPMgBSCtQZ/GtN/iIiOasAmJ18r79ew5HHGYBAHQkSFXj7yH8D6F5pEZHcNwC8FuBPgHciRLrztWVwXq52k5f7VBp2O4Rv4Tf7KRsiIpL7ioD7DAJYP7AxUNebj0FAXk56gUQ0kFgN4Vfxy300+YuI5I8ocCuEnwskVgYSeVn0naf73aEaeABsDX57lEj2M4NoBCKR9322y0XCkEj45whgyV8bifif9YFfH/wjnoB43D8nEjCUgETc/xyR7FEJ3A2hATgGnFIAkOPi1E4APgE8BIzX94BktIhBcSGUlEBJERQXQWkRVlToXxcnPxdGobAAigqhIOqfz39EL5PgSgQYGvLP0Yj//vOfLxcAxBMwOAS9/f55cAj6BqC3D/oHYWAQBoagr5/QNwA9ff7RN+AfiYReU8kkU8E+CxyPU/t0lFhLXq0Z8m3yN/gs8DV8319Ff5IZK/eiAqgog6oKqCzDqipgTLn/uKoCqsphTAVUlPrPl5f61xWl/vvOBwHnV+9mF63muXwGIITkF8lf/9PPl/s9F2UBQvDgIZHwyb+3H7p6oKcfOruhoxvaOqG1A9q7oK0L2rv9644uwrlOaO+E7j4Yius5kHQZANYD3w6QV0FAXmQAkhX/44EHgJ8DbtTkLyldxf90Uo74Sr48OZFXlsHYChhfBTXjYeokrGY8TJsEkydA9TgoLfbfF7noz7j4a/uISTsl32QhGVAkkkFBMjA4/zke9wxASzs0NMGpZjh2BhrOwOlzHih0JgOGjh7PGsTjF4KMoP0FGTVFQC2+EdYRp/bFCJGOfCgKzJctgBLztP//BSwEivXMS0oUFsCksTBpnE/oNeNherVP8NNrsOnVMGuyBwFYcrVuF1btliVJOjv/d45+eA/NmgmwdI7XX5Oc3Hv7oKkNTjYRjjTC8dNw/IwHCQ1NcOYcNLf7doXI6CgFag1+HegFXgFy/grhPMkAJKbgqf/les5lVEUjnpKfOgmmTPCJfvFsbPYUmD8dZk6GiWN9nz4fvSfAMV9zVZb7x7xp2B2rfHuhoxuOnYI9RwmHGmD/cThxBk63+Ednj/86kZFTDKwAPhdI7AKO5MW3Y65K3vw0G/gW2C+hoj8ZjZXv2EqYOxWbNQXmTIHZU2H+DP966iRP88u16erxrMDRU3C4AQ6cIBxogJ2H4GSTZxJErt0QcA74rwH+NtfbBed0ABCndgzwDYP/HzBZz7aM6Cp/wXRs2XxYOBMWz4I5U2FGDYwbk/49+Zz+xk5AcxscPglb98P2g4T9x+Fgg28X9A9ojORabQvw/wfeiBLrVACQZQJ1xYFwO4Q/wPs+63Y/uboVfsQ8ZT91IiyZgy2dA4tmw9ypMHsKjB/jVfgF0Qvn72X0g4B43I8dtnbCybNeL3DsFCG22wOD46f9SGJQEaFcsUHgHbD/bNhbRn2/AoAskWBNAdgC4I+A+1HFv1zppF8Q9er7aZOwm66DNUth6Vwv4pswxo/kXXzsTtIY7V906qBvwLMAJ5tgxyGI7fKAoPGs/7ehITUskuEaAF4CfjvAwSixnCs6ybl3rji1BQaLIHwpue8/Sc+xDGvSLyzwo3mTxmFzp0JtctKfNx1m1sC4Sq3usylD0NrhhYOHTsK7+whb9sG+Y8kjhz3qPSDD0Qj8YYBngWNRYjl1FCWnAoBAnQUSNcCXgV8E5qPUv3yYSARKi7zZzrRquGU5dstKT+/Pmuz7+QV6hLI+GDhzzk8VbN1PiO2CDbsunCZQICAfngXYHuAx4EngdC5lAnImAPDJP4yDcD/wLeBWoETPr1xSQdQn/ZoJsGIetnoRLJkDC2b4Ub1StYrIOYmE1wscPwPbD8CW/YT1OzwwaOtS8aBcTh+wOcBfAc9EiXUpAMi8AKAgEO6C8P8ANwDlem7lsiaOxT5/NzxY541paib4cb2oUvw5LwSvB2jt8C2B198lfO8V7zWgYkG5tH7ge8C/jxA7qQAg04J7aicB/xV4FN3wJ+95ys0L9spLYUYN9rE6+PzdvrdfVe4X5kh+isehqxeOnyHU74SnXoPYbm9ENKjOg/IeJyD8APhPETaeUwCQOav/0gSJ/2jws+i8v1w88VeU+t7+PTfAgzdjNy3ztruXu/FO8jcrEE9AexfhzS3wzJuwbpvXDnT1KjMgPtXA6QB/AvxjlFjWBwE58Q6YoPZ+4Duo4l/Az+1XlsOsKXDXKuwzd8AtK/0SHpHh6OqBddsJP3oL3twCRxs9EBCBYwH+PErsTxQApFmc2mkGv4tf8auiv7ye+COe0l+1CG5YjN26AlYs8Mt4yku14pcrWFUk/Jri0y2w4yDhtc3w+ruw96hODcggsDbAz0ezvB4ga98RA3URoDSQ+BLwf+M9//UOn68Tf1kJLJqJ3bAEPncnrFroZ/qLi9WWV64tEOgbhLOt8JN1hO+vhT1HoanV6wckXzUA/y3AD4DmKLGsfBiy9p0xQW0ZcDfwb4FbAFVy5RsDykph0Uy4cSn2QC3csBgmjYeyYq34ZWQDgdZO2HME1m4mvL3NLyJqOqdbCfM3C7AD+FPgqUiW3heQde+QcWoxQhFQB/bvgXs0+efbit+gtARWzIebl2P3roGVC/ya3aJC5YFkFN+A4tDRA/uPE/7lJXjseWhu17jk6dMA/Aj4zwG2A0NRYln1DyjIvojFCoBFwOc9CNDknz8Tf8QL+WZOhjVL4FO3e5/+6dU6vy+pEY16S+i502DKRCjU208+Pw3AjcBngWbgBFl200QWBgCMDT75fx6o0jOYL99qEaiqhDuuh8/dhd1+vQcCmvgl5eu+hJ8KeHubNxOSfDYD+LpBDDiZzAooABgtgbAEv+Fvhp69PFFcCMvmYZ+8FT51O6ycr+Y9kj6tHd5CeO8x6B/UeMgM4NYAO4GjCgBGK/CmdinwCLBcz1yemDUFPlaHffkBuH6h9+8XSZdE8JbBT73uVw6rQZC4LxnE49T+7yixBgUAI77yrxsXSHwG+DigWSCXRcwn+hXzsUfug7tvgIUzdSufpF93D+GdHbDvOPT2azzkvJnApwzeilPbmC03BmZFABCoiwTia8A+kRxobfzmIjM/u79iPva5u3y/f/4Mv6QnopdcMmD1f+AEPLcOWtq1+pf3LFvwXjSfAVri1G6LEsv4/aEsCQBCDdh9wBJU9Z+7q/5p1djDd8NXH4S50z0LoCY+kin6BghPvOrdAAe09y8fUAl8zqAgwB8D+xUAXGvQzZoKCPcA9+JV/5oRcm3iryyHOVP9et6Hk7f0FemyHskwDWfgpQ3Q1qmxkMtlASYBtxisilN7PEqsTwHAVa/86ywQXwR8Ez/7X6BnLIcURGFsJdQuxb50P9xbC1MnalwkA9+MAry22QsAVfkvlxcFpuG1anvi1O7M5HqADJ9QQ0ky9V+HLvrJHZbs5Ddvmhf4ffYOWL3Y9/9FMlFrJ+FHb3rhn/b+5cNVAp8EWvGL6jK2WUTGBgBxaiOBxGSwh4EyPVM5pLQYblqK/ewn4IE6qB6nCn/JbOu2wbYDfieAyEdnASYa4W7gz+LUdmRqi+CMDAC86j8sAH4buEHPUw6t/MePgUfvw37ry97Jr1C7OpLB4n4JUPjuS3DmnMZDruQNbznwO8C/AjIycszQACBRhfdX/jI68pcbigth5QLs0fvgKw9CzQSNiWS+jm546jWo3wlDuv5XrkgE+ALw53Fqj0SJ9WfiXzDTVv9F+HnK+9G+fw58C5hfnnLXDdh//BZ863NQPV7jItmx+j/ZRHhpA5w8q/GQqzHG4DeBFXFqi5QB+MgAIF4Ddg+wEN9LkWwVjcCsKd7U56HbvNCvvFTH+yQ79PbDpj2+9z84pPGQq1EMfMagDb8nIKMiyYwJAAJ1JEiUAiuBBwCdB8vqyT8Ki2Zin74DfvYTF872i2TL6r+hCTbu9s+q/JerY0ANfpJtcZzariixXgUAl/i7GFwHfAm4CaX/s/RxNygphusXYL/wWbg/ebZf1/ZKNunpg5c3EN7cCn3q+S/XHARcBzxqcCpO7cFMORWQEe/KgTogVAK3JD/U8S8bRZJV/jcvw37+036+f0a1Jn/JLgG/6e/NLXCowX8scm0mArcCyw3KfM5TBuD8d1w0eJ//u4GpelaycfKP+Er/gZuwR+6FumVQValxkezTPwDrdxI27dWNfzKS5gCfBjsWCNuAtB8ryYgAIBCK8ar/5XjRhGSbMeXY1z4OX7gXls2FIt3ZJFmqocmP/Z1s0ljISKoCVkFYFWCXAoAkw2YFwr3AdD0j2bbyNygrwX77y/CNT/kRPzX3kSwWXt8M63eo8l9Gw1y8v81h4LW0v32nf/VfFwkkHsK3ALT6zyYFUVgyG/t/fx1+7RGYVq3JX7Lb4ZPwcgwOnNBYyGioBK43wu1xatN+zD0DqrPCMrA78PSIZAMzKCmCZfOwX30EPn83jK3QuEh2G4oTNuyEHYdgYEDjIaMYBNg9kP5KwLQGAHFqKwKJR4FV6Krf7FFeAjcswX71YXj0XqgZr+Y+kv1ONsHr73oNQEKl/zJqCj0LwCfi1KZ15ZTWAMDgFrDzTX80g2RF7FoGN13nx/w+Vgfjxmjyl+w3MAjbD3rXv94+jYeMtjLgVp8D0yctq+7k3scE4CFgGVCk5yEbVv6lcO8a7Msfg3tvhLFjvAhQJNs1NMHz6+HQSYjr0h9Jydx7A/BQnNqNUWKteZEBCNRhUAHchu+BqONfNigqhGVzsW9+Gh64yVf+mvwlF8QTsOsI4fV3/fY/Zf9l9Bk+D9YBS9P1l0jDFkAoAJYafBNYoecgC5QWQ90y7He+CnevhqoKpf0ld7S0w3Nvw4HjMDio8ZBUWmFwf5za8XFqUz4fpyEDkCiAsAKoRan/LFj5F3jB37/6IjxQBxVlGhPJHUNxONRAeLHev9bqX1L8Dgs8CNxMGrbkUxoAxKm1gI0HWw5M0muf4cpK4OYV2O99Cz53F4zR5C85pvEs4anX4egpjYWky00GHwcmxqlNaWo1pQGAeeJ4MXCPXvMMV1oMtyzHfvNLcOdqXegjubn633UYnn5DYyHpdo/BYkvxabiUvqsHmBr82MM8vd4ZrKIM1izFfv2LcM8aTf6Sm9q7vPDvxBmNhaTbvAC3hBRfhpeyd/Y4tWZQZ/AI2vvPXGXFcOcq7Ld+Bu5do7S/5Kbefth7FF7aAP0q/JO0KzJ4xKAuldsAqVzaTcbvQ16g1zpDlRbDqkXYzzwAd97g5/5Fck0IcLKJ8MpGONroPxZJvwXJOXJyTgUAcWoLzNv93o0u/MlMxYXe2//rn4KP35I86qdhkRzUPwhbD8Czb0Fnj8ZDMuZdGLjbYFWqjgSmKgMwE2/8M5OMuIBI3iMahYUzsUfvg7tv9It9NPlLrq7+m1oJm/Z45X88oTGRTBFJzpFrSNEpuZRMxsnV/yfwqxAl04wpxx69Hx69H2ZNhohiNMlR8YRf+Vu/E9q6NB6SaSqBOw3m5EQAEKirBBYCM9CNfxm4+o9gv/RZn/ynTISCqMZEcldbJ2zYCYca/BigSGYpAOYDd8epHfUTAaMaAATqLBBfAtwFlOu1zSBmUFiAfe0T8BuPwtypOu4nuW1wyM/9v7oJTrdoPCRTTQG+YnD9aNcCjPI7fpgI9ilgJSr+y6DJH9/n/+St8H99DSZPVNpfcl9LO+G7LxF2HtTqXzI9CzANuA+ozsoAIFAXCYTFwKeAKr2mGaS0BGqvw37pczBnqgr+JPclErDnqO/9N7drPCTTFSYn/1FdOI9iABBKgHuB67T6z6TYMgpzpmJfuAduXg4l6skkeaC3n/DqRjh+2rcCRDJ8mQasNFgcp3bUCrNGMe8byoFP4l3/tMbMBJEIzJqCPXQ73HMjVKrLn+SJXYfhnR069y9Z826NbwN81Ubx9NxobvyOBVbodcwgFaXYPTfC5+7UcT/JH/EE4YV6OHBce/+STcYAn8ZP0GVHABCoI1Bn5gUMyi9nitJiuGMVfOMhWLXIm/+I5IMDx/3GP1X+S3aJAmMC/HKc2uho3BEwGkvAAqA6EL6u1y9DmMGNS7FvfApuXAyFascgeSAEGBgkPPGKzv1L9r59w+eBpUBJFmQAEpWBxH1guvQnU0yZiH3+LrjteijQ5C95YmAIDp6At7dp71+yWSlwCzDiRVsjGgAE6grA5gNfR21/M8O4Sp/8P3krTKxSOabkyeofaOskPLnWL/7RjX+SvUoM1ljmBwCJcgg34Uf/tNRMt2gUblrmbX5n1KjoT/LH0BAcOumr/xad+5esVgzcHmBZnNoRPVI/YjNCoA6gAqgFJug1SzMzWDAD+9ZnYPUiv+5XJF909cI722HHQW8CJJLdphl8ymB8nNpMzACEErBZ+HWGqv5P9+RfWQafucMr/8tK/OdE8kE8AUcbYf0Odf2TXFEGrAYWjeT8OiIBgF/6EyZDeBCYq9cq3Y9KCdQtw37hMzBxrMZD8kt7F7yznbB+O8RV+S+5sawDJge/WXfEtgFGKgNQAMwGbgUm67VKo4jB7CnYVx6EudM0HpJfQoB9x+DNrdDUqvGQXDLJYInBmJHqCTBCAUCIApOSH9psTluMaFBV4W1+713jwYBIPukbgPqdhPqdvhUgkjvKgTVAHSN0ImCEtgASpcDiZAAg6VIQhZmTsc/dCZNVhyl5qLEZYrvgZJPGQnLR9cCD5gX36Q8AAjcVQOQG4E4FAGle/S+ejX3rs8lWvzryJ3mmfxBiu7T6l1zPAswHmxWou+Y3+REIABLFEL6UjEzUYD4tkz9QPQ771K3w2TthrHowSR7avAeefUs9/yXXzYCwNEHimnvtjMAyMVLpEQlj9bqkSWkJ3LAE7r5RVf+Sn7p6CRt2ETbu8RbAIrlrOl4HcM0F9yMQAITVwFSt/tO1+k9W/T96H6xcAEWqwZQ8EwLsPw5vboGTZ9X4R3JdMbAEWJkBAQA3o73/9CkqgJuug1tXwAT1+pc81N3n6f9dh/0UgEjuq7ER6LlzTXsIgbrKQGI+Xpgg6TCtGnvodpgyUYV/cqlvUv+foTj0DkBnN/T1+0TZPwiDgzAY9975AV89h+D3RhQUQMH5z1EPNouLoLTIt51KivzH0Qhg6Qs+DzbAc+/AmXNa/Uu+mAncGqf2iSixU2kJACAsBmag9H96FER95b9mib8hSx5O8MEnvUTwyvf+Ae+D390LHd3+0dsHXX3Q1unNcTq7oaOH0N0Dvf0eCPQN+J8Vj/ufU5ic7IsK/B6JokKsrAQqSmFMhfebGFMGleX+c+Wl3oGyPPl1Vbn//kjE+1FEbHQuo+ob8Mr/dVv93yySH0rwtsCrgPQEAIHEbWALGOFbBWWY5k3HvngfTJ6o1H8+iid8Uj9xBk41w/Hk59PnCGdb/Sx8Y7P/mu6+a14dh8sFoZVlMGEsTBjjp1EmjYPpNTB5vPejmFYNUydC9biRr1E53UL4/lr1/Jd8Y8BkIywCfpLyACBQNyaQmMUINSSQK1RYAJ+/C1Yu9DdhyQ89fV7o1ngWGprg8EnYc5RwpNFT4efaPRuQKkNxaO30j4MXBQoR80zBzBps/nSYNwPmT/drqWdNhqmTPEtwLZdUxROw85A3/hHJP5Vgi+LUVkeJXVXnq2sIABK3ATeg1r/psXw+9vA9MG2ixiLX9fb73fab9xB2HYY9R+H4aTjRBB1dmdn0JhE889DWSdh+0AOC0mLPVs2bBrOnwor52A2LYcEMGFd55VsEPX2Ex573y39E8k8BMCv5kdoAAD+HuJprriOQK2IGJUXYw3fBolmjs68q6dc3AIca4LXNhE174MAJOHoKmttgYDD7/j2J4NsQhxr8wwzGjyHMngLzp2PXL4TbVsKqxVA+jHqW3n54dy+8slHPiuSrKDDPvA/PVX0jXNXkHbhpQiDMZoQuJJArUFIENy6FW1cO741SskMIMDjkley7jxDe3ALbDsDeY57u7+3PvX9vS7t/7DxMeGcHvFgPt6yEO67HVi+CqkoojF56m+BsK+F7r2j1L/ksAtQAd8Wp3RAldjhFAUBiNdiNGv80rP6rx2OfvxuuX3Rt+6eSISvjhFftHzsNr28mvLrRV/uNzdDVc6HCP5f1D3g9Q2MzbDsIz7xBWDIbu3M1PHATTJvk2wfns129/bBxjwcM8bieIclnhcAcvCvg6AYAgTqAwuTZ/1ka+xQrLoTr5sLNy/wIlmTp6hc/d9/W5av7tZvgzS2E82n+nj5fIedbBiQeh9YOrx04dpqw5yi2YZdvDdy2EmZN8e+BplbC65u953/Q4yR5rQhYZrAsTu2mKLEr6oR1hQFAIpqc+Fei9H9qRcyb/ty5CpbM1uo/Ww3FoemcV+xv2AXrthO27oOGsxea8OR9gBQ8+7HrMOHQSdiwE1u3DR6og5k1cOQUvL3NGxqJ5LcoMAW4DXgOODlqAQCEKLAU7AaNe4oVFGA3LPLGP5WKvbJST5+nul+OEX6yDnYdgdPN3ohHLq2vHw6cIJxq8e2BedO84c/RU7m/NSIyfLPwC/lGJwAI1BFIFEKYDKFanWdSbEIV3LEals/X6j8bV/2tHfDWNp/439zigcD57nvy0bp6YNchP0FwPjAQkfNWGMwO1O0Dhoz60cgA2Di89eBUjXdq2eJZcOMSGKNrF7LK4BCcOEN48lX4p2fh6GkvetPEf+XiCbX7Fbm0sWDL8eOAw+4JMOwAIEEi2XqQZejsf2qVl8AtK7wAULJn4j9zzov7/vR7sG2/31OviV9ERkW4KUF4ATjLMMtjhz2RG8GASWDzNdApNm861F7nR6Ek8/UNwNZ9hH/4MTz9Jpxt1ZiIyGhbaYQafH9+pAOAaFkgMRVQ79mUrv5L4bp5sHiWnwSQDA7Ag9+2t34H4fEX4fXN3iNfRGT0TQNbGmAdMKwOWVdQBJhYgB81UPo/lRbOxO6+AaZMUPFfJhsa8uNpP3yN8MybsP0g9PTqnLqIpEpRgOuAMSMeAECYBixR9X8qX84CbPk8qF0KpWr7m7H6B2DnYXjmTcLTr8P+4zraJyIpZ1ANYdiTxbACgECdBRLj8IYDkipTJ8Ety2HOVF35m6l6+rwpzd8/Q3hji/e2H1J7WhFJi8Vg4+PUHokS+8j84zCvkgvlwAy0/5860YgX/61YoOK/jJ78txL+/hnCC/Ve9a/JX0TSZzJ+Um/8cH7xsAKA5M1/K/GLByQVKsvg+gXe/7xAZRcZad12wh/8H3juHejo1niISLoVBFgBVI9YAABhHDBt+L9ertmUSX4bWmWZyi4yTXevr/z/53fhra1qTiMimaLQYKUNMwD4yKVloK48efvfVAUAKRKN+LG/21dBmYr/MkpPH7y8gfDnT8K67Ur5i0hGZQCAhcCEkcoAFON7/5VoLZoaleXw4M0wtsKDAckMLe3wnRcJ/++3feXfP6AxEZFMUwGsiVM77ZoDgEC8BlgOlGpcUyASgUljsbtWaywySWcPvLCe8M/PwrYD3upXRCQzswATgY+8Njby4ZN/XTR5AVANKgBM0eq/1C/9mTtNY5Ep+gchtstb+27cA726iU5EMlYZsMx83r6mDEAZMBs/WqBS9FQYN8Y7/+ncf2YYivuK//tr4d19MKAGPyKS0SLAWGBinNqia8gAhBJgEr7/r83oUX/ZIlAzHm5eobHIBCFA0zlv7/vcOmhTX38RyQqVwPzk56sOAKqAOQxjL0FGQGkRLJoN0ydpLDLBuQ5v7/vs23CqWeMhItmiCG8G9KFd5D40rW+EarDr8csFZLRVVWC3rlDf/0zQ3QvfX0v48yfg0Ekd9xORbFIBzDHC1WUAkgWANckMgE4AjLaIwfgqWDEfClVukVZDcdh5iPDYC7DnqPb9RSTblAIzwCb6XH6FAYD/ATYtmUqQ0VZSDMvmwrRJHgxIesQTsPcY4W+ehvXbNR4ikq3GAyvDh2QBCi6fAUiUANMVAKRIeQm2Yj6MKddYpFN7F+HHb8FTr3swICKSncqAqYFQfBUZgFAJLEDp/9EXjcKEKlg8W61/02koDjsOwdNv6HIfEcl2xcBEIxRdUQAQqAOsGD8CqA3p0VZaDLOnwrRq7f+nSwhwroPw4nrYdUirfxHJdmN8EW/lPqcPPwNQADYHmII6AI6+qnJs9SLvASDp0dwOr23y1X+POv2JSPYvLYFZYDOB6BUFAOYdAJWPHm1mMLbSq/8nVmk80mFwCHYc9Kr/I42eDRARyfLZBaiAsCQQCi+z0v+gQKIIuAE1ABp90ShMGgfTq30rQFIrBDjbRnhnB2zd733/RURyQ0GA8RAiww4A8KKB68CUARhtxYUwazJMGOutgCW1BoZg71F4fTOcOafVv4jkkiKDagjRKwgArADvI6wl6WgrKcLmToNxlRqLdGg6533+9x7TFb8ikmvKgMVgl9wCuNySsxC/TUhX0o0mM6iqgEWzYLy6Lafc4BDEdsGzb8PpFo2HiOSaKH4h0PADAMPm4QUEMqovTQTmToPJ43X8Lx1a2gnPr4dTLRBXr38RyUnlYMuGHQAEmKwxS4HCAmzZXC8ClNRKBHhnO6zbDt09Gg8RyVUVgbBi2AEAhKXKAKRAQRTmT/djgJJaff2EH77m1/wmVPgnIrm71DSYeCUBwBwFAClQXOQdAKsqNBaptvcYrN+p1b+I5HwGAPjoLYBAXbINMBM0ZikwpgwmT4AyHbZIqRAIazfCmRYYUstfEclppcDCS7UDfn8GIAqMBVuoDEAKzJ3mq3/TUKfU2TZ4fj30DmgsRCTXRZJBQNVHBQBFwDzUATA15kyFcvVaSrnYLjjUAAmt/kUkL5QEwrIPDQACoSAQJqLz/ylh86ZDhW5bTpkQoLef8OpGaGrVeIhIvigKXtv3Hu87fJ4oB24GUwZgtBUWwMzJUKoMQMr0D8LOQ7BlP/Tqxj8RyW+RSwQENVy2RbCMmAlV3v43qv7/KdM3ABt3w9FTGgsRyScFRqj5qADAuHx7YBlJsybDmHKNQ6okAjS3ETbu8ep/EZE8CgCA6kv95MXzfyEwVRmA0WfTa6BSOy0pMzQEx07B/uOeCRARyR+FYFMvmwEI1Bl++99kBQApzADoCGBqDAz55K/0v4jkaQYgUFeQnOvfGwAAUbBx6Arg1Jg2CSqUAUiJEKCjyyf/c+0aDxHJR6X4NkD0EhmAUID3C1YAMNqiEZherSOAqRJPwIETsO+Y0v8ikq/KAmFKcq7/QABQGAjVyShBRlN5qZ8CKCrUWKTCUBz2HyccadRYiEi+qgiE6YFQ+IEAIPl1Odr/H33V47wAUPv/qTE4BHuOggIAEclfBXiX38jFPwGAEYqBmWA6mzbaJlRBqXZaUqbpHBw5BZ26+U9E8lYU3+K/ZBEgyf+gZeloG1MBRUUah1RpbCY06ey/iOS1sUZYmlzsvzcDAHb+xiDdAzDKbEw5FGv/P2VONcOZcxoHEcnrpSfYHOCSNQAFwARAzelHW5UCgJRJBDjTCuc6NBYiks8i+I2/790CCNSd/7oYtQIefeMqoUQ1ACnR0Q2NZ6GrV2MhIgKWnPN/OtlH8OpAbUynJANQoQxAqpxqhtMt3gtARCS/FXHRSYDzAUAUbCJQqfEZZRGDqkooVqyVEg1nCI1nvRugiEh+q0zO9dGLAwBJlYICKCvWNcApEs62QYva/4qIfGA9qiFIseJCdQBMpdMtcLZVGQARESgPhHGBcCEDkCARhTAW3xuQ0VRY4B8y+uIJr/5v79ZYiIj4Uf8xgfCeGoBI8P1/3QOgACB3dPZ4ANA/qLEQEXmf8wGA4c0B1ARotBUVKgBIlY4uPwaY0AkAEZFLBgBGiBqMwy8DklENAJQBSJmuXuju1f6/iIiLJhf7BhdaAUcglIHpbNpoKyyAQiVaUhYA9PRpHEREXLkRxqFjgGlSUAAFCgBSIXQnAwBlAERESK7+S3hfIyBJWQAQhagCgJTo6/cCQAUAIiIfoAAgLXTjckr09nsQoPlfREQBQNqFAEFV6akLAAZQBCAiogAg/RJBKelUGRj0Dw23iMgH6DxayjMACQ8CZPQNxmEornEQEXElYBXnF//JAMAMbwOsS+qVAcgd8XiyCZDGW0QkOc+P5X3HAKOoEVCKMgAKAFIabMWD5n8RERfhoo6/F7cCLkA1AamZlLQFkKIMQEJtgEVEPiQakJQGAJqUUiZo9S8iogAgowIAzUqpebpNLRdERBQAZEoAEJQBSNnTHfEPERFRAJB2P61Ml1EXjfiHKQ0gIqIAIO0BgGoAUqYg6gGAiIgoAMiIACCuACAligv9Q0REFACkXUIBQMqUFENxkbYARESSS1DgpxeknA8AAtAHDGp8UpAB0CmAFGUAipIBgIZCRAToAlqTgcD5ACDEIbQC3RqfVAQAygCkRGlRcgtAEYCIiC/yQ9/5drQXZwAGgCGNzygb0gU1KVNeCmUlmv9FRC5BNQCp1j/gHzLqrLIcKspUAyAiogAgAwwOQU+fCgFTYWwljCn3joAiIqIAIO26+zwQkNFVVe4BgOkxFxFRAJAJehQApER5CYytUC8AEZHLBQABi4O1A70aEgUAOSMahepxMG6MCgFFRC6TAUgE6FAAkCK9/ToJkCrV42HSWBQBiIjQB9bJe/sAEC76kNHW1QMD6rmUCjZ9EkyeoPlfRAR6ArQFLHFxACApDQB6tQWQKlMmYZPH6yigiIiv/Id4Xytgkj+ps2kpENq6oE+9AFJi6kSYUQOFBRoLEcn7DIBBu2EXtgAMixumVsCp0tapACBVykqgZoIfBxQRyW99YJ128RZA8gc9QL/GJwVaO6BPQ50yUybApHEaBxGRi1y8BZBARYCp0dzuRwElNWZPwWZUaxxEJN8FLtrqPx8ADEFowo8Cymjr7YP2bhUCpsrkiTB9sl8NLCKSvzqSc/3QTwMAoz7g6X/NSKmQCNDcpm2AVCkvgfnTfStA5FoVFvg9E1UVKi6VbDME9CfnfPT0pktLO/T0Q6WK00ZdQRQWzIBZk+HoKY2HXJuJVXD3jX7d9PEzcOA4nGyCfvX2kIyf/Lu5aKv/4gAgkfyPQwoMUuB0i+oAUhkALJ2DLZ5NWLddXRjl6hkwdRL26H2wcBZ0dENsF7y1lXDgOBxu9J8LKqeSjNMN4RQweIkAIAwAJ8G6gSqN1egKp1uw7j6PxdSjZnRFIp7+nzsNKsugtVNjIlcnGoWpk2DJHJg3zb9/F82Eu1ZjG3cTnlsHu496gN/VA0NDKq2WTBHH2/1/MAMQsAHghHkWQAHAaDvaCO1dKAJIkZIiWDIb5kyF1n0aD7k6ZSWwbJ7fMhlJ1lBXVfhW3oIZ2Cdvg/3HCWs3wcbdsP2gBwNq/S3pN4AX+icukQEA1AkwdY6f8VShpEZBgWcA5kyDbQcgrkddri4AsKVzPBC4WMT8lEl1EVSPw25eDgdPwDNvEt7cAgcboKHJt/20PSDp0R2wE1xqC8CweCC0J6MEGW1n26ClzfejVUk8+qIRmDIRapfCO9vgVIvGRK6MmV8vvXgWlBR/9PO2aBb89pexr38Ktu4nvL0N3t4KW/ZDZ7eCUEm1PsNaSN4E+P4AYBBohqDKtFSIx+FMq1cOKwBIjdJi7NYVhB+9qQBArlxBFFYu8K6S0WHeoxaNetBwfy1212qI7YYX6gnb98PWA3C6GQZVlCop0WNYI5cuAiQOnEPtgFPndLP3Aqgo1VikQnEhXDfXPzbu0b6sXJmiAuzBm6HqKo7umnmgf+MSWDQTO3Ya3t1H2LLPTxHsPaZTQTLaeoGmS2YA8G6AbQoAUiecOIP1aLhTJhKBqgrs3lrCT97xPVmR4Zo1Be5Y5SdJroaZF6OWFMH4MbB4tgcUR07CzsPwykZCbBecbfWtQdUKyAhnAIBmLlUEaNSToLYfaABWAYUar1HW2KxugKlmBnXLYN50H/+E9mFlmI/ObSthQpWn9a9VNOqZv/ISmDYJViyAFfOxLftg817Cpj1w5hx09njLcAUDcm0GgVNG/Xve8N63+RyGgEawIQUAKXDyrBcDJYJXEUtqTK+Ge2701GuvAjAZhqJC+Fidfx7pgDRqMK4SblvpWwSfOIvtOQKb9hBejsHhk9DW5VtWCgTk6sQhNL//JwsuESWcLxLQxvRoa+/yKL+331cCkhoRw+6rJfzt08oCyPDMnQqLZ184+z8qKQaD0mK/t2L2FM8IXL+QsGU/bNgF+4/DuXbo7tUJArlSgUuc8Ct476+wIaDZLioSkFE0MOi96Vd1KQBItRXzYdUiX1l19Wg85MPdvNyr+VOVqCuIes3BjBrs9uvhQAPsOEB4KQZrN0GrLm6VKzIY/ATAh2YAJJWG4oQDJ7DWTj+jLqlTVox97eOEnYcVAMjlRSMwphy7r9ZvAEy1SAQmjvWiweXzsLtugEMN8NQbhGff9n4iQ7rEVT5Sr2HbPjQAiBDpBd4NJHQeJSUx2RDsOOzfxJL6N9Y7V8Pdq6G51YutRN5vTDk8fLcXjhZE0/u8VpTCwpkwfwbcshL7na8Q6nfAD17z5kLNbdA/4DVFIu8VN+zcR2UAeoDN+HlBSUEGgMMN0HTOv07nG0w+Gj8Ge/geQmw37D2qWwLlfav/KMyYjH3mTt+Tz5jg1bwXQVU5tmAGPHwP7D4Cj79IeGsrnDhzoWhQxI/9tQJ7PjQAMOqTv7pWG0yp0tvvl4V09/qlIpLCN/gILJkF998EbZ3qCyDvVVYMqxd6S9/RLP67FucLB6+bC7/1Zexzd8GGXYSNu/1WwiMn/T1GpwfyepaBcPj8/P5hGYCksAdsJbqmbvT1DyaPA/YoAEjHm+eUSdjHbyZsPwBNrVo1yYXgcOpEWHMdzKjO7L9rJOKXE5WVwOQJsHw+9tBtUL8Lnn2bsOeIZwV6+nR6ID8l2/wz3ADAjqB7alMUAAx4R0DdDJgexYV+TfBtK2HfMQ/GtFqSshK4fRVWu9Rv+csWhQXeU6CyDKrHw6qF2FY/Rhj2HfMCwtMtvvCQfNEFtmPYAUDwo4CSCgODHp2fa/fz6JmaaszlLEDNeOz26z1t2twGfboQM++fiSkTsfvWeHBoWfhuWBD1kwNjK72vwL1rsJ2HCG9sgXe2w4ET0NKuLoP5odew/cMOAAzbpKciRYbi3gxo33FYNh/GahsgLaum6xdiX3+IcPQU7Dmqxz+fVZZi998Et66E8izvhxYx/zcktwfstpXQ2ExYuwne3ubPekMTdHR5MKAtgpxcZuKXAA0vAIBwyBMBkhLtXYTtB7A7VysASJcJVX4kcNOt3npVWYD8deNS+OSt3jI6l7IaBVGoKIOFM/30wFc+Bpv2Eup3Qv0O2H7QgwGdhsm5DEAgHLmCAIB4MmKoRs2CRl9Xr18Hqu5e6TW2Evv4LYQXNsD2AxqPfH0GHrwZVi/K7X+nGVSWw903eJ3DPTcQNuyCN7f4VdmnW1QQmxuG8COAl4zqLnnw/HeZWgi2CpgOqEftaIv7a2N1y7zJh+oA0iMSgapKLBFg6wE/min5IxqB+2uxbz4Ec6dl597/1SgqhOnV2I1LsLrl2PK52Kwp/u9v7VQgkN36IGwGnvs9GvuHlQEIWBxoMo8eZLQlgh8DPHLKz+xe7X3jcu2qyv3Wt52H4NvP/zQ4kzwwo8ZX/3Om5mcQHo16w6MZ1XDzCuzuG2DjHsLajb49cK5DtTHZpwdsj/lRQIYVABg2BOEgoLtSU/Yy9RF2H8E6exQApDsLsGA69s2HCIdOwrptui0wH0wahz18D9xX66nxfBaNwpQJMGks1F6Hffxm75T56ibC+h1wsskXKvG4KsWyIwA4CHZFAcBAgO0QulA/gNToG4D9yTqAKRPyJ/2YiYqLYOUC7NF7CS1tfnZa56ZzV1Eh3LoCvnivr36j2oIjEoGiiI/NqkWweBYsnYvdcyPsPEjYfgh2HPQtgt4+3T+QmRIeAIRGPKv/wVjvUj/5ezSE32VaFLgVmIEKAUdfSIAZtnoRLJipN6FMCAKmV2NmcKTRWwUr/Zl7Cgtg8Szsd74Ct18PRUVa7nxgRWgeCEyZAEvnwMqF2OpF2Lzp2NhKn/zjCb+VUNmyDFv9sxt4IsKG5mEHAAC/y9QSsOuBJUCxxjJFUffMydgtK/yNSdKrogxqJsKpFjh6ylupSm5NbNOq4ZF7sa9+3Hvqa/L/6PeoilJvk7xolmfKFs/2AGFwyAsGh+IKBDJDB4StwCu/R2PbpX7BZWeZgA0atHOZ4wMyCrp7Ycch/1ymwxcZMUHMnoz9zAPQ0UV45k0FAbn02o4pgwdqffJX3c1VjF+5f8yowa6bCwdPEH70Fry8AY6d1hil3zmwLR4IXNplAwDD+iGc4TLVgzIKBoZ8pXnsNEwap/HIBIUFUHcdtHVguw4Tth/UmOTK6zpvOvazn4BFMzUe16K8xFsmL5jh1xBvO6AAIDP0gp027LJdzSIfEgB04vcHd6Jaz9QIAZpaCS/WaywySXER3LMGfvULsGyed1ST7FVUiN25Cvvdb3nXP/XdGBmRiN8xoGu1M2I28QxA2Ad0X0UAUB8gnAZ2JoMASYVz7fByzPfTJHNUlsGj92G/+ah3iVMQkL0r/9WL4Jc/D/et8X1/GRnNrbD7MJxt1VikXydwEEKrz+WXVvDhIYR1GzSifgCpc/52wEMNsHi2xiPTgoD767DOXkLvAOw5or7p2Tb5L53j5/3XLIWSIo3JSDqYvGpYp2UyQTdwPGBdH/aLPir31Qoc+LAUgoywRPBrOl/dpLHINNEoTJsED9+N/ZsvwYr5Sh9ni4jBsrnYrzwMn78Lpk7SazeS4gnCrsPQcFY9ATJDH3DGsN5rCQBa8DqALo1nCvX0+Q1dnT36Zsq4ICDik8fn78Z+7pMeEGg7IPPNnQ5fvA8+fTvMmqw+GyOttcOvNG9Vu+BMWEYCbUCzXaYD4LACgCixOHAWOIHfKSypMDTk6eUt+6FPuy8ZGQRUVcDPfAz7tS/A8vl+DloyjxnMnYb9xhf9OGfNeM/kyMjaecir/9UxMxMMAqcgNPAR2/cfGQYHzwIcSKYUJBUC0NBEeHurXxUsmWniWPjmQ/DVB2Hp7GQjGXWSyYyJH3895k+HL38MHr0fZk5W2n80xBPev2TvUV+8SCYEACeBtg8rABxWAIA3AzqOCgFTq7kN3tkOp5t1I10mTzKTxmHf+BT82iNeWKYGTpmhqAhWLcJ+56vYb3wRqscpOBstrR2EbQegsdmDAUm3XmCvET37Ub/wIwOAKLFW4IgCgDRE1UdP+b7agKLqjDZuDPbZO7Ff/QJ25yooL9WYpFNJEaxaiP3S5+CRe9VUa7QdOw17j6n9b2ZIAGeA08OZs4fbcP40cBCYwofcHyAj7Php2LIP7lztb2pawWSuiWO9wGxGDTb/JcK3X/CCKEmtshLsE7fANz8Nd6zyLnUyiguVOOw65Of/JUMCgHASOGHUf+T+8bACgACtyX4AcQUAKdTV64WA2w7AHdd7RzrJXKXFULsUZlRjUyYR/vqHnsWR1Jheg33qVviVh701rS7UGn3N7fDODm8BLJlgAOwtvHD/Iw33O+QksAv4NKBZKFVCIOw8jL21xbuXKQDIfAXJXgFf+Rg2aSzhiVdg3XbdmT5azLzuYtlc7NH74YE6WDhDk3+K3p84edaPLOvoX6boA2JGZFgrj+F+l3RA2AN2CFipMU6h081QvxMeOgnjxnhDE8lskYhfM/ul+7EbFsM//pjwQj0cO6VjUiM6zgaTJ8DD92BffdAvo6mqUKV/qnT2eIZy71GNReY4kWzhP6yavWF9p0SJATTjpwEklYbihIMNfsxGR2yya3IqL4Wlc+Bbn8X+9aNw60qoHq/GQSMxtlUV3tb3W5/FfvGzcP1CGFupyT+VGs8S3tyioDajhCagzxjehXLDzpMF7KyBQr20fKM1E2K7sQfrvAudZI+iQr9BcFo1tmAm4ZUYvLYZ9h9PdnpU5fSwWTKomlEDdddh998Ed67S90Q6DA7BzsPwmlqWZ9g3ySbg3HB/9ZVslJ0AtuJNBtT2LJX6B7zT1qGTMGWSnz+X7DKuEu5Yhc2fDotmEZ5+A97dC2fO6UKh4YhGYWIVXDcXe/BmeOAmmD9DVf7p0tXjz2/jWY1FBs0UwE7D2kc8ADBCD3AM7DgwT2OdYgeOw9b9nuqsLNN4ZKPiQpg1xe8RWLXIswE/eSd5hWqbCqkut+qvLIdFM+Gh27C7bvAK/7GV2kpJl3gCTrUQ3tqqxj+ZZS+ExkAY9oriSrYAAJoNDisASIO2Lli/A2qvg5uu03hkq4jB2ApYuQBbMAM+cyfhf/8QfvganGuH7j5tC5yf+EuKvInPLcuxb30W1iyBslJd5JNuHd2w67CK/zLP9oA1XclvGHYAECESAuG0nwbgfo11ig0OEV5/F5bMwZbNU+ozFwKBilJYOAP7w1+Hn/804Zk34OUYHDzhWwODeVj0WRCF8WM8vX/bSuwrH4NFs3QENlOEAPuOEX6w1mtYJFP0AG9HiDTit8mMbABg1JOgth3vB3AGqNGYp/gbr6UdYrs88r5hscYklya96+ZgC2fAQ7cTXtsEL9Z7e9XGs9CXBxdxFhf6CYkls+GWFdhn7/SvdctiZunu8+K/DbvUojyzJoiTwHajvm1UMgB4WBEHDhnsVgCQBkNx2HuUENuFLZypWoBcDAQWzsCmTIC65bD9IGzZS9h9xFOuLe251UwoGvHLlBbOhMWz/FrlNUu9TmJshZr5ZOIipPEsxHbCuQ7VrGQUO45f3HdlbzlX+Ovj+NXAO4C7NegplkjA6RZ49m24YYm3nZUc+h42T3UXF/kEuGQ2fOwm7Ogp7wOxbT9hxyEPBrp6s7NWIBrxhlZzpnqTpFWLvHPf7Kl+UqKsBAoKdNIlEyUr/8PmfdCta8ozzAbDTlzpb7rSACCBtwVeDzyCXw4kqdTbDxv3+FbA8nnef15yMBtQAJUFXicwZaJv+Rxbje08DNsOwonThONnvLvg+XqBRMIzBJmwMotEvM4hYn4178QqmF6NzZzsq/0FM3zynznZMx8FUV12lekam+GNLdDQpOr/zBEH2nxhbt2jGgAkOwLG49QeSjYFUgCQ8pc7Ac1thB+/hd1X68ej9MaZ21mB8xPkotmwYCZ88lY414E1NMHRRr8y+mgj4cAJOH4G2rs8IIgn3hsUjHRgYMm/n1lywo/4Cr+0GKZMgCmTsBnVsHCmT/QzamB6NVSP873980GCZL6+Adiyzzv/6ZbLjHplgLfw/f8r/ga/2k22Rvw44I2oKVDqhQCb9/pe3JQJ3hZVcl/EIJIMBspKfDKtXQqtndBwBg6ehBNnfJuopQ3OtEJTqwcEnd3esjUeTwYGAULiQk3BT2sLwoXZ3ey9k7wlV/SRCBREfBIvK/FalLGVfiXy5Al+dG/6JGzyRJg+CeZM9Q5+ClSz17l2n/xPNuXn6ZTM1QtsMiIHr+Y3X1UAECV2MkHt60AtMBddEZx6Le2E772CLZ/vqVTJ06AgAhOqYEIVtnKhT+Q9fT7pN7USzrb6ue2Obl+5NbV6T4mObt/T7e2HgQHoG/TAIB73ibqowIvwigq9Qr+w0Ff2FWUwYYyv4qsqffKvKMXGVvrP10yAMeVq0pNL4gk4cMKv/e3S3n8GSeDp/yNG/VWlZa66zDbACYNTwGwFAGny9ja/anbhTF9hiZzvL1BRCtMmXailC8HTuB3dfpSrp89bTPcP+iVTg0MePMQTvuo/v+1QUHDh66JCb85TWeaTfEmRLt/JBwODhLWbPLukvf9MMgTswU/lXZVrOWezCT8NsAZtA6RHRzfhX17G7lvjzVKUYpXLMfMVvIpG5UrtOOQNqtq7NBaZFwDsMyJ7rnq9cLW/MUqsBcJW4CxX0HlIRvqb8yD86C3fBxYRGUmDQ4S/fBL2HNHqP7MEoAnCbqO+P+UBgP8NbC9wWgFAGnX3El5Y790Be/s1HiIyclPMO9t9m1FtfzPx1TlkRNZfyx9yrRt47+LbAAN6PdIknoDdRwjPrfM9OnXnEpFrlUj4FuP3X/Puf1r9Z5oBoAHsUNoCgCixHmAr0KrXI41aO2HtZj8aqA5dInKtuvtg3Ta/gbRf67sM1ADsMOoH05kBIMAxoFuvRzpjwUHfAnh9M5xo0nWyInINq/8AZ1vhuXVw4Hhu3T+RO/aCrbvWP+SaAwDz0wD78OsIJV06ugnrtvs37TkVBIrI1a3o6OqBd3YQ3nhXe/+ZaRBvxLfnWv+gkcgAtAA/Ag7qdUln1O7NOsJjL/g9AUNxjYmIXOHUMgjbD8Azb8D+46opykwngG2QuOb93si1/wGRONh24IxelzQbGIRDDYSn34DmNo2HiFyZ1k54dh3hnR1q+ZuZ+oDtYBuN6DW/QCOwBVAfh3AQ70akWyLSrasHXt4AT76q6F1Ehi+egI27/Vhx41mNR2baBzwP4ZhRf83FXiPSxzNAJ/AGcESvTwY41Ux4/EU42KCxEJHhaWmDVzd633/JVA1gu4xI30j8YSPVyHsgwE7gHbwmQNJpYAj2HIV/elbtO0Xko/UPwobdhDe2qKFY5urB9/+bgBEp8hqRACBKLADtAfYC5/Q6pVkIfirg2be9i9fgkLYDROTy7xe7jxB+sg4OntB7ReZqBDZAOJ1RAUBSp3lToOMj9ZeTa/ym3neM8Pc/gk17FNWLyKXfJxqbCU++Aj95R43EMviVAg6DbTIiPUb9iPyhIxYARIn1AruAN/E7iiXd+ge9OdDjL+oqTxH5oL4Bv1b8xXov/FPTn0zVAxyC0OyF9yNjRC/zDtAZPAA4DGi2yQTnOgivbPSWnqoHEJHzEsl7RJ5+Aw43qndIZq/+TwExRrjr7ogGAFFiA+bHFF4ClEvKFAeTvQG2HVCbYBFxnb3wwnp4cT10qJt7BhsCtgExIzKiL1TBKIQqA8Bx87+0ZMTjMwSvbYaa8TBnKsyeojERyWeDQ7D9AOGfn/PmP5LJOoB/Bk6OxNn/UcsAJLUCj+NHFSRjHqFuwg9fh++8oCYfIvksEWDTHsLv/R0cUq+QzBfeNCIvGpERb7Q34gFAlFgiSqwrwGN64TJMayfhiVcIT74KbYr6RfJvLglwugW+vxY27FTRXxYwIt8z6geM+hF/sSKj+Pd+Ar+zWDJFPO4XfDy5lvDaZh0NFMmryR+/3e/VjV4T1KUyrUx/xwbaAmHTaP0fjGYAcDLA03oNM0z/AOw8BE+/AfuOeVAgIrmvpxe2HyQ88Qo0aIc2C3QD9RDasjEA6DF4Dm9dOKjXMoNWAV098NZWePwlONGkzl8iuW5wyI/8Pf6iNwbTTX/ZsPo/CTwJkVE7vz1qAUCUWDzAQWAjOhKYYY9WAhqaCI+/AM+/Az19CgJEclUiwJlzsHaTN/xpatX3e+ZLXvvL24YNZGMGAKAdvx9Am82ZuCJobCb80WPwSsyzAiKSe9o64YevE/7lZTh2Sr1AMt8QXj/3vGEnRqP4LyUBQJTYWbwp0Al0P0DmCQGOniL8j+/Cpr3eOlhEcmgqiUP9DsI/Pes3hKodeDbowPf+NxsbRjV7PtoZAALsTgYBOneWqUHA+p2Ev3gCNu3WZSAiuaR+J+EvnoTdR7wAWLJBM/CiET082v9Hox4AJLMArwNb0FZAZhoYhJdj8GdPwMY9ygSI5IKDJwj/7Z/grW2a/LNHP75o3mPUj/q+bEEq/kUBthm8CCwHivUaZ6DOHsLaTVjNeJg8HhbMhGhE4yKSjdq74Dsvwrptqu/Jsndi4K0AR1Pxf5aqd/im5C2BW1AtQOY610H44WvwN09rv1AkGyUScK4DHnve+/x3aPLPInFgS4B38AL63AgAosQS+GmAdXiBg2Tqm8epFsKP3oR/+jEcbVTFsEi2CAHOtsEzbxK+/TycOKPjftmlIzlH7osSS8kLl7Icb5RYK/A23hdAMjkIOHqK8L1XvFHQ6Ra9iYhkxfTRDW9sIfz1U7Blv58AkGyyEXg7OVemREo3eQPsxE8E9Om1zmDxBJxuITz5Cnz3JV9VKAgQyVy9/fDWVsJ3X4QdB72wV7JJH/BSco5MmZQGAAZn8RTHu3q9M9xQHPYcJfzvp+CJVzwToBWFSObp7oMNOwl/85Q39erR+ioLvQusS86RuRkABG9KuR94w38oGR8EHDlJ+KsfwFOvQ2unMgEimSQEv+Dn/zwPb2zRDX9Z+ioCbwTYHyClRVfRVP6f/T4n+V2mDQFmsBCYrtc+C95gWjpgy34MYOFMKC/xV1BE0mdwiLB+J/zRt+G5d6CrW2OSnfZAeMJgR5SNKU2zpvygd5TYoMEm4FX8ukPJdIkENJ4l/PFj8Pc/gqOndERQJJ0GBqF+J/zRY/DSBujsVk41O3UD3zeibxjRlBdupKXTS4Ce4I2BNun1z6ZHtZfw2Avwz8/BoQZdKSqSDvE4xHYT/vwJ2LBTXf6y26YAPwZaSHH6H1LUCfAS+oEYXhC4BijTc5ANbzwJONRA+N4rWEEUPn83LJoFBVGNjUhK3jkH4eAJb9b1k3e8+j+hpX8WCnizn3XAdqM+Lcc20pIBiBILUWJ9wPPAYT0LWWQoDkcbCY+/CP/4LOw8pEyASEom/wHYvIfwF08SfrLOL+5So65s1Q2sBZ6PEktbCietS7f/wLQzBnOA+cksgCrLsiUT0NoJh05iTa1eFDizRpkAkdFaK7Z2EN7eCn/9FDzzJrSqoWqWv6L7gf8Z4M3f52TazlcXpHMUosT649T+IHki4D50UVB2BQFnzhF+/BY0t2FFhXDLCiguhIguERIZEYkAbZ2E19+Fx56H1zZDm25Wz3L9wDYjUh+hPq035KZ9yfb7nDzxu0ydBFYLlOvZyDJ9A14QeOgkVj0OqsdBabGOCYqMhI5ueGE9/OWT8Ma7OuefG04BfxNhw/p0/0UyYqkWsPV4MUSLno0sFALEdhH+8DH40VvQ1KpjgiIjMfk/9Trhb5+Gjbs92JZs1wvsNZ/z0q4gE/4Shm2F8CTeGGiCnpEsNDgEG3cTWjuwfcfg0fthxXzVBYhcjaZW+O5LhP/1fTjSqELb3HEIeDoQTmbCXyYjMgCGDYFtwY8GntMzkqXicThwgvDY84S//iFsO+A/JyLD13gW/u5p7+1/+KQm/9xa/e8Ae9OIZEQ6JyOWZ/+R6QFCB9ADLAemoRMB2SkEP550qNFrA/qHsFmTobgIInpJRS77fdPRDfU7CX/+pBf8nTijC7hyaHkE7AUeM6werP/3aFAAAPB7NPAfmBYH+gwmAktQQWAWv5kBff1w/LT3CWjt9ALB8WMgqi0BkfdIJPy+jR++Tvijb8PajXCuQw1+cusd8SDwf4BnjchZoz4j/mIZ826cvCho0KAEuA6YQoZsUcjVxrwJ6OyBhjNYSwdUlUNlmbIBIucNDMKRU/DCesI//Mi3zXSdb67pwpvefTfA0QgbMmZPJ6OWY7/PyfjveiagCm8QVIW2ArJ/ddPmTYNoaMKiUagZDxWlOioo+b0m7Ov374snXib883M++Q8Mamxy75U+DvwdWH2UWEad4yzIwNE6CXwPGGPwLaBUz1AOfAu0dcJrmwnHT2OnmuHhu2F6DZSpZ4Dk2/dDgPYuWL+T8KM34ZWY37Cp/f5cNAC8BWw1LOOaOGTkO2+c2qjBbcATQLWeoRxi5hmAz9yBffljULvUtwQUBEi+6O6Fl2N+m9/mvX6Vr/b7c9VZw74KrDc2ZFwLx4x9141TOxn4JYN/h1oE51gQAJQUw9I52Bfvg4duh3nToKhQYyO5u+qPJ7ww9m+eIjy51r/Wqj+X9Qf4z4b9oWGDRn3GRXkZvexKsKYG7Pt4NkByUWkx3LzcswH3roHZUzQmklvicTh5FjbuIfzJd2DDTq3488NPIPxchI3NmfoXLMjk0QtYF/C8wWygBtASMdf09sPrmwkbd2Nf+wT8wqdh0SzdJyC5serv6IbdR+CZNwnffRFONPnPSy5L4JX/TxrRjL62MaMPZSdPBJxMHg1cil8ZLDn3RomnQo80elV0Tx9WXgJjKtRKWLJ0CkjAsdPwg9cI//BjeDkGp8/pjoz80A48DeFxI3I2Exr+ZGUA8PucDL/LtA7zrYobgUmoN0DurpZ6k82D9h+Hrl4sGoFxlVBaosOgkj16+mDrfl/1f/t52LTHT8Fo8s8Hg8BW4M+ALRFiGf2iZ8XbapzayQZfBb4JLMz0wEWuNSyNwtgKmD8De6AWbrseVi+CsZX+3xQMSCau+PsGYN8xeGsb4fl1sP2gX+qjQr98EQcOA38F9g8RNrRn+l+4IEsG9kyAx5K1ANOBSj1rufxtFIeWdmjtIOw7Ci+sxx6og4/VwbJ5MKYcokoESQYIwMCAp/tff5fw9BuwZR80t0N8yP+75ItO4BVgrWHt2fAXzooAIEosAKcT1L4D3ACsQkcD82BVFaCtCzbvI5w8C3uOYHfdAHethhk1UKlAQNI18Qfv2tfR7R38nltHeG0z7DysGzDz0wCwC3jBiBzOlF7/Hz23ZpH/wLRm84FeDEzQM5dHb7adPbD3GGzbD0dO+S7AjBrdKyDp0T8IOw/6BT5//yN49m1V+Oe3k8B3AjyVDan/87LunTNO7XSDPwC+omcuTxUUwMIZcNv12F2rPSNQM0GBgKTgDSgOBxoIL9XDKxvh3b1wqsVrACSfPQb8aYTYpqx6K83CgT4T4CXzGwOv13OXh4aG/Gz1oQbCC+thzRJ44CbsU7fDlAnqHyCj48QZv7TnlRjsOQqtnTA4pHGRrQG+GyGyK+vWUtn49g+8CpQDv4FvB0g+6h+EE2fgXDtsP0h4KQZ3rMI+eStMHu/bA+ojIFf9ThP3I31nzhHe3gZPvw7v7oOmczCgiV8AOAU8YbApQSLr7nHOyqVSnFozPw3wb4GfRw2CBKCwwI8KLp0Dd9+A1S2DxbNg4lgoK1FmQIantx+a22D3EcK67fDaJm9QdVZH+uQ92oB/Av4UwvEIG7NuHyhr3xHj1JYBdxj8FnAHOhUgFwcCVRUwvRpuX4nduRpWLoRJY6GyDCI6OSDvE4Kf429ugx0H/UjfKxvhxGlo71aqX96vH3ga7A+APRE29GfjP6Igi1+APiAG/B3eH2CBnkkB/M26uQ06uqDpHCG2G65fiF2/CJbPg1k1XjRYWKCxynfxhJ8wOXzSO1Bu2kPYuNvbUp8550f9RN73DgPsBr5r2D78ZFpWyvqcaJzaSoN/Bj6NOgTKpUQjUFEG1eNg8WzsrtVw52qYNsmbCpUU6wRBPgaJPX1wugV2HfYGPpv3+qTf3qVUv3yY0xD+OWD/JUqsM5v/ITnxrhen9gsG/zGZBSjS8ymXftoNigt90p82CW5Zgd24BK6bC1MmwvgxuoUwlyUCdPX4Xv7BBth+iPDaJth12DtP9vTpHL98lH7gBcN+z9iwJevfEnMkACgEftn8VMBcdGGQDCcYKCyAilJYOBPuXI3duQqWzIGJVV40qFqB3Fntd3T7ttCW/YRXYvDaZl/99w+qc58Mf6qBwwH+nwiR7xv1Wf/g5MxSJ8GaMuB3wH4DdQmUq/lOKC2G6xdh962B5fNh1mTPDEwcCyVKLGWVvgFP5zee9b39N7YQ1m3zvX6l9+XqHAUeh/BfI2zszpW3vZyRoHY+8N+BO4Exel7lqrMDpcWeGVg2D1sx348TLp0D06oVDGSqrl441exNovYeI2zdB1sPwNFGDwhErt5Z4HHD/g5sTy6s/nMuAIhTWwk8aH5t8B2oP4CMhMoymDsNaq+D6+bA7KnY9GqYOdm3C1QzkB4heDe+Qw2Eo6cuVPJv3ONfd/dqjGQk9AAvAn9oRDYZ9TlzJjTXzkF1A2sDhOSlQffiHQNFruGp6vXWr0cavWagejxhygSYXgPzp2PzpsOiWTCjWn0GRnvC7+713vsHGwgHjl94XRqafI+/p8/39XVuX0bGALAd+CGwB68DyBk5t3SJU2vARPPJ/z8B83Lx3ynp+o4xPzIYifjxwvFVMGsKLJiOzZoCM2tg6kTfKqgZ78cPC6IXfn3ElDH4MImEV+vHE/51dx80t8Kx0972ubGZ0HAGDjf6x+lm39OPJyAkQEX8MoIhJ3Acwh8DzwSsIXk1vQKADA8CMKgBfhv4FaBCz7KM3neQXThVMK4SJo3Fpk6E6vEweaIXE06vhtlTYPIEKC+9EAz8NDDIw6xBPHFhoo8nYGAAmlqhsdlX9CfO+P79qRbCybNe0NfR7c15QvC3Zx3bk9HTAfw18N8hnImwMecetpxdigTqCgNhDYT/BNwClOh5lpSKmGcApk2CyROwGTV+qqBmAkwY48HC+DEwcRyMLfdrjgsKoCAC0agHBtGIBxfZmjVIJHyFfn6VHk9+PTDo+/dN5+Bsu6fvz7bCybNwuoVwqsUL+prOKZ0v6TAE/ATsDwzbZNTnZEvInM5Fxqktx4sCfwm4CyjUcy1pVVjgjYiqKmBsBVZVAROqPBCYUOVZgwljfGthTDmMrfCMQVHhheCgIOoBQiSZRTifQTgfLIz6pB4upOrPr97j8QuTe/yiSb8z2XinpR3OdfhHc6tP/uc6CM3t0Jr8+fYur9bXql7SuXb0yX938Htm6qPEenL1H5vrzdC78erNAmApME3Pt6TV4JBPhi3tP3238e/EqDcfqiiD8hL/uqTIjyOWlXgvgskTfHuhZjyMG+MBQkUZVCU/jyn3Xz/a+gegs9v357t6fDI/20Y41+4p/HMdvqJv6/QAoLPHi/O6+5Kfe6Gv3wMIkcxb+TdC+HODdyJs7Mvlf2xeVCMFbpoWCP8V+BoqCJRsEzEoLPSAoLgISgr9x4UFHjhc/Lm40IOBilLPHJQUe9YgEvH/Fo36ry0u/GC2YCh+4bz84JCn6Qfj0JucvLu6/az9+Z//aTp/yIOCgUGvwD//eXDIMwQi2aMFeMywPzI2nMz1f2yeXIdmpwLhr82bA92Nf1YgINkhEXyC7R9GMxsD7KLTBmYXFSpyYdK/3FbB+fT7xUV25z/Op/xFcnGd6EV/Pwnw94adyouZMV9e3Ti1ZcBN5qcC7gHGKwgQEdHkD5zDe8j8BbAhSqw/H/7heXMhepRYT5za9QEKk2uhTwKlevZFRPJaH/BagL8DYvky+UOe3ZoXJdYHbAB+AOzAuzyJiEh+6sU7/f0gufLvy6d/fN51H4kSa4fwMvAvQKOefxGRvLUP+A6El31uyC8F+fiKG5HWAK9DWIXfFTAR1QOIiOSLADQDL4O9alhrPg5Cnt5aYgnDdoH9CfB9oFPfDyIieeMc8DjYPxi2Hywvj7fkaQagHmAgTu1e4AcGU4D70c2BIiK5rht4PsBjEI5E2JC3vaYL8vkpiBLrj1O7LkCLQRfwCFCs7w8RkZzUDzwV4L9Fie3O98HQvndSnNobDf4RWITuDBARyTWDwL4A34gS26ThyNsagEvaDuEPgW3JB0VERHJn8t+WfI/fruFQBuADEqwpAx4A+23gRnw7QGMkIpKdAl7kvRnCnwEvRdjYo2FxBRqC9+gBXgPGJR+c1agwUEQkW/UC6yH8Y/K9XZO/AoBLi7ARoD3Bmqf9RhXGA0vQVomISLaJAyeA58FeieRho5+PnvPkUoFAK7AW+AnQwEXXtouISMYLwFHgO2BrjUiHhkQBwLAZkRNgj4F9Dz86IiIi2aEf+B7YPxu2x6hXYbcCgCsJAOqHDNsD/DV+PDCuURERyXhD/p5tf29Yg1E/pCG53DwnHypQFwmEORB+G/hZVBQoIpKp2sD+wDz132jUJzQkCgCuSZzaAmCZwa8BXwIqNCoiIhkjAXQAf2pE/odRr4K/YdApgOHGAN5B6h+ACvN7A8ahLRQRkUyY/BsDPAf8ONnWXZQBGPFMQBFwu8GXgU8B1RoVEZG0CUATfrnPXwFbo8QGNCzKAIy4KLGBOLX1+D3S/cBDwHSNjIhIWpzEr3R/CtiuyV8ZgBSEnHXRQFgB/CKER/CGQRpLEZHUrfzPgT0B/HkgHNLkrwAg1UHAIuBXIHweqAGiGhkRkVEVB46CPQP8vWH7jHod01YAkJYgYAnwNQhfRtsBIiKjrQH4X2D/YthxTf4KANIZBBQFwjwIPwd8GlioTICIyKis/N/Fm/y8oMlfAUDGZAIg1AQSS8H+EFilURERGTH9wC4I/yrCxnc0HCND59hHJIqqjxsbGoFXIfwPPEWl3tMiItduENgO4Y+B9RqOkaNjgCOdDICnIAyB/TJwI1CKMi0iIlfzftoDbIbw34GX0c2sI7x4lRGXYE0V2IPA14CbgbEo2yIiMvy3UTgDvAX8EMILETaqva8CgOwQp7bC4C7gi8ADeNdAjbeIyEev/BuBnwR4DHg3SkztfUeBtgBGSZRYV5za14FugxLgHmCCRkZE5EO1AM8F+A4QixLr05AoA5CtmYBygxuAL+A3CVYBRRoZEZH3GADage8G+Cdgf5RYt4ZFAUBWC9QVQlgUCF8HPgPM09iLiFz0NgmHgGcM+0e8u9+QhkUBQC4EAECIQCgOhI+D/d/AMmUCREQTP73AXgj/xbDnwfrBEka9RkcBQO6JU3uzwX8G6oAyjYiI5KkOfL//z6PEdMY/xXQ0LT22Bfgz4LVk9Csikm9agdcC/A2wTcOhDEA+ZQHKgFsNvgrUAnOAYo2MiOS4fuAI8FKAZ4B1UWL9GhYFAHklwU1lwHIID+CnBBajugARyV1dydX+c2DPAQcibFAWNE3UByCt0Zf1QHg3EE4Cx8C+gDcPqtToiEiO6QBeDPB/DDYCLYap0l8ZgHzPBKyxgJUDtxn8PHAvMAZdKywi2S3gl/l0AM8G+B7e3rc7SkyjowBALnyn1BUGwmIIjwDfAiZrVEQki/UAu4Gnjcg/A2eMet2UqgBAPiQjMAb412CPAvNRcaCIZJ9+IBbgb4HvR4lpr18BgAxHnNoxBnfgWwIfw68VFhHJBgPAhgD/E3hRLX0VAMgVSm4JXA/hc/gpgXmod4OIZK4EsA/4X8ArAY5HifVoWBQAyNUFASWBRDVwfzIbsDyZDVCBoIhk2qr/DeBxCD+OsLFFQ6IAQEZAnNpKw26E8AngE8lsgGoDRCTd+oFmYHOAvwDWR4l1aVgyn/oAZIkIkS5gUyA0A0fxWwVvRXcJiEh6J/8NwPcCbAH2ANrvVwZARikTEAHGGlwP/BxeKDgDbQmISGon/kPATuAHAV6IEuvQsCgAkBQI1BUFEtOAXwS+DExTECAiKZr8j+KFfj8JcCZKrFPDogBAUizBmilgX8ZPCcwHxikQEJFR0Ac0ADuAlwM8GSXWrGHJXqoByPpMgJ0y7O+AnRAexk8LTNdrKyIjaAjYlby693ngNN7iV5QBkEwQp7YauMfgG8BqYKJGRUSu0X5gS4DvAGvV1EcBgGRsRqCuGML1gcSnwD4NrNCoiMhVOh7gPwFPRYnpXH+OUZo49wwA28GOg22B8A3gNvx2QXURFJGPMgQcBuoDrAOeBdo0LMoASPZkAiwQSiCxBOw+4G7gBnxbQK+7iHzwbcMb+uwIPuk/AzThV/cGDY8CAMkyCdZEjci4QFgFPIg3EJoBFOr1FxEgjmcOG4AfQ3gtYDuBY5r4FQBI9mcDCMQtYNMMPgk8ANQlswEKBETye/I/BDwHvB2g3ginAhaixDQ6CgAkhwIBAyYGEsvxLYE6vEhwkp4FkbwyhLfsPQ48E+BxYH+UWFxDowBAcjsQiEAYH0gsAPsa8HX8hkERyQ+nSe7zB9gGnIoSG9KwKACQPBKndr75nQIPAUuAIo2KSM7qAl4Bvhtgc5TYIQ2JAgDJ32wAEKoCYR5wF/Ap4EagHB0bFMkFCeAc8CrYWuCNQDiodL8oAJDzgYAFQhFQA+E+CJ8Hq8P7BxRqhESyUjuE+oC9DPzEsEOGDRr1qu4XBQDygUAgGkiMwe8TuAvCx8FuAao0OiLZNPFzNMCPDX4c4AjQHiU2oKERBQDykYEAJCoSsNTg54Gbk0FBpZ4bkQz9toVO/Dz/+gDfBnYDbVFiurhHFADIlYlTW2owG1iG9xC4C5gMFGt0RDJGB7AX2Ay8BuwMcDRKrFdDIwoAZAQCAbsOwpfw/gGzgQlAiUZHJG368Ha92/AVf32AZk38ogBARlSgriAQrwKbEKDWvKPgGmBmMhDQqQGR0ZcAevB9/a0BfmjwboCmKLE+DY8oAJDRzAZYcrKvBG5JBgJ3ASs1OiKjqhdP9a8N8CNgO77vH1frXlEAIOkICCLAFw1+BlgKzEJHB0VGUg9wDNgU4MfAD6LEEhoWUQAgmRIIVBncyYU7BtagewZErsUgsBN4CXgtQH2UWLuGRRQASEYK1JUFEtfhpwbuAG7AGwqJyPDs50Jx39oAb0SJtWhYZCQVaAhkFPQm37hOAmuBW5KBwK0KBEQ+dLV/GHg5+FG+c+ZX9bbg1f4iygBI1mQCDEIBMC4QlgGrgWn41sB1+H0DBXoOJU/F8Wt54/jtfGuBtwOsMzgWvLFPPEpMbXtFAYDkRFBQGYivBLsLqMUbDFUDZXoeJY8m/nbgIHAAOI4X99VHiTVqeEQBgORyEBCBMCEQFiSDgHvwosEx+NFCbU1Jrkngafwh4AzwAvBC8Mr+FvwMv6r6RQGA5E0gUOCBQGIB2HR8W+ABYAbeargcKNJzKlm+2u9OTvrbgGa8mn8HcEiX84gCAFEw4PUCEwJhFdgyCIvx5kJzlRGQLNYBvA48GWATfp7/hPb1RQGAyCWzApQGEuXJyf9hCA+CzcabC6loUDJVggtFfW+DrQU2QDgU4Kz684sCAJHhBwORQKIYKDNsfiB8Gt8imIrXC5TrGZa0P6Z+7LUT78v/JvAusNGwLsMGjXrt7YsCAJGrDAQMQnEgjAOqwOZAuJsLnQanAGPRZUSSGv144V4ncAbCjoC9ZbA5eMq/G+hRml8UAIiMfEBQnLyVcBy+TXAzsBC/g2A+MFGjJKPgUHLS3x5gC3DCYC+EloB1RIn1aIhEAYBIijIDgVAMoRzfElgB4V6wpcmswET8WGGJnnW5CnGgFWgDTgV4DjhlsD54m95BoE/H90QBgEh6g4EIhPJAYizYBPxI4fzkxzy8A+EEvG5A5HK68a58TcAJYB/QAOwM3rhnAOiMEotrqEQBgEgGBgOBRBGEQiMyNcBSSKwEW4D3GJgGVABVQBSvHTB9P+TRI3Lhoz+5wu8EuoADAbYbtgvY4el9hoB+TfqiAEAku4IBCwSDRASs0IhMDCRuhTANbB4wB68fqEp+qOdA7utKru6bga0Q9gWsAb+8ai/QZ1jCsGDUq5BPFACI5GBwUBBIrIRwK9hMfLugGpiN301QpVHKGQfw/vtHgUa8kO8o8I7O54sCAJH8DgZKgfEQqgOJ5WBTgKV4ceFs/ITBOI1UVmhNrvDP4NfrNiW78J0FtkeJdWuIRBQAiFwuICgBxgZCOSSmJAOC+WA1EK6DMD3ZnbBEo5VWfcB24BRwDj+qtx/C6YCdAxqjxFo1TCIKAESuNiCI4FsCYwOJGRAmejBgCyBMAJbjxw4n8t4LjEzfa9cw7P5x/usBvCL/OF60dwLYA+FIwFqBzgiRM0Cruu+JfDQVPIkMTwKvFm9PTjwGRAyqApRBmARWhtcOVILN8iCBxReCBSbipw6ivPdOg3z9PhxKTuzxi8Z4CD9f3wzsC9BsWCOEnXiXvY5kt71+wzogtCZ/f7joz1DhnogyACJpyRYUAeOT2wfVQAVYJV5LUAVMxo8inm9QNB2/6Ojij4LkZy76cTZO8D3JCTqRnNjPT/CDydV8b3KyH8LP4B9O/rgbwpmAdRvWYdgpo75fT5eIAgCRbAsKINmVMJCogjA2ObEb2Fi8QdFk/BTCBPyOgwnJ3z7xoq8LuHDngV30Y7soYLicQoZ3X0I8OSFfblJPvC8zcn6CT7zv950D9iQzJ/14c502vNFON4Q2YCBgXUDCsH7DmoFuo14PjYgCAJG8CBAKgTGBUBkIZUaoxLcTDKycCx0Ma5I/f35Cn5jMJBTjtyR+mGqG1wmxDa+kv9Tk35xctZ93/mKc/uRq/8xF/6rugJ1J/vohwzoN6wU6jPpBveoiCgBEZPhZhPLzK/1AiARCGcm6gmTQ8GHf7mV4geJH6YPQd4mfTwSs532r/LhhvYbF8bS+Vu8iIiIiIpnq/wPnTaZnpKrH0AAAAABJRU5ErkJggg=="

# Create a streaming image by streaming the base64 string to a bitmap streamsource
$bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
$bitmap.BeginInit()
$bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
$bitmap.EndInit()

# Freeze() prevents memory leaks.
$bitmap.Freeze()

# Set source here. Take note in the XAML as to where the variable name was taken.
$Window.Icon = $bitmap




$Window.add_Loaded({
        #$Icon                            = New-Object system.drawing.icon ("$($workingdir)\Files\Icon.ico")
        #$Window.Icon = $Icon 
    })

#Read Parameters
#$parameterfile = $workingdir + "\Files\" + "Customerparameters.ps1"
#. $parameterfile


#Read INI Parameters
#$INIRead = $workingdir + "\Files\" + "Config.ps1"
#. $INIRead


Function New-ProgressBar {
 
    [void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
    $syncHash = [hashtable]::Synchronized(@{})
    $newRunspace = [runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $newRunspace
    $syncHash.AdditionalInfo = ''
    $newRunspace.ApartmentState = "STA" 
    $newRunspace.ThreadOptions = "ReuseThread"           
    $data = $newRunspace.Open() | Out-Null
    $newRunspace.SessionStateProxy.SetVariable("syncHash", $syncHash)           
    $PowerShellCommand = [PowerShell]::Create().AddScript({    
            [string]$xaml = @" 
        <Window 
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" 
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml" 
            Name="Window" Title="Progress..." WindowStartupLocation = "CenterScreen" 
            Width = "560" Height="130" SizeToContent="Height" ShowInTaskbar = "True"> 
            <StackPanel Margin="20">
               <ProgressBar Width="560" Name="ProgressBar" IsIndeterminate="True" Opacity="0.9" BorderBrush="Transparent">
               
                    <ProgressBar.Foreground>
                        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                            <GradientStop Color="#00FFCB00"/>
                            <GradientStop Color="#00FFCB00" Offset="1"/>
                            <GradientStop Color="#7FFFCB00" Offset="0.49"/>
                        </LinearGradientBrush>
                    </ProgressBar.Foreground>
               </ProgressBar>
               <TextBlock Text="{Binding ElementName=ProgressBar, Path=Value, StringFormat={}{0:0}%}" HorizontalAlignment="Center" VerticalAlignment="Center"  Visibility="Hidden" />
               <TextBlock Name="AdditionalInfoTextBlock" Text="" HorizontalAlignment="Center" VerticalAlignment="Center" />
            </StackPanel> 
        </Window>
"@ 
   
            $syncHash.Window = [Windows.Markup.XamlReader]::parse( $xaml ) 
            #===========================================================================
            # Store Form Objects In PowerShell
            #===========================================================================
        ([xml]$xaml).SelectNodes("//*[@Name]") | % { $SyncHash."$($_.Name)" = $SyncHash.Window.FindName($_.Name) }

            $updateBlock = {            
            
                $SyncHash.Window.Title = $SyncHash.Activity
                $SyncHash.ProgressBar.Value = $SyncHash.PercentComplete
                $SyncHash.AdditionalInfoTextBlock.Text = $SyncHash.AdditionalInfo
                #$SyncHash.Window.MinWidth = $SyncHash.Window.ActualWidth
                       
            }

            ############### New Blog ##############
            $syncHash.Window.Add_SourceInitialized( {            
                    ## Before the window's even displayed ...            
                    ## We'll create a timer            
                    $timer = new-object System.Windows.Threading.DispatcherTimer            
                    ## Which will fire 4 times every second            
                    $timer.Interval = [TimeSpan]"0:0:0.01"            
                    ## And will invoke the $updateBlock            
                    $timer.Add_Tick( $updateBlock )            
                    ## Now start the timer running            
                    $timer.Start()            
                    if ( $timer.IsEnabled ) {            
                        Write-Host "Clock is running. Don't forget: RIGHT-CLICK to close it."            
                    }
                    else {            
                        $clock.Close()            
                        Write-Error "Timer didn't start"            
                    }            
                } )

            $syncHash.Window.ShowDialog() | Out-Null 
            $syncHash.Error = $Error 

        }) 
    $PowerShellCommand.Runspace = $newRunspace 
    $data = $PowerShellCommand.BeginInvoke() 
   
    
    Register-ObjectEvent -InputObject $SyncHash.Runspace `
        -EventName 'AvailabilityChanged' `
        -Action { 
                
        if ($Sender.RunspaceAvailability -eq "Available") {
            $Sender.Closeasync()
            $Sender.Dispose()
        } 
                
    } | Out-Null

    return $syncHash

}

function Show-Toast {
    param (
        [string]$Message = "Done!",
        [int]$Duration = 3000  # in milliseconds
    )

    $toast = New-Object System.Windows.Forms.NotifyIcon
    $toast.Icon = [System.Drawing.SystemIcons]::Information
    $toast.BalloonTipTitle = "Application Manager"
    $toast.BalloonTipText = $Message
    $toast.Visible = $true
    $toast.ShowBalloonTip($Duration)

    Start-Sleep -Milliseconds $Duration
    $toast.Dispose()
}

function Start-ProgressBar {

    <#Severity
    1 = Info
    2 = Warning
    3 = Error
    #>

 
    Param (
        [Parameter(Mandatory = $true)]
        $ProgressBar,
        [Parameter(Mandatory = $false)]
        [String]$Activity,
        [int]$PercentComplete,
        [String]$Status = $Null,
        [int]$SecondsRemaining = $Null,
        [int]$Severity = '1',
        [String]$CurrentOperation = $Null
    ) 
   
    Write-Verbose -Message "Setting activity to $Activity"
    $ProgressBar.Activity = $Activity

    if ($PercentComplete) {
       
        Write-Verbose -Message "Setting PercentComplete to $PercentComplete"
        $ProgressBar.PercentComplete = $PercentComplete

    }
   
    if ($SecondsRemaining) {

        [String]$SecondsRemaining = "$SecondsRemaining Seconds Remaining"

    }
    else {

        [String]$SecondsRemaining = $Null

    }

    Write-Verbose -Message "Setting AdditionalInfo to $Status       $SecondsRemaining$(if($SecondsRemaining){ " seconds remaining..." }else {''})       $CurrentOperation"
    $ProgressBar.AdditionalInfo = "$CurrentOperation"
    #$ProgressBar.AdditionalInfo.Foreground = "Red"
    #$ProgressBar.AdditionalInfo = "$Status       $SecondsRemaining       $CurrentOperation"


}



function Stop-ProgressBar {
   
   
    $labelProgressbar.Visibility = "Hidden"
    $RectangleProgressBar.Visibility = "Hidden"
    $ProgessBarGlobal.Visibility = "Hidden"
    #$ProgressBar.Visibility = "Visible"
    $LabelProgressBar.Content = ""


}




#############################################################################
#Begin Config.ps1
#############################################################################
$ConfigFile = $workingdir + "\Files\" + "config.ini"
$registryPath = "HKCU:\SOFTWARE\Kapsch\ApplicationManager"

<#
  Windows-only functions for reading from / updating INI files, via the Windows API.
  
  You must dot-source this script for the functions to become available in a session, in the course
  of which helper type [WinApiHelper.IniFile] is on-demand-compiled and also added to the session; e.g.:
    
    . .\IniFileHelper.ps1
    
  * Get-IniValue returns a specific value for a given section and key from an INI file, as a string.
    It optionally enumerates all section names or all key names inside a given section.
  
  * Set-IniValue updates a specific value for a given section and key from an INI file.
    It optionally deletes an entry or even an entire section of entries.

  Invoke the functions with Get-Help -Examples to learn more.

  Requires version 3 or higher.

  These functions were inspired by https://stackoverflow.com/a/55437750/45375
  
#>


# Check prerequisites:
# PS version:
# IMPORTANT: Be sure to make any modifications PSv3-compatible, which notably requires avoiding use of ::new()
#            v3 is required at a minimum due to use of [NullString]::Value
#requires -Version 3
# Windows-only.
#if ($env:OS -ne 'Windows_NT') { Throw "These functions are only supported on Windows." }
# Make sure this script is being dot-sourced.
#if (-not ($MyInvocation.InvocationName -eq '.' -or $MyInvocation.Line -eq '')) { Throw "Please invoke this script dot-sourced." }


# Create helper type [WinApiHelper.IniFile], which wraps P/Invoke calls to the Get/WritePrivateProfileString()
# Windows API functions.
Add-Type -Namespace WinApiHelper -Name IniFile -MemberDefinition @'
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
  // Note the need to use `[Out] byte[]` instead of `System.Text.StringBuilder` in order to support strings with embedded NUL chars.
  public static extern uint GetPrivateProfileString(string lpAppName, string lpKeyName, string lpDefault, [Out] byte[] lpBuffer, uint nSize, string lpFileName);
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
  public static extern bool WritePrivateProfileString(string lpAppName, string lpKeyName, string lpString, string lpFileName);
'@




Function Set-IniValue {
    <#
.SYNOPSIS
Updates a given entry's value in an INI file.
Optionally *deletes* from the file:
* an entry (if -Value is omitted)
* a entire section (if -Key is omitted)
.EXAMPLE
Set-IniValue file.ini section1 key1 value1
Updates the value of the entry whose key is key1 in section section1 in file
file.ini.
.EXAMPLE
Set-IniValue file.ini section1 key1
Deletes the entry whose key is key1 from section section1 in file file.ini.
.EXAMPLE
Set-IniValue file.ini section1
Deletes the entire section section1 from file file.ini.
#>
    param(
        [Parameter(Mandatory)] [string] $LiteralPath,
        [Parameter(Mandatory)] [string] $Section,
        [string] $Key,
        [string] $Value
    )
    # Make sure that bona fide `null` is passed for omitted parameters, as only true `null`
    # values are recognized as requests to *delete* entries.
    if (-not $PSBoundParameters.ContainsKey('Key')) { $Key = [NullString]::Value }
    if (-not $PSBoundParameters.ContainsKey('Value')) { $Value = [NullString]::Value }

    # Convert the path to an *absolute* one, since .NET's and the WinAPI's 
    # current dir. is usually differs from PowerShell's.
    $fullPath = Convert-Path -ErrorAction Stop -LiteralPath $LiteralPath
    $ok = [WinApiHelper.IniFile]::WritePrivateProfileString($Section, $Key, $Value, $fullPath)
    if (-not $ok) { Throw "Updating INI file failed: $fullPath" }

}


#Read Config.ini File
function Get-IniFile {  
    param(  
        [parameter(Mandatory = $true)] [string] $filePath  
    )  
    
    $anonymous = "NoSection"
  
    $ini = @{}  
    switch -regex -file $filePath {  
        "^\[(.+)\]$" {
            # Section    
            $section = $matches[1]  
            $ini[$section] = @{}  
            $CommentCount = 0  
        }  

        "^(;.*)$" {
            # Comment    
            if (!($section)) {  
                $section = $anonymous  
                $ini[$section] = @{}  
            }  
            $value = $matches[1]  
            
            $CommentCount = $CommentCount + 1  
            $name = "Comment" + $CommentCount  
            $ini[$section][$name] = $value  
        }   

        "(.+?)\s*=\s*(.*)" {
            # Key    
            if (!($section)) {  
                $section = $anonymous  
                $ini[$section] = @{}  
            }  
            
            $name, $value = $matches[1..2] 
            $ini[$section][$name] = $value  
        }  
    }  

    return $ini  
}  



#############################################################################
# BEGIN Init Variables
#############################################################################
function Init-Variables {
    $iniFile = Get-IniFile $ConfigFile
    # Global
    $global:logfile = $iniFile.Global.logfile
    $global:Packagefolderpath = $iniFile.Global.Packagefolderpath
    $global:Testpackagestring = $iniFile.Global.Testpackagestring
    $global:Packagedelimiter = $iniFile.Global.Packagedelimiter
    $global:Retiredpackagefolderpath = $iniFile.Global.Retiredpackagefolderpath
    $global:Showonlynewpackages = $iniFile.Global.Showonlynewpackages
    $global:MaxLogSizeInKB = $iniFile.Global.MaxLogSizeInKB
    $global:mailserver = $iniFile.Global.mailserver
    $global:mailfrom = $iniFile.Global.mailfrom
    $global:mailrecipients = $iniFile.Global.mailrecipients
    $global:showsummaryfile = $iniFile.Global.showsummaryfile
    $global:PSADTTemplate = $iniFile.Global.PSADTTemplate

    #ConfigMgr
    $global:SiteCode = $iniFile.ConfigMgr.SiteCode
    $global:DPGroup = $iniFile.ConfigMgr.DPGroup
    $global:ApplicationFolderName = $iniFile.ConfigMgr.ApplicationFolderName
    $global:ApplicationtestCollectionname = $iniFile.ConfigMgr.ApplicationtestCollectionname
    $global:CollectionFolderName = $iniFile.ConfigMgr.CollectionFolderName
    $global:Standardapplicationfoldername = $iniFile.ConfigMgr.Standardapplicationfoldername
    $global:CollectionUninstallFolderName = $iniFile.ConfigMgr.CollectionUninstallFolderName
    $global:DeviceLimitingCollection = $iniFile.ConfigMgr.DeviceLimitingCollection
    $global:UserLimitingCollection = $iniFile.ConfigMgr.UserLimitingCollection
    # Assuming $iniFile.ConfigMgr.RunInstallAs32Bit returns a string like "True" or "False"
    $global:RunInstallAs32Bit = [bool]::Parse($iniFile.ConfigMgr.RunInstallAs32Bit)
    $global:DownloadOnSlowNetwork = [bool]::Parse($iniFile.ConfigMgr.DownloadOnSlowNetwork)
    $global:AllowFallbackSourceLocation = [bool]::Parse($iniFile.ConfigMgr.AllowFallbackSourceLocation)
    $global:AllowInteractionDefault = [bool]::Parse($iniFile.ConfigMgr.AllowInteractionDefault)
   
    #AD
    $global:CreateADGroup = $iniFile.AD.CreateADGroup
    $global:DomainNetbiosName = $iniFile.AD.DomainNetbiosName
    $global:DeviceOUPath = $iniFile.AD.DeviceOUPath
    $global:DeviceUninstallOUPath = $iniFile.AD.DeviceUninstallOUPath
    $global:DeviceOURetiredPath = $iniFile.AD.DeviceOURetiredPath
    $global:DeviceOUProdPath = $iniFile.AD.DeviceOUProdPath
    $global:UserOUPath = $iniFile.AD.UserOUPath
    $global:ADGroupNamePrefix = $iniFile.AD.ADGroupNamePrefix
    $global:ADUninstallGroupNamePrefix = $iniFile.AD.ADUninstallGroupNamePrefix

    #Azure
    $global:TenantName = $iniFile.Azure.TenantName
    #$global:AADUser= $iniFile.Azure.AADUser
    $global:IntuneOutputFolder = $iniFile.Azure.IntuneOutputFolder
    $global:CleanUpIntuneOutputFolder = $iniFile.Azure.CleanUpIntuneOutputFolder
    $global:PilotAADGroup = $iniFile.Azure.PilotAADGroup
    $global:AADGroupNamePrefix = $iniFile.Azure.AADGroupNamePrefix
    $global:AzApplicationName = $iniFile.Azure.AzApplicationName
    $AzApplicationNameFromIni = $iniFile.Azure.AzApplicationName

    # Check if the Azure Application Name from the INI file is empty or null
    if ($null -eq $AzApplicationNameFromIni -or $AzApplicationNameFromIni -eq "") {
        # Set the global variable to the value from the TextBox
        $global:AzApplicationName = $TextBoxAzApplicationName.Text
    }
    else {
        # Set the global variable to the value from the INI file
        $global:AzApplicationName = $AzApplicationNameFromIni
    }
    #$global:ClientSecret = $iniFile.Azure.ClientSecret
    $global:AppId = $iniFile.Azure.AppId
    #### Get AADUser from User Registry ####
    $val = Get-ItemProperty -Path $registryPath -Name "AADUser" -ErrorAction SilentlyContinue
    $global:AADUser = $val.AADUser


    #SWMapping
    $global:SWMappingenabled = $iniFile.SWMapping.SWMappingenabled
    $global:SQLServer = $iniFile.SWMapping.SQLServer
    $global:SWMapDBName = $iniFile.SWMapping.SWMapDBName
    $global:SCCMDBName = $iniFile.SWMapping.SCCMDBName
    $global:SWMappingTable = $iniFile.SWMapping.SWMappingTable
    $global:SWMappingTableFiles = $iniFile.SWMapping.SWMappingTableFiles
    $global:SWProductTable = $iniFile.SWMapping.SWProductTable
    $global:SWProductTable = $iniFile.SWMapping.SWProductTable
}


#############################################################################
# END END Init Variables
#############################################################################

#############################################################################
# BEGIN Get Values
#############################################################################
function Load-Config {
    # Global
    $TextBoxlogfile.Text = $logfile
    $TextBoxPackagefolderpath.Text = $iniFile.Global.Packagefolderpath
    $TextBoxTestpackagestring.Text = $iniFile.Global.Testpackagestring
    $TextBoxPackagedelimiter.Text = $iniFile.Global.Packagedelimiter
    $TextBoxRetiredpackagefolderpath.Text = $iniFile.Global.Retiredpackagefolderpath
    if ($iniFile.Global.Showonlynewpackages -eq $true) {
        $comboBoxShowonlynewpackages.SelectedIndex = 0
    }
    else {
        $comboBoxShowonlynewpackages.SelectedIndex = 1
    }
    $TextBoxMaxLogSizeInKB.Text = $iniFile.Global.MaxLogSizeInKB
    $TextBoxmailserver.Text = $iniFile.Global.mailserver
    $TextBoxmailfrom.Text = $iniFile.Global.mailfrom
    $TextBoxmailrecipients.Text = $iniFile.Global.mailrecipients
    if ($iniFile.Global.showsummaryfile -eq $true) {
        $comboBoxshowsummaryfile.SelectedIndex = 0
    }
    else {
        $comboBoxshowsummaryfile.SelectedIndex = 1
    }
    $TextBoxPSADTTemplate.Text = $iniFile.Global.PSADTTemplate


    # ConfigMgr
    $TextBoxSiteCode.Text = $iniFile.ConfigMgr.SiteCode
    $TextBoxDPGroup.Text = $iniFile.ConfigMgr.DPGroup
    $TextBoxApplicationFolderName.Text = $iniFile.ConfigMgr.ApplicationFolderName
    $TextBoxApplicationtestCollectionname.Text = $iniFile.ConfigMgr.ApplicationtestCollectionname
    $TextBoxCollectionFolderName.Text = $iniFile.ConfigMgr.CollectionFolderName
    #$TextBoxStandardapplicationfoldername.Text = $iniFile.ConfigMgr.Standardapplicationfoldername
    $TextBoxCollectionUninstallFolderName.Text = $iniFile.ConfigMgr.CollectionUninstallFolderName
    $TextBoxDeviceLimitingCollection.Text = $iniFile.ConfigMgr.DeviceLimitingCollection
    $TextBoxUserLimitingCollection.Text = $iniFile.ConfigMgr.UserLimitingCollection
    
    if ($iniFile.ConfigMgr.RunInstallAs32Bit -eq $true) {
        $comboBoxRunInstallAs32Bit.SelectedIndex = 0
    }
    else {
        $comboBoxRunInstallAs32Bit.SelectedIndex = 1
    }

    if ($iniFile.ConfigMgr.DownloadOnSlowNetwork -eq $true) {
        $comboBoxDownloadOnSlowNetwork.SelectedIndex = 0
    }
    else {
        $comboBoxDownloadOnSlowNetwork.SelectedIndex = 1
    }

    if ($iniFile.ConfigMgr.AllowFallbackSourceLocation -eq $true) {
        $comboBoxAllowFallbackSourceLocation.SelectedIndex = 0
    }
    else {
        $comboBoxAllowFallbackSourceLocation.SelectedIndex = 1
    }

    if ($iniFile.ConfigMgr.AllowInteractionDefault -eq $true) {
        $comboBoxAllowInteractionDefault.SelectedIndex = 0
    }
    else {
        $comboBoxAllowInteractionDefault.SelectedIndex = 1
    }


    #ActiveDirectory
    if ($iniFile.AD.CreateADGroup -eq $true) {
        $comboBoxCreateADGroup.SelectedIndex = 0
    }
    else {
        $comboBoxCreateADGroup.SelectedIndex = 1
    }

    $TextBoxDomainNetbiosName.Text = $iniFile.AD.DomainNetbiosName
    $TextBoxDeviceOUPath.Text = $iniFile.AD.DeviceOUPath
    $TextBoxDeviceUninstallOUPath.Text = $iniFile.AD.DeviceUninstallOUPath
    $TextBoxDeviceOURetiredPath.Text = $iniFile.AD.DeviceOURetiredPath
    $TextBoxUserOUPath.Text = $iniFile.AD.UserOUPath
    $TextBoxADGroupNamePrefix.Text = $iniFile.AD.ADGroupNamePrefix
    $TextBoxADUninstallGroupNamePrefix.Text = $iniFile.AD.ADUninstallGroupNamePrefix

    #Azure
    $TextBoxTenantName.Text = $iniFile.Azure.TenantName
    #$TextBoxAADUser.Text = $iniFile.Azure.AADUser
    $TextBoxIntuneOutputFolder.Text = $iniFile.Azure.IntuneOutputFolder
    $TextBoxPilotAADGroup.Text = $iniFile.Azure.PilotAADGroup
    $TextBoxAADGroupNamePrefix.Text = $iniFile.Azure.AADGroupNamePrefix
    if ($iniFile.Azure.CleanUpIntuneOutputFolder -eq $true) {
        $comboBoxCleanUpIntuneOutputFolder.SelectedIndex = 0
    }
    else {
        $comboBoxCleanUpIntuneOutputFolder.SelectedIndex = 1
    }
    # Check if $iniFile.Azure.AzApplicationName is not empty or null
    if (-not [string]::IsNullOrEmpty($iniFile.Azure.AzApplicationName)) {
        # Assign the value to the TextBox only if it is not empty or null
        $TextBoxAzApplicationName.Text = $iniFile.Azure.AzApplicationName
    }

    #$TextBoxClientSecret.Password = $iniFile.Azure.ClientSecret

    #### Get AADUser from User Registry ####
    $val = Get-ItemProperty -Path $registryPath -Name "AADUser" -ErrorAction SilentlyContinue
    $TextBoxAADUser.Text = $val.AADUser
    
}


#############################################################################
# END Get Values
#############################################################################


#############################################################################
# BEGIN Set Values
#############################################################################

function Save-Config {
    #Set Config
    #Global
    Set-IniValue $ConfigFile 'Global' 'logfile' $TextBoxlogfile.Text
    Set-IniValue $ConfigFile 'Global' 'Packagefolderpath' $TextBoxPackagefolderpath.Text
    Set-IniValue $ConfigFile 'Global' 'Testpackagestring' $TextBoxTestpackagestring.Text
    Set-IniValue $ConfigFile 'Global' 'Packagedelimiter' $TextBoxPackagedelimiter.Text
    Set-IniValue $ConfigFile 'Global' 'Retiredpackagefolderpath' $TextBoxRetiredpackagefolderpath.Text
    Set-IniValue $ConfigFile 'Global' 'Showonlynewpackages' $comboBoxShowonlynewpackages.SelectedItem.Content
    Set-IniValue $ConfigFile 'Global' 'MaxLogSizeInKB' $TextBoxMaxLogSizeInKB.Text
    Set-IniValue $ConfigFile 'Global' 'mailserver' $TextBoxmailserver.Text
    Set-IniValue $ConfigFile 'Global' 'mailfrom' $TextBoxmailfrom.Text
    Set-IniValue $ConfigFile 'Global' 'mailrecipients' $TextBoxmailrecipients.Text
    Set-IniValue $ConfigFile 'Global' 'showsummaryfile' $comboBoxShowsummaryfile.SelectedItem.Content
    Set-IniValue $ConfigFile 'Global' 'PSADTTemplate' $TextBoxPSADTTemplate.Text

    #ConfigMgr
    Set-IniValue $ConfigFile 'ConfigMgr' 'SiteCode' $TextBoxSiteCode.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'DPGroup' $TextBoxDPGroup.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'ApplicationFolderName' $TextBoxApplicationFolderName.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'ApplicationtestCollectionname' $TextBoxApplicationtestCollectionname.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'CollectionFolderName' $TextBoxCollectionFolderName.Text
    #Set-IniValue $ConfigFile 'ConfigMgr' 'Standardapplicationfoldername' $TextBoxStandardapplicationfoldername.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'CollectionUninstallFolderName' $TextBoxCollectionUninstallFolderName.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'DeviceLimitingCollection' $TextBoxDeviceLimitingCollection.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'UserLimitingCollection' $TextBoxUserLimitingCollection.Text
    Set-IniValue $ConfigFile 'ConfigMgr' 'RunInstallAs32Bit' $comboBoxRunInstallAs32Bit.SelectedItem.Content
    Set-IniValue $ConfigFile 'ConfigMgr' 'DownloadOnSlowNetwork' $comboBoxDownloadOnSlowNetwork.SelectedItem.Content
    Set-IniValue $ConfigFile 'ConfigMgr' 'AllowFallbackSourceLocation' $comboBoxAllowFallbackSourceLocation.SelectedItem.Content
    Set-IniValue $ConfigFile 'ConfigMgr' 'AllowInteractionDefault' $comboBoxAllowInteractionDefault.SelectedItem.Content
    
    #AD
    Set-IniValue $ConfigFile 'AD' 'CreateADGroup' $comboBoxCreateADGroup.SelectedItem.Content
    Set-IniValue $ConfigFile 'AD' 'DomainNetbiosName' $TextBoxDomainNetbiosName.Text
    Set-IniValue $ConfigFile 'AD' 'DeviceOUPath' $TextBoxDeviceOUPath.Text
    Set-IniValue $ConfigFile 'AD' 'DeviceUninstallOUPath' $TextBoxDeviceUninstallOUPath.Text
    Set-IniValue $ConfigFile 'AD' 'DeviceOURetiredPath' $TextBoxDeviceOURetiredPath.Text
    #Set-IniValue $ConfigFile 'AD' 'DeviceOUProdPath' $TextBoxDeviceOUProdPath.Text
    Set-IniValue $ConfigFile 'AD' 'UserOUPath' $TextBoxUserOUPath.Text
    Set-IniValue $ConfigFile 'AD' 'ADGroupNamePrefix' $TextBoxADGroupNamePrefix.Text
    Set-IniValue $ConfigFile 'AD' 'ADUninstallGroupNamePrefix' $TextBoxADUninstallGroupNamePrefix.Text

    #Azure
    Set-IniValue $ConfigFile 'Azure' 'TenantName' $TextBoxTenantName.Text
    #Set-IniValue $ConfigFile 'Azure' 'AADUser' $TextBoxAADUser.Text
    Set-IniValue $ConfigFile 'Azure' 'IntuneOutputFolder' $TextBoxIntuneOutputFolder.Text
    Set-IniValue $ConfigFile 'Azure' 'PilotAADGroup' $TextBoxPilotAADGroup.Text
    Set-IniValue $ConfigFile 'Azure' 'CleanUpIntuneOutputFolder' $comboBoxCleanUpIntuneOutputFolder.SelectedItem.Content
    Set-IniValue $ConfigFile 'Azure' 'AADGroupNamePrefix' $TextBoxAADGroupNamePrefix.Text
    Set-IniValue $ConfigFile 'Azure' 'AzApplicationName' $TextBoxAzApplicationName.Text

    #### Save AADUser in User Registry ####
    If (!(Test-Path $registryPath)) {
        #Create Registry Hive if not exist
        New-Item -Path $registryPath -Force | Out-Null
        New-ItemProperty -Path $registryPath -Name AADUser -Value $TextBoxAADUser.Text -PropertyType String -Force | Out-Null
    }
    ELSE {
        New-ItemProperty -Path $registryPath -Name AADUser -Value $TextBoxAADUser.Text -PropertyType String -Force | Out-Null
    }

    #For ReInitialize Variables
    Init-Variables
    Write-Host "Config successfully saved!"
}

#############################################################################
# END Set Values
#############################################################################

function OpenFolderDialog {
    param(
        $InitialDirectory
    )
    $AssemblyFullName = 'System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'
    $Assembly = [System.Reflection.Assembly]::Load($AssemblyFullName)
    $OpenFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $OpenFileDialog.AddExtension = $false
    $OpenFileDialog.CheckFileExists = $false
    $OpenFileDialog.DereferenceLinks = $true
    $OpenFileDialog.Filter = "Folders|`n"
    $OpenFileDialog.Multiselect = $false
    $OpenFileDialog.Title = "Select folder"
    $OpenFileDialog.InitialDirectory = $InitialDirectory
    $OpenFileDialogType = $OpenFileDialog.GetType()
    $FileDialogInterfaceType = $Assembly.GetType('System.Windows.Forms.FileDialogNative+IFileDialog')
    $IFileDialog = $OpenFileDialogType.GetMethod('CreateVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $null)
    $null = $OpenFileDialogType.GetMethod('OnBeforeVistaDialog', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $IFileDialog)
    [uint32]$PickFoldersOption = $Assembly.GetType('System.Windows.Forms.FileDialogNative+FOS').GetField('FOS_PICKFOLDERS').GetValue($null)
    $FolderOptions = $OpenFileDialogType.GetMethod('get_Options', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($OpenFileDialog, $null) -bor $PickFoldersOption
    $null = $FileDialogInterfaceType.GetMethod('SetOptions', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $FolderOptions)
    $VistaDialogEvent = [System.Activator]::CreateInstance($AssemblyFullName, 'System.Windows.Forms.FileDialog+VistaDialogEvents', $false, 0, $null, $OpenFileDialog, $null, $null).Unwrap()
    [uint32]$AdviceCookie = 0
    $AdvisoryParameters = @($VistaDialogEvent, $AdviceCookie)
    $AdviseResult = $FileDialogInterfaceType.GetMethod('Advise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdvisoryParameters)
    $AdviceCookie = $AdvisoryParameters[1]
    $Result = $FileDialogInterfaceType.GetMethod('Show', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, [System.IntPtr]::Zero)
    $null = $FileDialogInterfaceType.GetMethod('Unadvise', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $AdviceCookie)
    if ($Result -eq [System.Windows.Forms.DialogResult]::OK) {
        $FileDialogInterfaceType.GetMethod('GetResult', @('NonPublic', 'Public', 'Static', 'Instance')).Invoke($IFileDialog, $null)
    }
    Write-Output $OpenFileDialog.FileName

}

#Read Config
$iniFile = Get-IniFile $ConfigFile


$logfile = $iniFile.Global.Logfile -replace "%USERNAME%", $env:UserName

write-host "Current Logpath: $logfile"



Init-Variables
Load-Config

$buttonPackagefolderpath.add_Click({
        $Packagefolderpath = OpenFolderDialog $Packagefolderpath
        $TextBoxPackagefolderpath.Text = $Packagefolderpath
    })

$buttonRetiredpackagefolderpath.add_Click({
        if (!$TextBoxRetiredpackagefolderpath.Text) {
            $InitialDirectory = $Packagefolderpath
        }
        else {
            $InitialDirectory = $TextBoxRetiredpackagefolderpath.Text
        }
        $Retiredpackagefolderpath = OpenFolderDialog $InitialDirectory
        $TextBoxRetiredpackagefolderpath.Text = $Retiredpackagefolderpath
    })

$buttonPSADTTemplate.add_Click({
        if (!$TextBoxPSADTTemplate.Text) {
            $InitialDirectory = $Packagefolderpath
        }
        else {
            $InitialDirectory = $TextBoxPSADTTemplate.Text
        }
        $PSADTTemplate = OpenFolderDialog $InitialDirectory
        $TextBoxRPSADTTemplate.Text = $PSADTTemplate
    })

$buttonIntuneOutputFolder.add_Click({
        if (!$TextBoxIntuneOutputFolder.Text) {
            $InitialDirectory = $Packagefolderpath
        }
        else {
            $InitialDirectory = $TextBoxIntuneOutputFolder.Text
        }
        $IntuneOutputFolder = OpenFolderDialog $InitialDirectory
        $TextBoxIntuneOutputFolder.Text = $IntuneOutputFolder
    })


$buttonInstallIntuneModule.add_Click({ Install-IntuneModule })
$buttonInstallWingetModule.add_Click({ Install-WingetClient })

$buttonConfigSave.add_Click({
        $ProgressBar = New-ProgressBar
        Start-ProgressBar -ProgressBar $ProgressBar -Activity "Save Config" -CurrentOperation "Saving Config"
        Save-Config
        Close-ProgressBar $ProgressBar
    })


#############################################################################
#END Config.ps1
#############################################################################





$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path $logfile -append
#$ErrorActionPreference = "Stop"

#############################################################################
#BEGIN Test-Administrator
#############################################################################
function Test-Administrator {  
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)  
}


function Get-WingetCmd {

    #WinGet Path (if User/Admin context)
    $UserWingetPath = Get-Command winget.exe -ErrorAction SilentlyContinue
    #WinGet Path (if system context)
    $SystemWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"

    #Get Winget Location in User/Admin context
    if ($UserWingetPath) {
        $Script:Winget = $UserWingetPath.Source
    }
    #Get Winget Location in System context
    elseif ($SystemWingetPath) {
        #If multiple version, pick last one
        $Script:Winget = $SystemWingetPath[-1].Path
    }
    else {
        Write-Host "WinGet is not installed, mandatory to run WinGet integration"
        
    }

}



function Get-StoreApp-Icon ($AppID) {
    $icon = @{
        "@odata.type" = "#microsoft.graph.mimeContent"
        type          = "image/png"
        value         = [Convert]::ToBase64String(((Invoke-WebRequest -Uri ((Invoke-RestMethod -Method "Get" -Uri ("https://apps.microsoft.com/store/api/ProductsDetails/GetProductDetailsById/" + $AppID + "?hl=en-US&gl=US")).IconUrl)).Content))
    }
}



function Get-WingetAppInfo-Native ($SearchApp) {
    class Software {
        [string]$Name
        [string]$Id
    }

    #Search for winget apps
    $AppResult = & $Winget search $SearchApp --accept-source-agreements

    #Start Convertion of winget format to an array. Check if "-----" exists
    if (!($AppResult -match "-----")) {
        Write-Host "No application found."
        return
    }

    #Split winget output to lines
    $lines = $AppResult.Split([Environment]::NewLine) | Where-Object { $_ }

    # Find the line that starts with "------"
    $fl = 0
    while (-not $lines[$fl].StartsWith("-----")) {
        $fl++
    }

    $fl = $fl - 1

    #Get header titles
    $index = $lines[$fl] -split '\s+'

    # Line $fl has the header, we can find char where we find ID and Version
    $idStart = $lines[$fl].IndexOf($index[1])
    $versionStart = $lines[$fl].IndexOf($index[2])
    $TagsStart = $lines[$fl].IndexOf($index[3])
    $SourceStart = $lines[$fl].IndexOf($index[4])


    # Now cycle in real package and split accordingly
    $searchList = @()
    For ($i = $fl + 2; $i -le $lines.Length; $i++) {
        $line = $lines[$i]
        if ($line.Length -gt ($sourceStart + 5)) {
            $software = [Software]::new()
            $software.Name = $line.Substring(0, $idStart).TrimEnd()
            $software.Id = $line.Substring($idStart, $versionStart - $idStart).TrimEnd()
            
           
            #add formated soft to list
            $searchList += $software
        }
    }
    return $searchList
}


function Get-WingetAppInfo {
    param (
        [Parameter(Mandatory = $true)]
        [string]$SearchApp
    )

    class Software {
        [string]$Name
        [string]$Id
        [string]$Version
        [string]$Source
    }

    $useModule = $false

    if (Get-Command -Name "Find-WinGetPackage" -ErrorAction SilentlyContinue) {
        Write-Host "Winget module is available. Trying to search for app: $SearchApp"

        try {
            $AppResult = Find-WinGetPackage -Name $SearchApp -ErrorAction Stop

            $results = foreach ($app in $AppResult) {
                $software = [Software]::new()
                $software.Name = $app.Name
                $software.Id = $app.Id
                $software.Version = $app.Version
                $software.Source = $app.Source
                $software
            }

            return $results
        }
        catch {
            Write-Warning "Winget module failed with: $($_.Exception.Message). Falling back to native 'winget search'"
            # Continue to fallback below
        }
    }

    # If we get here, either module is missing or it failed
    return Get-WingetAppInfo-Native $SearchApp
}


function Get-WingetPackageDetails {
    param (
        [Parameter(Mandatory = $true)]
        [string]$PackageId
    )

    $output = winget show --id $PackageId --accept-source-agreements 2>&1

    if (-not $output) {
        Write-Warning "No output received from winget show."
        return $null
    }

    # Create hash table to store values
    $info = @{
        Id              = $PackageId
        Name            = $null
        Version         = $null
        Publisher       = $null
        PublisherUrl    = $null
        Description     = $null
        Homepage        = $null
        License         = $null
        LicenseUrl      = $null
        InstallerType   = $null
        InstallerUrl    = $null
        InstallerSha256 = $null
        ReleaseDate     = $null
    }

    foreach ($line in $output) {
        if ($line -match '^Name:\s+(.+)$') { $info.Name = $matches[1].Trim() }
        elseif ($line -match '^Version:\s+(.+)$') { $info.Version = $matches[1].Trim() }
        elseif ($line -match '^Herausgeber:\s+(.+)$') { $info.Publisher = $matches[1].Trim() }
        elseif ($line -match '^Herausgeber-URL:\s+(.+)$') { $info.PublisherUrl = $matches[1].Trim() }
        elseif ($line -match '^Startseite:\s+(.+)$') { $info.Homepage = $matches[1].Trim() }
        elseif ($line -match '^Lizenz:\s+(.+)$') { $info.License = $matches[1].Trim() }
        elseif ($line -match '^Lizenz-URL:\s+(.+)$') { $info.LicenseUrl = $matches[1].Trim() }
        elseif ($line -match '^Installertyp:\s+(.+)$') { $info.InstallerType = $matches[1].Trim() }
        elseif ($line -match '^Installer-URL:\s+(.+)$') { $info.InstallerUrl = $matches[1].Trim() }
        elseif ($line -match '^Sha256-Installer:\s+(.+)$') { $info.InstallerSha256 = $matches[1].Trim() }
        elseif ($line -match '^Freigabedatum:\s+(.+)$') { $info.ReleaseDate = $matches[1].Trim() }
        elseif ($line -match '^Beschreibung:\s+(.+)$') { $info.Description = $matches[1].Trim() }
    }

    return [PSCustomObject]$info
}

function Get-WingetInstallContent {
    param (
        [Parameter(Mandatory = $true)][string]$PackageId,
        [Parameter(Mandatory = $true)][string]$PackageVersion
    )

    try {
        $firstLetter = $PackageId.Substring(0,1).ToLower()
        $repoPath = "$($PackageId.Replace('.', '/'))"
        $url = "https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/$firstLetter/$repoPath/$PackageVersion/$PackageId.installer.yaml"

        $tmpPath = Join-Path $env:TEMP "$PackageId.installer.yaml"

        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Downloading YAML for $PackageId..."
        Invoke-WebRequest -Uri $url -OutFile $tmpPath -UseBasicParsing -ErrorAction Stop

        Import-Module powershell-yaml -ErrorAction Stop

        return Get-Content $tmpPath -Raw | ConvertFrom-Yaml
    }
    catch {
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "❌ Failed to fetch or parse YAML: $($_.Exception.Message)" -Severity 3
        return $null
    }
}

function Get-WingetLocaleContent {
    param (
        [Parameter(Mandatory = $true)][string]$PackageId,
        [Parameter(Mandatory = $true)][string]$PackageVersion,
        [string]$Locale = "en-US"
    )

    try {
        $firstLetter = $PackageId.Substring(0,1).ToLower()
        $repoPath = "$($PackageId.Replace('.', '/'))"
        $url = "https://raw.githubusercontent.com/microsoft/winget-pkgs/master/manifests/$firstLetter/$repoPath/$PackageVersion/$PackageId.locale.$Locale.yaml"

        $tmpPath = Join-Path $env:TEMP "$PackageId.locale.$Locale.yaml"

        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Downloading locale YAML for $PackageId..."
        Invoke-WebRequest -Uri $url -OutFile $tmpPath -UseBasicParsing -ErrorAction Stop

        Import-Module powershell-yaml -ErrorAction Stop

        return Get-Content $tmpPath -Raw | ConvertFrom-Yaml
    }
    catch {
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "⚠️ Failed to fetch or parse locale YAML: $($_.Exception.Message)" -Severity 2
        return $null
    }
}


function Get-WingetYamlValue {
    param (
        [Parameter(Mandatory = $true)][PSCustomObject]$YamlContent,
        [Parameter(Mandatory = $true)][string]$PropertyPath
    )

    try {
        $value = $YamlContent | Select-Object -ExpandProperty $PropertyPath -ErrorAction Stop
        return $value
    }
    catch {
        return $null
    }
}


function Download-PSADTTemplate {
    param (
        [string]$DownloadUrl = "https://github.com/PSAppDeployToolkit/PSAppDeployToolkit/releases/latest/download/PSAppDeployToolkit_Template_v4.zip",
        [Parameter(Mandatory = $true)]
        [string]$DestinationFolder,
        [Parameter(Mandatory = $true)]
        $ProgressBar
    )

    $templateSubFolder = Join-Path $DestinationFolder "_PSADT_Template_v4"

    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Preparing template download folder..."
    if (-not (Test-Path $templateSubFolder)) {
        New-Item -ItemType Directory -Path $templateSubFolder | Out-Null
    }

    $destinationZip = Join-Path $templateSubFolder "PSAppDeployToolkit_Template_v4.zip"

    try {
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Downloading PSADT v4 Template from GitHub..."
        Invoke-WebRequest -Uri $DownloadUrl -OutFile $destinationZip -UseBasicParsing

        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Extracting PSADT Template..."
        Expand-Archive -Path $destinationZip -DestinationPath $templateSubFolder -Force
        Remove-Item $destinationZip -Force

        # Return the template folder path
        return Get-Item $templateSubFolder
    }
    catch {
        [System.Windows.MessageBox]::Show("Download failed: $($_.Exception.Message)", "Error", "OK", "Error")
        return $null
    }
}



#############################################################################
#BEGIN Functions_Intune
#############################################################################
#Read Intune Functions
#$IntuneFunctions = $workingdir + "\Files\" + "Functions_Intune.ps1"
#. $IntuneFunctions

<#
  Script for Application Manager to Manage Intune / AAD Objects
  
  You must dot-source this script for the functions to become available in a session
    
    . .\Files\Functions_Intune.ps1

  Requires version 3 or higher.
  
#>

# Make sure this script is being dot-sourced.
#if (-not ($MyInvocation.InvocationName -eq '.' -or $MyInvocation.Line -eq '')) { Throw "Please invoke this script dot-sourced." }


function Connect-Azure {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$User
    )

    $requiredModules = @{
        "IntuneWin32App"         = "Install-Module -Name IntuneWin32App -Scope CurrentUser -Force";
        "AzureAD"                = "Install-Module -Name AzureAD -Scope CurrentUser -Force";
        "Microsoft.Graph.Groups" = "Install-Module -Name Microsoft.Graph.Groups -Scope CurrentUser -Force";
        "Microsoft.Graph.Intune" = "Install-Module -Name Microsoft.Graph.Intune -Scope CurrentUser -Force"
    }

    foreach ($module in $requiredModules.Keys) {
        if (-not (Get-Module -Name $module -ListAvailable)) {
            Write-Host "$module module not found. Attempting to install..."
            Invoke-Expression $requiredModules[$module]
        }
    }

    if (-not $User) {
        $User = Read-Host "Please specify your user principal name for Azure Authentication"
    }

    $userUpn = New-Object System.Net.Mail.MailAddress -ArgumentList $User
    $tenant = $userUpn.Host

    Write-Host "Connecting to Azure with tenant ID: $tenant"

    if (-not (Connect-Intune -tenant $tenant)) {
        Write-Host "Failed to connect to Azure. Exiting..."
        return $false
    }
    
    return $true
}



function Connect-Intune {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$tenant
    )

    Write-Host "Connecting to Intune..."

    # Check if appId and clientSecret are specified
    if (-not $global:AppId) {
        $message = "AppId not specified. Please execute 'Configure Intune integration'."
        Write-Warning $message
        write-log -message $message -type 3
        Write-ProgressBar `
            -ProgressBar $ProgressBar `
            -Activity "Failed to connect to Intune" `
            -PercentComplete "0" `
            -CurrentOperation $message
        return $false
    }

    # Import the IntuneWin32App module, continue silently if not found
    Import-Module IntuneWin32App -ErrorAction SilentlyContinue

    # Check if the module is loaded
    $moduleItem = Get-Module -Name IntuneWin32App

    if ($moduleItem) {
        # Attempt to connect using Connect-MSIntuneGraph with TenantID and Client credentials
        try {
            Write-Host "Attempting to connect using TenantID and Client credentials..."
            Connect-MSIntuneGraph -TenantID $tenant -ClientId $global:AppId
            #-ClientSecret $global:ClientSecret
            Write-Host "Connected successfully using TenantID and Client credentials."
            return $true
        }
        catch {
            Write-Host "Failed to connect using TenantID and Client credentials. Error: $_"
            Write-Warning "Please verify your tenant information and connectivity."
            Write-Warning "If the problem persists, please execute 'Configure Intune integration'."
            return $false
        }
    }
    else {
        $message = "IntuneWin32App module is not loaded. Please ensure it's installed and try again."
        Write-Host $message
        write-log -message $message -type 3
        return $false
    }
}






####################################################


function Install-IntuneModule {
    # Define the required modules and their installation commands
    $requiredModules = @(
        @{ Name = "IntuneWin32App"; Source = "PSGallery" },
        @{ Name = "Microsoft.Graph.Intune"; Source = "PSGallery" },
        @{ Name = "Microsoft.Graph.Authentication"; Source = "PSGallery" },
        @{ Name = "Microsoft.Graph.Applications"; Source = "PSGallery" },
        @{ Name = "Microsoft.Graph.Groups"; Source = "PSGallery" }
    )

    foreach ($module in $requiredModules) {
        $moduleName = $module.Name
        $moduleSource = $module.Source

        # Check if the module is already installed
        $installedModule = Get-Module -Name $moduleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1

        if ($installedModule) {
            Write-Host "$moduleName is already installed. Checking for updates..."
            # Call the function to check for and install updates
            CheckAndUpdate-Module -ModuleName $moduleName
        }
        else {
            Write-Host "$moduleName is not installed. Installing..."
            # Install the module from the specified source
            Install-Module -Name $moduleName -Repository $moduleSource -Force -Confirm:$false
        }
    }

    # Import the primary module to use in the script
    Import-Module IntuneWin32App
    Write-Host "Required modules are installed and updated."
    Write-Host "Start CheckAndCreate-EnterpriseApplication..."
    Force-ConnectGraph
    $global:AppId = CheckAndCreate-EnterpriseApplication -applicationName $global:AzApplicationName -tenantDomain $global:TenantName
    Write-Host "Intune Installation and configuration done!"

}

function Install-WingetClient {
    # Remove Cobalt module if installed to prevent conflicts
    if (Get-Module -ListAvailable -Name "Cobalt") {
        Write-Warning "⚠️ Conflicting module 'Cobalt' detected. Uninstalling to prevent issues with official Winget module..."
        try {
            Uninstall-Module -Name "Cobalt" -AllVersions -Force -ErrorAction Stop
            Write-Host "✅ Cobalt module uninstalled successfully." -ForegroundColor Yellow
        }
        catch {
            Write-Warning "❌ Failed to uninstall Cobalt module: $($_.Exception.Message)"
        }
    }

    try {
        # Check and install/update WinGet PowerShell module
        CheckAndUpdate-Module -ModuleName "Microsoft.WinGet.Client"
        Import-Module Microsoft.WinGet.Client -Force -ErrorAction SilentlyContinue
        Write-Host "✅ Winget Client module is ready to use." -ForegroundColor Green
    }
    catch {
        Write-Warning "❌ Failed to install or update Winget Client module: $($_.Exception.Message)"
    }

    try {
        # Check and install/update powershell-yaml module
        CheckAndUpdate-Module -ModuleName "powershell-yaml"
        Import-Module powershell-yaml -Force -ErrorAction SilentlyContinue
        Write-Host "✅ YAML module is ready to use." -ForegroundColor Green
    }
    catch {
        Write-Warning "❌ Failed to install or update powershell-yaml module: $($_.Exception.Message)"
    }
}



function CheckAndUpdate-Module {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName
    )

    # Get the currently installed version of the module
    $installedModule = Get-Module -Name $ModuleName -ListAvailable | Sort-Object Version -Descending | Select-Object -First 1

    if ($installedModule) {
        Write-Host "Currently installed version of $ModuleName is $($installedModule.Version). Checking for updates..."

        # Find the latest version available in the repository
        $latestModule = Find-Module -Name $ModuleName

        if ($latestModule.Version -gt $installedModule.Version) {
            Write-Host "An update is available. Updating $ModuleName to $($latestModule.Version)..."
            Update-Module -Name $ModuleName -Force
            Write-Host "$ModuleName updated successfully."
        }
        else {
            Write-Host "$ModuleName is already up to date."
        }
    }
    else {
        Write-Host "$ModuleName is not installed. Installing now..."
        Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber
        Write-Host "$ModuleName installed successfully."
    }
}


function Force-ConnectGraph {
    Write-Host "Forcing interactive Microsoft Graph authentication..."

    # Ensure Microsoft Graph module is loaded
    if (-not (Get-Module -Name Microsoft.Graph.Authentication)) {
        Import-Module -Name Microsoft.Graph.Authentication
    }
    if (-not (Get-Module -Name Microsoft.Graph.Applications)) {
        Import-Module -Name Microsoft.Graph.Applications
    }

    try {
        # Force the interactive login
        Connect-MgGraph -Scopes "Application.ReadWrite.All", "Directory.ReadWrite.All", "User.Read.All" -ErrorAction Stop
        Write-Host "Successfully connected to Microsoft Graph."
    }
    catch {
        Write-Host "Failed to connect to Microsoft Graph: $_"
        throw $_
    }
}



function Force-ConnectAzureAD {
    Write-Host "Forcing interactive Azure AD authentication..."

    # Ensure AzureAD module is loaded
    if (-not (Get-Module -Name AzureAD)) {
        Import-Module -Name AzureAD
    }

    try {
        # Force the interactive login
        Connect-AzureAD -ErrorAction Stop
        Write-Host "Successfully connected to Azure AD."
    }
    catch {
        Write-Host "Failed to connect to Azure AD: $_"
        throw $_
    }
}



function CheckAndCreate-EnterpriseApplication {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$applicationName,
        [Parameter(Mandatory = $true)]
        [string]$tenantDomain,
        [Parameter(Mandatory = $false)]
        [string]$clientSecret
    )
    Write-Host "Checking and creating the enterprise application..."

    # Define the application details
    $appDetails = @{
        DisplayName            = $applicationName
        SignInAudience         = "AzureADMyOrg"
        PublicClient           = @{
            RedirectUris = @("https://login.microsoftonline.com/common/oauth2/nativeclient")
        }
        RequiredResourceAccess = @(
            @{
                ResourceAppId  = "00000003-0000-0000-c000-000000000000"
                ResourceAccess = @(
                    @{
                        Id   = "7b3f05d5-f68c-4b8d-8c59-a2ecd12f24af"
                        Type = "Scope"
                    },
                    @{
                        Id   = "e1fe6dd8-ba31-4d61-89e7-88639da4683d"
                        Type = "Scope"
                    }
                )
            }
        )
    }

    # Check if application already exists
    $existingApps = Get-MgApplication -Filter "displayName eq '$applicationName'"
    $existingApp = $existingApps | Where-Object { $_.DisplayName -eq $applicationName }

    if ($existingApp) {
        Write-Host "Application with the name '$applicationName' already exists. Using the existing application."
        $appId = [string]$existingApp.AppId
        $existingApp | Format-List

        <# Retrieve existing passwords
        $passwordCredentials = Get-MgApplicationPasswordCredential -ApplicationId $existingApp.Id
        if ($passwordCredentials) {
            foreach ($password in $passwordCredentials) {
                Write-Host "Existing Password Display Name: $($password.DisplayName), End Date: $($password.EndDateTime)"
            }
        }
        #>
    }
    else {
        Write-Host "Application with the name '$applicationName' does not exist. Creating a new application."

        # Create new application
        $newApp = New-MgApplication -DisplayName $appDetails.DisplayName -SignInAudience $appDetails.SignInAudience -PublicClient $appDetails.PublicClient -RequiredResourceAccess $appDetails.RequiredResourceAccess


        # Output the new application's details
        $newApp | Format-List

        # Retrieve the AppId of the newly created application
        $appId = [string]$newApp.AppId

        # Add client secret using the provided clientSecret
        $passwordCred = @{
            displayName = 'Default'
            endDateTime = (Get-Date).AddYears(1)
        }
        #$newClientSecret = Add-MgApplicationPassword -ApplicationId $newApp.Id -PasswordCredential $passwordCred
        #Write-Host "New Client Secret: $($newClientSecret.SecretText)"
    }

    # Save the AppId back to the INI file
    Set-IniValue $configFile 'Azure' 'AppId' $appId
    #Set-IniValue $configFile 'Azure' 'ClientSecret' $($newClientSecret.SecretText)

    return $appId
}





               

Function Create-AADGroup {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$DisplayName,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Description
    )

    # Ensure you're connected to Microsoft Graph
    Write-Host "Connecting to Microsoft Graph..."
    Connect-MgGraph -Scopes "Group.ReadWrite.All"

    # Check if Group Exists
    Write-Host "Checking if Group already exists..."
    $GroupExists = Get-MgGroup -Filter "displayName eq '$DisplayName'"

    if ($GroupExists) {
        Write-Host "Group $($DisplayName) already Exists, using existing one" -ForegroundColor Yellow
        $IntuneInstallGroup = $GroupExists
    }
    else {
        Write-Host "Creating new Group..."
        # Create a new group using Microsoft Graph SDK
        $groupParams = @{
            DisplayName     = $DisplayName
            MailEnabled     = $false
            MailNickname    = "NotSet"
            SecurityEnabled = $true
            Description     = "AutoGenerated Group by <Cancom Austria Application Manager> for $($Appfullname) deployment"
        }
        $IntuneInstallGroup = New-MgGroup @groupParams
    }

    return $IntuneInstallGroup
}



Function Create-IntuneApp {  


    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Publisher,
        [string]$DisplayName,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ContentSourcePath,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallationProgram,
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$UninstallationProgram,
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Description,
        [string]$DetectionType = "Script",
        [string]$DetectionMethod = 'if (Test-Path C:\DummyDetectionMethod) {Write-Host "IMPORTANT! This detection method does not work. You must manually change it."}',
        [string]$ProductCode,
        [string]$ProductVersion,
        [string]$FilePath,                # E.g. C:\Program Files\MyApp
        [string]$FileOrFolder,            # E.g. app.exe
        [string]$FileDetectionMode = "existence", # one of: existence, version, size, datecreated, datemodified
        [string]$FileDetectionOperator = "exists", # for existence or version check
        [string]$FileDetectionValue = $null,       # version, size, or datetime
        [bool]$FileCheck32BitOn64System = $false
    )

    

    if (!$Description) {
        $Description = 'AutoGenerated by Cancom Austria Application Manager'
    }
   
    If (!(test-path $IntuneOutputFolder)) {
        New-Item -ItemType Directory -Force -Path $IntuneOutputFolder
    }

    #Check if App with specified Name already exiss
    $IntuneApp = Get-IntuneWin32App -DisplayName $DisplayName -Verbose
    if (!$IntuneApp) {
        $IntuneWinAppUtilPath = Join-Path $workingdir "Files\IntuneWinAppUtil.exe"
        $NewIntuneWinFile = Join-Path $IntuneOutputFolder "$DisplayName.intunewin"

        # Ensure the file is (re)created, even if it already exists
        try {
            $IntuneWinFile = New-IntuneWin32AppPackage `
                -SourceFolder $ContentSourcePath `
                -SetupFile $InstallationProgram `
                -OutputFolder $IntuneOutputFolder `
                -IntuneWinAppUtilPath $IntuneWinAppUtilPath `
                -Force `
                -Verbose
        } catch {
            Write-Warning "❌ Failed to create IntuneWin package: $_"
            return
        }

        # Validate the package object
        if (-not $IntuneWinFile -or -not $IntuneWinFile.Path -or !(Test-Path $IntuneWinFile.Path)) {
            Write-Warning "❌ New-IntuneWin32AppPackage did not return a valid file. Possibly skipped due to existing file or incorrect parameters."
            return
        }

        if (!(Test-Path $NewIntuneWinFile -PathType Leaf)) {
            Write-Host "Final .intunewin file doesn't exist yet — renaming generated file..."
            try {
                Rename-Item -Path $IntuneWinFile.Path -NewName "$DisplayName.intunewin" -Force
                $IntuneWinFile = Get-Item $NewIntuneWinFile
                Write-Host "Renamed file to: $($IntuneWinFile.FullName)"
            } catch {
                Write-Warning "❌ Failed to rename generated file: $_"
                return
            }
        } else {
            Write-Host "Final .intunewin file already exists — backing up and replacing..."

            $timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
            $backupName = "$DisplayName_$timestamp.backup.intunewin"
            $backupPath = Join-Path $IntuneOutputFolder $backupName

            try {
                Rename-Item -Path $NewIntuneWinFile -NewName $backupName -Force
                Write-Host "Backed up existing file to: $backupPath"
            } catch {
                Write-Warning "❌ Failed to back up existing .intunewin file: $_"
                return
            }

            try {
                Rename-Item -Path $IntuneWinFile.Path -NewName "$DisplayName.intunewin" -Force
                $IntuneWinFile = Get-Item $NewIntuneWinFile
                Write-Host "Renamed file to: $($IntuneWinFile.FullName)"
            } catch {
                Write-Warning "❌ Failed to rename generated file after backup: $_"
                return
            }
        }
    

    


        <#
    
        if ($TextBoxMSIPackage.Text.Length -gt 0) {
            Write-Host "ProductCode: $($ProductCode)"
            # Create detection rule using the en-US MSI product code (1033 in the GUID below correlates to the lcid)
            $DetectionRule = New-IntuneWin32AppDetectionRuleMSI -ProductCode $ProductCode -ProductVersionOperator greaterThanOrEqual -ProductVersion $ProductVersion
            #New-IntuneWin32AppDetectionRuleScript
        }
        else {
            #Create Dummy Detection
            $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry -DetectionType exists -Existence -KeyPath HKEY_LOCAL_MACHINE\SOFTWARE\ToBeEdited -Confirm
        }
        #>

        switch ($DetectionType.ToLower()) {
            "msi" {
                if (-not $ProductCode -or -not $ProductVersion) {
                    throw "ProductCode and ProductVersion must be provided for MSI detection"
                }
                Write-Host "Using MSI detection rule with ProductCode: $ProductCode"
                $DetectionRule = New-IntuneWin32AppDetectionRuleMSI `
                    -ProductCode $ProductCode `
                    -ProductVersionOperator greaterThanOrEqual `
                    -ProductVersion $ProductVersion
            }

            "registry" {
                Write-Host "Using default registry detection rule"
                $DetectionRule = New-IntuneWin32AppDetectionRuleRegistry `
                    -DetectionType exists `
                    -Existence `
                    -KeyPath "HKEY_LOCAL_MACHINE\SOFTWARE\ToBeEdited" `
                    -Confirm
            }

            "script" {
                # Check for existing detection script in SupportFiles
                $supportFilesPath = Join-Path $ContentSourcePath "SupportFiles"
                $existingDetectScript = Get-ChildItem -Path $supportFilesPath -Filter "Detect-*.ps1" -File -ErrorAction SilentlyContinue | Select-Object -First 1

                if ($existingDetectScript) {
                    Write-Host "Found existing detection script in SupportFiles: $($existingDetectScript.FullName)"
                    $scriptPath = $existingDetectScript.FullName
                } else {
                    # Fall back to writing the provided DetectionMethod string to a temp file
                    $scriptFolder = Join-Path $env:TEMP "IntuneDetectionScripts"
                    if (!(Test-Path $scriptFolder)) {
                        New-Item -Path $scriptFolder -ItemType Directory -Force | Out-Null
                    }

                    $scriptPath = Join-Path $scriptFolder "$($DisplayName)_Detection.ps1"

                    Write-Host "No existing script found. Writing detection script to: $scriptPath"
                    $DetectionMethod | Out-File -FilePath $scriptPath -Encoding UTF8 -Force

                    if (!(Test-Path $scriptPath)) {
                        throw "Failed to create detection script at: $scriptPath"
                    }
                }

                Write-Host "Using script detection rule with file: $scriptPath"
                $DetectionRule = New-IntuneWin32AppDetectionRuleScript `
                    -ScriptFile $scriptPath `
                    -EnforceSignatureCheck $false `
                    -RunAs32Bit $false
            }


            "file" {
                Write-Host "Using file-based detection rule"

                if (-not $FilePath -or -not $FileOrFolder) {
                    throw "FilePath and FileOrFolder must be provided for file detection"
                }

                $params = @{
                    Path                     = $FilePath
                    FileOrFolder            = $FileOrFolder
                    Check32BitOn64System    = $FileCheck32BitOn64System
                }

                switch ($FileDetectionMode.ToLower()) {
                    "existence" {
                        $params.Existence = $true
                        $params.DetectionType = $FileDetectionOperator  # "exists" or "doesNotExist"
                    }

                    "version" {
                        if (-not $FileDetectionValue) {
                            throw "FileDetectionValue (e.g. 1.0.0.0) must be provided for version check"
                        }
                        $params.Version = $true
                        $params.Operator = $FileDetectionOperator
                        $params.VersionValue = $FileDetectionValue
                    }

                    "size" {
                        if (-not $FileDetectionValue) {
                            throw "FileDetectionValue (e.g. 100) must be provided for size check"
                        }
                        $params.Size = $true
                        $params.Operator = $FileDetectionOperator
                        $params.SizeInMBValue = $FileDetectionValue
                    }

                    "datecreated" {
                        if (-not $FileDetectionValue) {
                            throw "FileDetectionValue (datetime) must be provided for creation date check"
                        }
                        $params.DateCreated = $true
                        $params.Operator = $FileDetectionOperator
                        $params.DateTimeValue = [datetime]$FileDetectionValue
                    }

                    "datemodified" {
                        if (-not $FileDetectionValue) {
                            throw "FileDetectionValue (datetime) must be provided for modified date check"
                        }
                        $params.DateModified = $true
                        $params.Operator = $FileDetectionOperator
                        $params.DateTimeValue = [datetime]$FileDetectionValue
                    }

                    default {
                        throw "Unsupported FileDetectionMode: $FileDetectionMode"
                    }
                }

                $DetectionRule = New-IntuneWin32AppDetectionRuleFile @params
            }

            default {
                throw "Unknown DetectionType: $DetectionType. Valid types are: MSI, Registry, Script, File."
            }
        }


        # Create custom requirement rule
        $RequirementRule = New-IntuneWin32AppRequirementRule -Architecture All -MinimumSupportedWindowsRelease "W10_1607"

        # Convert image file to icon
        #$ImageFile = "C:\IntuneWinAppUtil\Icons\AdobeReader.png"
        #$Icon = New-IntuneWin32AppIcon -FilePath $ImageFile

        #Add new EXE Win32 app
        #Check for ServiceUI (Intune workaround for User interaction)
        # Apply ServiceUI only if NOT using .bat scripts
        $usingBatchFiles = ($InstallationProgram -like "*.bat" -or $UninstallationProgram -like "*.bat")

        if (-not $usingBatchFiles) {
            # First try to get ServiceUI_x64.exe
            $serviceUIFile = Get-ChildItem -Path $ContentSourcePath -File -Include "ServiceUI_x64.exe" -ErrorAction SilentlyContinue | Select-Object -First 1

            # If not found, fall back to any other ServiceUI*
            if (-not $serviceUIFile) {
                $serviceUIFile = Get-ChildItem -Path $ContentSourcePath -File -Include "ServiceUI*" -ErrorAction SilentlyContinue | Select-Object -First 1
            }

            if ($serviceUIFile) {
                Write-Host "✅ ServiceUI found: $($serviceUIFile.Name) — wrapping install/uninstall commands..."

                $InstallationProgram = "$($serviceUIFile.Name) -process:explorer.exe $InstallationProgram"
                $UninstallationProgram = "$($serviceUIFile.Name) -process:explorer.exe $UninstallationProgram"
            } else {
                Write-Host "ℹ️ No ServiceUI executable found — using default commands."
            }
        } else {
            Write-Host "⚠️ Skipping ServiceUI wrapping because .bat files are being used"
        }

        Write-Host "Add-IntuneWin32App -FilePath $NewIntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -InstallExperience system -RestartBehavior suppress -DetectionRule $DetectionRule -RequirementRule $RequirementRule  -AppVersion $ProductVersion -Notes $ContentSourcePath -InstallCommandLine $InstallationProgram -UninstallCommandLine $UninstallationProgram -Verbose"
        $IntuneApp = Add-IntuneWin32App -FilePath $NewIntuneWinFile -DisplayName $DisplayName -Description $Description -Publisher $Publisher -InstallExperience system -RestartBehavior suppress -DetectionRule $DetectionRule -AppVersion $ProductVersion -Notes $ContentSourcePath -RequirementRule $RequirementRule -InstallCommandLine $InstallationProgram -UninstallCommandLine $UninstallationProgram -Verbose
        
        if ($CleanUpIntuneOutputFolder -eq $true) {
            Write-Host "CleanUpIntuneOutputFolder is true, remove $($DisplayName).intunewin"
            Remove-Item –path $NewIntuneWinFile
        }
    }
    else {
        Write-Host "Application $($DisplayName) already exists, use existing One" -ForegroundColor Yellow
    }

    return $IntuneApp
}



#############################################################################
#END Functions_Intune
#############################################################################

$Global:ScriptStatus = 'Success'

# We set the RunInstallAs32Bit parameter to a global variable, because it can be modified in multiple functions
$global:RunInstallAs32Bit = [bool]::Parse($RunInstallAs32Bit)
$global:DownloadOnSlowNetwork = [bool]::Parse($DownloadOnSlowNetwork)
$global:AllowFallbackSourceLocation = [bool]::Parse($AllowFallbackSourceLocation)
$global:AllowInteractionDefault = [bool]::Parse($AllowInteractionDefault)

$driveloc = Get-Location

# Import the SCCM powershell module
$CurrentLocation = Get-Location

#Set Image Source
$CompanyLogo.Source = $workingdir + "\" + "files\logo.png"

    

#############################################################################
# BEGIN Functions
#############################################################################


# Workaround for 1st run - handle non-existent path more gracefully
$Packagefolderpathfiles = "FileSystem::" + $iniFile.Global.Packagefolderpath

if (Test-Path -Path $Packagefolderpathfiles) {
    try {
        $Packagefolders = Get-ChildItem -Path $Packagefolderpathfiles | Where-Object { $_.PSIsContainer }
    }
    catch {
        Write-Warning "Failed to retrieve folders from '$Packagefolderpathfiles': $_"
    }
}
else {
    Write-Warning "Package folder path does not exist, please check you config: $Packagefolderpathfiles"
    $Packagefolders = @() # define as empty to avoid further errors
}



#Function to write logfile in CMTrace Format (TH 2016)
function write-log {
    param (
        [Parameter(Mandatory = $true)]
        $message,
        [Parameter(Mandatory = $true)]
        $type = 1 )

    $Currentprov = get-location
  
    set-location "c:"
     
    switch ($type) {
        1 { $type = "Info" }
        2 { $type = "Warning" }
        3 { $type = "Error" }
        4 { $type = "Verbose" }
    }
  
    $component = $env:USERNAME

    If (!(test-path (Split-Path -Path $logfile))) {
        New-Item -ItemType Directory -Force -Path (Split-Path -Path $logfile)
    }

    if (($type -eq "Verbose") -and ($Global:Verbose)) {
        $toLog = "{0} `$$<{1}><{2} {3}><thread={4}>" -f ($type + ":" + $message), ($component), (Get-Date -Format "MM-dd-yyyy"), (Get-Date -Format "HH:mm:ss.ffffff"), $pid
        $toLog | Out-File -Append -Encoding UTF8 -FilePath $logfile
        #Write-Host $message
    }
    elseif ($type -ne "Verbose") {
        $toLog = "{0} `$$<{1}><{2} {3}><thread={4}>" -f ($type + ":" + $message), ($component), (Get-Date -Format "MM-dd-yyyy"), (Get-Date -Format "HH:mm:ss.ffffff"), $pid
        
        
        
        ####Disabled because of Transcript Output
        #$toLog | Out-File -Append -Encoding UTF8 -FilePath $logfile


        #Write-Host $message
    }
    if (($type -eq 'Warning') -and ($Global:ScriptStatus -ne 'Error')) { $Global:ScriptStatus = $type }
    if ($type -eq 'Error') { $Global:ScriptStatus = $type }

    if ((Get-Item $logfile).Length / 1KB -gt $iniFile.Global.MaxLogSizeInKB) {
        $log = $logfile
        Remove-Item ($log.Replace(".log", ".lo_"))
        Rename-Item $logfile ($log.Replace(".log", ".lo_")) -Force
    }

    set-location $Currentprov

} 



#Control ErrorHandling (Set Or Clear)
function ErrorHandler {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        $Object,
 
        [parameter(Mandatory = $true)]
        $Handling
    )

    if ($Handling -eq 'Set') {
        $($Object).BorderBrush = "#ff0000"
    }
    else {
        $($Object).BorderBrush = "#FFABADB3"
    }
    

}





function Write-ProgressBar {

    Param (
        [Parameter(Mandatory = $true)]
        $ProgressBar,
        [Parameter(Mandatory = $true)]
        [String]$Activity,
        [int]$PercentComplete,
        [String]$Status = $Null,
        [int]$SecondsRemaining = $Null,
        [String]$CurrentOperation = $Null
    ) 
   
    Write-Verbose -Message "Setting activity to $Activity"
    $ProgressBar.Activity = $Activity

    if ($PercentComplete) {
       
        Write-Verbose -Message "Setting PercentComplete to $PercentComplete"
        $ProgressBar.PercentComplete = $PercentComplete

    }
   
    if ($SecondsRemaining) {

        [String]$SecondsRemaining = "$SecondsRemaining Seconds Remaining"

    }
    else {

        [String]$SecondsRemaining = $Null

    }

    Write-Verbose -Message "Setting AdditionalInfo to $Status       $SecondsRemaining$(if($SecondsRemaining){ " seconds remaining..." }else {''})       $CurrentOperation"
    $ProgressBar.AdditionalInfo = "$Status       $SecondsRemaining       $CurrentOperation"

}

function Close-ProgressBar {

    Param (
        [Parameter(Mandatory = $true)]
        [System.Object[]]$ProgressBar
    )

    $ProgressBar.Window.Dispatcher.Invoke([action] { 
      
            $ProgressBar.Window.close()

        }, "Normal")
 
}



function Start-Appimportform {
    try {
        $Packagefolderpathfiles = "FileSystem::" + $iniFile.Global.Packagefolderpath

        if (Test-Path -Path $Packagefolderpathfiles) {
            $Packagefolders = Get-ChildItem -Path $Packagefolderpathfiles | Where-Object { $_.PSIsContainer }
        }
        else {
            throw "Package folder path does not exist, please check you config: $Packagefolderpathfiles"
        }

    }
    catch {
        $ButtonCreate.IsEnabled = $false
        $Message = "Unable to connect to package repository. Error: $($_.Exception.Message)"
        $LabelOutput.Content = $Message
        Write-Host $Message -ForegroundColor Red
    }
}


# read Properties from MSI Files for Detection Methods
function Get-MsiFileInformation {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo]$Path,
 
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateSet("ProductCode", "ProductVersion", "ProductName", "Manufacturer", "ProductLanguage", "FullVersion")]
        [string]$Property
    )
    Process {
        try {
            # Read property from MSI database
            $WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
            $MSIDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $null, $WindowsInstaller, @($Path.FullName, 0))
            $Query = "SELECT Value FROM Property WHERE Property = '$($Property)'"
            $View = $MSIDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSIDatabase, ($Query))
            $View.GetType().InvokeMember("Execute", "InvokeMethod", $null, $View, $null)
            $Record = $View.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $View, $null)
            $Value = $Record.GetType().InvokeMember("StringData", "GetProperty", $null, $Record, 1)
 
            # Commit database and close view
            $MSIDatabase.GetType().InvokeMember("Commit", "InvokeMethod", $null, $MSIDatabase, $null)
            $View.GetType().InvokeMember("Close", "InvokeMethod", $null, $View, $null)           
            $MSIDatabase = $null
            $View = $null
 
            # Return the value
            return $Value
        } 
        catch {
            Write-Warning -Message $_.Exception.Message ; break
        }
    }
    End {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($WindowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}

# Functions nur für SWMapping notwendig

if ($SWMappingenabled -eq $true) {

    function query-SQLSWMAP ($querytext, $DBname) {
        $SqlQuery = $querytext
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $DBName; Integrated Security = True;"
        $SqlConnection.Open() 
        $Command = New-Object System.Data.SQLClient.SQLCommand 
        $Command.Connection = $SqlConnection 
        $Command.CommandText = $SQLQuery 
        $DataAdapter = new-object System.Data.SqlClient.SqlDataAdapter $Command 
        $Dataset = new-object System.Data.Dataset 
        $DataAdapter.Fill($Dataset) 
        $SqlConnection.Close() 
      
        if ($Dataset.Tables[0]) {
            $result = $Dataset.Tables[0]
            return $result
        } 


    }


    function Insert-SQLSWProduct ($SWPName, $SWPVersion, $SWInstallmethod, $SWPOwner, $SWPADGroup, $CatalogID) {
 
        $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
        $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SWMapDBName; Integrated Security = True;"
        $SqlConnection.Open() 
        $SqlCommand = New-Object System.Data.SQLClient.SQLCommand 
        $SqlCommand.Connection = $SqlConnection
        $SqlCommand.CommandText = "INSERT INTO $SWProductTable (DKPID,SWPName,SWPVersion,Installationsmethode,Applikationsverantwortlicher,ADGruppe,Quelle) VALUES ('$CatalogID','$SWPName','$SWPVersion','$SWInstallmethod','$SWPOwner','$SWPADGroup','Applicationimport')"
        # Run the query and get the scope ID back into $InsertedID

        $InsertedID = $sqlCommand.ExecuteScalar()

        # Write to the console.

        write-host "Inserted row ID $InsertedID into SWP"

   

    }
}

# Function with actions when form is loaded
function Load-Form {

    $ButtonCreate.IsEnabled = $False
    $RadioButtonAvailable.IsEnabled = $false
    $RadioButtonRequired.IsEnabled = $false
    $TextBoxMSIPackage.Clear()
    #$CheckboxStandardapp.text = ""
    #$TextBoxSWCatalogID.Clear()
    $TextBoxAppName.Clear()
    $TextBoxPublisher.Clear()
    $TextBoxVersion.Clear()
    $TextBoxInstallProgram.Clear()
    $TextBoxInstallProgram.IsEnabled = $true
    $TextBoxUnInstallProgram.Clear()
    $TextBoxUnInstallProgram.IsEnabled = $true
    $TextBoxSourcePath.Clear()
    $TextBoxSourcePath.IsEnabled = $true
    $TextBoxADGroup.Clear()
    $TextBoxADGroup.IsEnabled = $true
    $RadioButtonRequired.IsChecked = $true
    $RadioButtonDevice.IsChecked = $true
    $CheckBoxCreateCollection.IsChecked = $false
    $CheckBoxCreateCollection.IsEnabled = $false
    $CheckBoxCreateADGroup.IsChecked = $false
    $CheckBoxCreateADGroup.IsEnabled = $false
    $CheckBoxDistributeContent.IsChecked = $false
    $CheckBoxDistributeContent.IsEnabled = $false
    $CheckBoxCreateDeployment.IsChecked = $false
    $CheckBoxCreateDeployment.IsEnabled = $false
}

# Function to validate all input in the form. Initiated by clicking the Create-button, and must return True before an application is created.
function Validate-Form {
    if ($checkboxCreateInConfigMgr.IsChecked) {
        try {
            Import-Module(Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) | Out-Null
            Set-Location($iniFile.ConfigMgr.SiteCode + ":")
        }
        catch {
            
            $Message = "no Connection to ConfigMgr Provider established..."
            #If no Connection to ConfigMgr Uncheck and Disable ConfigMgr Checkbox
            $checkboxCreateInConfigMgr.IsChecked = $false
            $checkboxCreateInConfigMgr.IsEnabled = $false
            $CheckBoxCreateCollection.IsChecked = $false
            $CheckBoxCreateCollection.IsEnabled = $false
            $CheckBoxDistributeContent.IsChecked = $false
            $CheckBoxDistributeContent.IsEnabled = $false
            $TextBoxCollection.IsEnabled = $false
            $TextBoxCollection.Text = $Message
            $LabelOutput.Content = $Message
            Write-Host $Message -ForegroundColor DarkYellow
            
        }
    }


    $OkToProceed = $true
    $MSIPackageName = $TextBoxMSIPackage.Text
    $MSTFileName = $TextBoxMSTFile.Text
    $SWCatalogID = $TextBoxSWCatalogID.Text
    $ApplicationVersion = $TextBoxVersion.Text
    if ($ApplicationVersion -ne "" -and $ApplicationVersion -ne $null) {
        $ApplicationName = $TextBoxAppName.Text + " " + $ApplicationVersion
    }
    else {
        $ApplicationName = $TextBoxAppName.Text
    }
    $Publisher = $TextBoxPublisher.Text
    $ApplicationVersion = $TextBoxVersion.Text
    $InstallationProgram = $TextBoxInstallProgram.Text
    $UninstallationProgram = $TextBoxUnInstallProgram.Text
    $ContentSourcePath = $TextBoxSourcePath.Text
    $InstallCollectionName = $TextBoxInstallCollection.Text
    $ADGroupName = $TextBoxADGroup.Text
    $CollectionName = $TextBoxCollection.Text

    # Clear the error providers
  

    # Check if an Application exists in SCCM with the same name
    <#   if ((Check-ApplicationExist $ApplicationName) -eq $true)
    {
    	$OkToProceed = $false
        $logtext = "An application called $ApplicationName already exists. Please check the name and try again."
        $ErrorProviderAppName.SetError($TextBoxAppName, $logtext)
        write-log -message $logtext -type 3
    }
#>
    # Check if we try to create a new collection with the same name as one that already exists
    <#  if ($TextBoxCollection.Text.Length -gt 0 -and (Check-CollectionExist $CollectionName) -eq $true -and $CheckBoxCreateCollection.IsChecked -eq $true)
    {
    	$OkToProceed = $false	
        $logtext = "A collection called $CollectionName already exists. Please change the name or clear the Create Collection checkbox."
        $ErrorProviderCollection.SetError($TextBoxCollection, $logtext)
        write-log -message $logtext -type 3
    }
  #>
    <#
    if ($checkboxCreateInConfigMgr.IsChecked) {
        # Check if a collection name is specified, but the Create Collection checkbox is unchecked, and the collection does not already exist
        if ($CollectionName.Length -gt 0 -and (Check-CollectionExist $CollectionName) -eq $false -and $CheckBoxCreateCollection.IsChecked -eq $false) {
            $OkToProceed = $false
            $logtext = "The collection $CollectionName does not exist. Please clear the collection name or change it to the name of an existing collection, or check the Create Collection checkbox to create a new collection."
            $LabelOutput.Content = $logtext
            $ErrorProviderCollection.SetError($TextBoxCollection, $logtext)
            write-log -message $logtext -type 3
        }
    }
    #>

    <# # Check if we try to create a new AD-group with the same name as one that already exists
    if ($TextBoxADGroup.Text.Length -gt 0 -and (Check-ADGroupExist $ADGroupName) -eq $true -and $CheckBoxCreateADGroup.IsChecked -eq $true)
    {
    	$OkToProceed = $false
        $logtext = "An AD Group called $ADGroupName already exists. Please change the name or clear the Create AD Group checkbox."
        $ErrorProviderADGroup.SetError($TextBoxADGroup, $logtext)
        write-log -message $logtext -type 3
    }
    #>
    # Validate that the path to the MSI-package is a UNC-path
    if ($TextBoxMSIPackage.Text.Length -gt 0) {
        if (-not $MSIPackageName.StartsWith("\\")) {
            $OkToProceed = $false
            $logtext = "Local paths are not supported. Please specify a UNC-path."
            $LabelOutput.Content = $logtext
            ErrorHandler -Object $TextBoxMSIPackage -Handling 'Set'
            write-log -message $logtext -type 3
            Write-Host $logtext
        }
    }

    

    if ($SWMappingenabled -eq $true) {
        # Validate that a Catalog ID is entered
    
        if ($TextBoxSWCatalogID.Text.Length -eq 0) {
            $OkToProceed = $false
            $logtext = "SW Catalog ID must not be empty"
            $LabelOutput.Content = $logtext
            $ErrorProviderSWCatalog.SetError($TextBoxSWCatalogID, $logtext)
            write-log -message $logtext -type 3
        }
    }
    # Validate that the content source path is a UNC-path
    if ($TextBoxSourcePath.Text.Length -gt 0) {
        if (-not $ContentSourcePath.StartsWith("\\")) {
            $OkToProceed = $false
            $logtext = "Local paths are not supported. Please specify a UNC-path."
            $LabelOutput.Content = $logtext
            $ErrorProviderSourcePath.SetError($TextBoxSourcePath, $logtext)
            ErrorHandler -Object $ContentSourcePath -Handling 'Set'
            write-log -message $logtext -type 3
        }
    }

    # If only the content source path is specified but not the installation program, we can't proceed
    if ($TextBoxSourcePath.Text.Length -gt 0 -and $TextBoxInstallProgram.Text -eq "") {
        $OkToProceed = $false
        $logtext = "If you specify a content source path, you must also specify an installation program."
        $ErrorProviderInstallProgram.SetError($TextBoxInstallProgram, $logtext)
        $LabelOutput.Content = $logtext
        ErrorHandler -Object $TextBoxInstallProgram -Handling 'Set'
        write-log -message $logtext -type 3
    }

    # If only the installation program is specified but not the content source path, we can't proceed
    if ($TextBoxInstallProgram.Text -gt 0 -and $TextBoxSourcePath.Text -eq "") {
        $OkToProceed = $false
        $logtext = "If you specify an installation program, you must also specify a content source path."
        $ErrorProviderSourcePath.SetError($TextBoxSourcePath, $logtext)
        ErrorHandler -Object $TextBoxSourcePath -Handling 'Set'
        $LabelOutput.Content = $logtext
        write-log -message $logtext -type 3
    }

    # If we're ok to proceed, return True
    if ($OkToProceed) {
        Return $true
    }
    else {
        Return $false
    }
}

#Get Infos from selected Application
function ButtonLoadPackageInfo {
    param(
        $Packagename,
        $Path
    )
    if (!$Packagename) {
        $Packagename = $DDPackages.SelectedItem.ToString()
    }

    if (!$Path) {
        $Packagefullname = $iniFile.Global.Packagefolderpath + "\" + $Packagename
    }
    else {
        $Packagefullname = $Path
    }


    $Packagevendor = $Packagename.Split($iniFile.Global.PackageDelimiter)[0]
    $PackageApplication = $Packagename.Split($iniFile.Global.PackageDelimiter)[1]
    $PackageVersion = $Packagename.Split($iniFile.Global.PackageDelimiter)[2]
  
    write-host "Path: " $Packagefullname
    write-host "Packagename: " $Packagename
    write-host "Vendor: " $Packagevendor
    write-host "App: " $PackageApplication
    write-host "Version: " $Packageversion

    $Packagefullpath = "filesystem::" + $Packagefullname
  
 
    $look4MSI = Get-ChildItem -Path $Packagefullpath -Include "*.msi" -Recurse
    if ($look4MSI.count -gt 0) {
        $TextBoxMSIPackage.text = $look4MSI[0].FullName
        $LabelOutput.Font.Bold
        #$LabelOutput.ForeColor = "Red"
        $LabelOutput.Content = "MSI Detectionrule found... PLEASE CHECK IF CORRECT!"
                                
    }

 
    $TextBoxAppName.Text = $Packagevendor + " " + $PackageApplication
    $TextBoxPublisher.Text = $Packagevendor
    $TextBoxVersion.Text = $Packageversion
    #$path = 'filesystem::' + $packagefullname + "\_SW_silent.vbs"
     

    # Check for Deployment Wrappers
    $installScriptPath = 'filesystem::' + $packagefullname
    $cmdFile = Get-ChildItem -Path $installScriptPath -Include "*.cmd" -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
    $ps1File = Get-ChildItem -Path $installScriptPath -Include "*.ps1" -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne "Deploy-Application.ps1" } | Select-Object -First 1

    if (Test-Path ($installScriptPath + "\_SW_silent.vbs")) {
        $TextBoxInstallProgram.Text = "cscript.exe _SW_silent.vbs"
        $TextBoxUnInstallProgram.Text = "cscript.exe _SW_silent.vbs Uninstall"
    }
    elseif (Test-Path ($installScriptPath + "\Deploy-Application.exe")) {
        $TextBoxInstallProgram.Text = "Deploy-Application.exe"
        $TextBoxUnInstallProgram.Text = "Deploy-Application.exe Uninstall"
    }
    elseif (Test-Path ($installScriptPath + "\Invoke-AppDeployToolkit.exe")) {
        $TextBoxInstallProgram.Text = "Invoke-AppDeployToolkit.exe"
        $TextBoxUnInstallProgram.Text = "Invoke-AppDeployToolkit.exe Uninstall"
    }
    elseif ($cmdFile) {
        $TextBoxInstallProgram.Text = $cmdFile.Name
        $TextBoxUnInstallProgram.Text = "tbd"
    }
    elseif ($ps1File) {
        $TextBoxInstallProgram.Text = "powershell.exe -ExecutionPolicy Bypass -File `"$($ps1File.Name)`""
        $TextBoxUnInstallProgram.Text = "tbd"
    }
    else {
        $TextBoxInstallProgram.Text = 'tbd'
        $TextBoxUnInstallProgram.Text = 'tbd'
    }

    $TextBoxSourcePath.Text = $Packagefullname



    #
    #$CheckBoxCreateADGroup.IsChecked = $true
    $RadioButtonDevice.IsChecked = $true
    $RadioButtonRequired.IsChecked = $true
    if ($AllowInteractionDefault -eq $true) { $checkboxInteraction.IsChecked = $true }


}

# Function with actions when something is typed in the application name textbox
function AppName-Changed {
    if ($TextBoxAppName.Text -ne "") {
        $TextBoxADGroup.Text = $ADGroupNamePrefix + $TextBoxAppName.Text + "_" + $TextBoxVersion.Text
        $TextBoxCollection.Text = $TextBoxAppName.Text + " " + $TextBoxVersion.Text
        #$ErrorProviderAppName.Clear()
        ErrorHandler -Object $TextBoxAppName -Handling 'Clear'
        $ButtonCreate.IsEnabled = $True
    }
    else {
        $TextBoxADGroup.Text = ""
        $TextBoxCollection.Text = ""
        ErrorHandler -Object $TextBoxAppName -Handling 'Set'
        $logtext = "Please enter a name for the application"
        #$ErrorProviderAppName.SetError($TextBoxAppName, "Please enter a name for the application")
        $LabelOutput.Content = $logtext
        $ButtonCreate.IsEnabled = $False
    }
}

# Function to control actions when opening an MSI file in the file dialog
function OpenMSIFile {    
    # Set variables based on properties from MSI file
    $MSIFilePath = $OpenFileDialogMSI.FileName
    Write-Host $MSIFilePath
    #set-location $iniFile.Global.Packagefolderpath
    Set-Location $CurrentLocation
    [string]$MSIFileName = (Split-Path -leaf $MSIFilePath)
    [string]$SourcePath = (Split-Path -Parent $MSIFilePath)
    [string]$ApplicationName = Get-MsiProperty $MSIFilePath "'ProductName'"
    $ApplicationName = $ApplicationName.Trim()
    Write-Host "MSIName= $ApplicationName"
    [string]$ApplicationPublisher = Get-MsiProperty $MSIFilePath "'Manufacturer'"
    $ApplicationPublisher = $ApplicationPublisher.Trim()
    [string]$ApplicationVersion = Get-MsiProperty $MSIFilePath "'ProductVersion'"
    $ApplicationVersion = $ApplicationVersion.Trim()
    [string]$ProductCode = Get-MsiProperty $MSIFilePath "'ProductCode'"
    $ProductCode = $ProductCode.Trim()
	
    # Enable and populate text boxes
    #$TextBoxInstallProgram.IsEnabled = $true
    #$TextBoxUnInstallProgram.IsEnabled = $true
    #$TextBoxSourcePath.IsEnabled = $true
    $TextBoxMSIPackage.Text = $MSIFilePath
    #$LabelMSIDetectionmethod.Text = "MSI Detectionrule set..."
    #$TextBoxMSTFile.Text = ""
    #$TextBoxAppName.Text = $ApplicationName
    #$TextBoxPublisher.Text = $ApplicationPublisher
    #$TextBoxVersion.Text = $ApplicationVersion
    #$TextBoxSourcePath.Text = $SourcePath
    #$TextBoxInstallProgram.Text = "msiexec /i ""$MSIFileName"" /q /norestart"
    #$TextBoxUnInstallProgram.Text = "msiexec /x $ProductCode /q /norestart"
	
	
    # Enable the PADT button
    $ButtonVB.IsEnabled = $true
	
    # Enable and check the checkboxes for content and deployment
    $CheckBoxDistributeContent.IsEnabled = $true
    $CheckBoxDistributeContent.IsChecked = $true
    $CheckBoxCreateDeployment.IsEnabled = $true
    $CheckBoxCreateDeployment.IsChecked = $true
    $RadioButtonAvailable.IsEnabled = $true
    $RadioButtonRequired.IsEnabled = $true
}

# Function to control actions when clicking on button 'Browse' for package selection
function ButtonBrowseAppClick {
    param(
        $Packagefullname
    )

    #~~< OpenFileDialogMSI >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $OpenFileDialogMSI = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialogMSI.Filter = "MSI Files | *.msi"
    $OpenFileDialogMSI.ShowHelp = $true
    $OpenFileDialogMSI.Title = "Select MSI File"
    $OpenFileDialogMSI.InitialDirectory = $Packagefullname
    $OpenFileDialogMSI.add_FileOK({ OpenMSIFile })
    #Param($Packagefullname ='')
    # Open the file dialog
    #$OpenFileDialogMSI.InitialDirectory = $Packagefullname
    $OpenFileDialogMSI.ShowDialog()
    
}

# Function to control actions when clicking on button 'Browse' for MSI package
function ButtonMSIClick {
    param(
        $Packagefullname
    )

    #~~< OpenFileDialogMSI >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    $OpenFileDialogMSI = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialogMSI.Filter = "MSI Files | *.msi"
    $OpenFileDialogMSI.ShowHelp = $true
    $OpenFileDialogMSI.Title = "Select MSI File"
    $OpenFileDialogMSI.InitialDirectory = $Packagefullname
    $OpenFileDialogMSI.add_FileOK({ OpenMSIFile })
    #Param($Packagefullname ='')
    # Open the file dialog
    #$OpenFileDialogMSI.InitialDirectory = $Packagefullname
    $OpenFileDialogMSI.ShowDialog()
    
}


function ButtonMSIClearClick {
    # Open the file dialog
    $TextBoxMSIPackage.Clear()

}

# Function to control actions when opening an MSI file in the file dialog
function OpenMSIFile {    
    # Set variables based on properties from MSI file
    $MSIFilePath = $OpenFileDialogMSI.FileName
    #set-location $Packagefolderpath
    Set-Location $CurrentLocation
    [string]$MSIFileName = (Split-Path -leaf $MSIFilePath)
    [string]$SourcePath = (Split-Path -Parent $MSIFilePath)
    [string]$ApplicationName = Get-MsiProperty $MSIFilePath "'ProductName'"
    $ApplicationName = $ApplicationName.Trim()
    [string]$ApplicationPublisher = Get-MsiProperty $MSIFilePath "'Manufacturer'"
    $ApplicationPublisher = $ApplicationPublisher.Trim()
    [string]$ApplicationVersion = Get-MsiProperty $MSIFilePath "'ProductVersion'"
    $ApplicationVersion = $ApplicationVersion.Trim()
    [string]$ProductCode = Get-MsiProperty $MSIFilePath "'ProductCode'"
    $ProductCode = $ProductCode.Trim()
	
    # Enable and populate text boxes
    #$TextBoxInstallProgram.IsEnabled = $true
    #$TextBoxUnInstallProgram.IsEnabled = $true
    #$TextBoxSourcePath.IsEnabled = $true
    $TextBoxMSIPackage.Text = $MSIFilePath
    #$LabelMSIDetectionmethod.Text = "MSI Detectionrule set..."
    #$TextBoxMSTFile.Text = ""
    #$TextBoxAppName.Text = $ApplicationName
    #$TextBoxPublisher.Text = $ApplicationPublisher
    #$TextBoxVersion.Text = $ApplicationVersion
    #$TextBoxSourcePath.Text = $SourcePath
    #$TextBoxInstallProgram.Text = "msiexec /i ""$MSIFileName"" /q /norestart"
    #$TextBoxUnInstallProgram.Text = "msiexec /x $ProductCode /q /norestart"
	
	
    # Enable the PADT button
    #$ButtonVB.IsEnabled = $true
	
    # Enable and check the checkboxes for content and deployment
    $CheckBoxDistributeContent.IsEnabled = $true
    $CheckBoxDistributeContent.IsChecked = $true
    $CheckBoxCreateDeployment.IsEnabled = $true
    $CheckBoxCreateDeployment.IsChecked = $true
    $RadioButtonAvailable.IsEnabled = $true
    $RadioButtonRequired.IsEnabled = $true
}

# Function to get properties from an MSI package
function Get-MsiProperty {
    param(
        [string]$Path,
        [string]$Property
    )
	    
    function Get-Property($Object, $PropertyName, [object[]]$ArgumentList) {
        return $Object.GetType().InvokeMember($PropertyName, 'Public, Instance, GetProperty', $null, $Object, $ArgumentList)
    }
	 
    function Invoke-Method($Object, $MethodName, $ArgumentList) {
        return $Object.GetType().InvokeMember($MethodName, 'Public, Instance, InvokeMethod', $null, $Object, $ArgumentList)
    }
	 
    $ErrorActionPreference = 'Stop'
    Set-StrictMode -Version Latest
	 
    $msiOpenDatabaseModeReadOnly = 0
    $Installer = New-Object -ComObject WindowsInstaller.Installer
	 
    $Database = Invoke-Method $Installer OpenDatabase @($Path, $msiOpenDatabaseModeReadOnly)
	 
    $View = Invoke-Method $Database OpenView  @("SELECT Value FROM Property WHERE Property=$Property")
	 
    Invoke-Method $View Execute
	 
    $Record = Invoke-Method $View Fetch
    if ($Record) {
        Write-Output(Get-Property $Record StringData 1)
    }
	 
    Invoke-Method $View Close @( )
    Remove-Variable -Name Record, View, Database, Installer
	 
}

# Function to check if an application exists
function Check-ApplicationExist {
    param(
        [Parameter(
            Position = 0)]
        $AppName
    )
	
    if (Get-CMApplication -Name $AppName) {
        return $true
    }
    else {
        return $false
    }
}

# Function to check if a collection exists
function Check-CollectionExist {
    param(
        [Parameter(
            Position = 0)]
        $CollectionName
    )
	
    if (Get-CMCollection -Name $CollectionName) {
        return $true
    }
    else {
        return $false
    }
}

# Function to check if an AD group exists
function Check-ADGroupExist {
    param(
        [Parameter(
            Position = 0)]
        $ADGroupName
    )
	
    $GroupExist = Get-ADGroup -Filter { name -eq $ADGroupName }
    if ($GroupExist) {
        return $true
    }
    else {
        return $false
    }
}

# Function to create a random time stamp
function Random-StartTime {
    [string]$RandomHour = (Get-Random -Maximum 12) 
    [string]$RandomMinute = (Get-Random -Maximum 59)
    [string]$RandomStartTime = $RandomHour + ":" + $RandomMinute
    return $RandomStartTime
}

function LabelSelectDistributionPointClick {
    $SelectedDPGroup = Get-CMDistributionPointGroup | Select-Object -Property Name, Description, MemberCount |
    Sort-Object -Property Name -Descending | Out-GridView -Title "Select DistributionPoint Group" -OutputMode Multiple
    if ($SelectedDPGroup -and $ConfigMgrConnect) {
        $CheckBoxDistributeContent.IsChecked = $false  
    }
}

function CheckBoxDistributeContentClick {
    if ($ConfigMgrConnect) {
        $CheckBoxCreateDeployment.IsEnabled = $true
    }
}

# Function to control actions when the distribute content checkbox is checked or unchecked
function CheckBoxDistributeContentChanged {
    # If checkbox to distribute content is not selected, then clear selection to create deployment
    if ($CheckboxDistributeContent.IsChecked -eq $false) {
        $CheckBoxCreateDeployment.IsChecked = $false
    }
}

# Function to control actions when the create deployment checkbox is checked or unchecked
function CheckBoxCreateDeploymentChanged {
    # If a deployment is selected to be created, then also create a collection and distribute the content
 
    if ($CheckboxCreateDeployment.IsChecked) {
        $CheckBoxCreateCollection.IsChecked = $true
        $CheckBoxDistributeContent.IsChecked = $true
        $RadioButtonAvailable.IsEnabled = $true
        $RadioButtonRequired.IsEnabled = $true
    }
    else {
        #$RadioButtonAvailable.IsEnabled = $false
        #$RadioButtonRequired.IsEnabled = $false
    }
}

# Function to control actions when the create AD Group checkbox is checked or unchecked
function CheckBoxCreateADGroupChanged {
    # Currently not in use
}

# Function with actions when the MSI-package textbox is changed
function MSIFile-Changed {
    #$ErrorProviderMSIPackage.Clear()
    ErrorHandler -Object $TextBoxMSIPackage -Handling 'Clear'
    # Update textboxes
    if ($TextBoxMSIPackage.Text.Length -gt 0) {
        #$TextBoxAppVPackage.Clear()
        $TextBoxInstallProgram.IsEnabled = $true
        $TextBoxUnInstallProgram.IsEnabled = $true
        $TextBoxSourcePath.IsEnabled = $true
    }

}



# Function with actions when the collection name textbox is changed
function CollectionName-Changed {
    ErrorHandler -Object $TextBoxCollection -Handling 'Clear'
    if ($TextBoxCollection.Text -eq "") {
        # If no name for the collection has been specified, we disable the controls to create a collection and a deployment
        $CheckBoxCreateDeployment.IsChecked = $false
        $CheckBoxCreateDeployment.IsEnabled = $false
        $CheckBoxCreateCollection.IsChecked = $false
        $CheckBoxCreateCollection.IsEnabled = $false
    }
    else {
        If ($ConfigMgrConnect) {
            # If a collection name is specified, we enable the control to create a collection
            $CheckBoxCreateCollection.IsEnabled = $true
            $CheckBoxCreateCollection.IsChecked = $true
        
            # If also an install program and source path, then we enable the control to create a deployment
            if (($TextBoxInstallProgram.Text.Length -gt 0 -and $TextBoxSourcePath.Text.Length -gt 0)) {
                $CheckBoxCreateDeployment.IsChecked = $true
                $CheckBoxCreateDeployment.IsEnabled = $true
            }
        }
    }
}


# Function with actions when the installation program textbox is changed
function InstallProgram-Changed {
    ErrorHandler -Object $TextBoxInstallProgram -Handling 'Clear'
    # If both a source path and an installation program has been specified, then we can enable the controls for distributing content and creating deployment and the PADT-button
    if ($TextBoxSourcePath.Text.Length -gt 0 -and $TextBoxInstallProgram.Text.Length -gt 0) {
        $CheckBoxDistributeContent.IsEnabled = $true
        $CheckBoxDistributeContent.IsChecked = $true
        #$ButtonVB.IsEnabled = $true
        
        # If a collection is also specified, we enable the control to create a deployment
        if ($TextBoxCollection.Text.Length -gt 0) {
            $CheckBoxCreateDeployment.IsEnabled = $true
            $CheckBoxCreateDeployment.IsChecked = $true
        }
        else {
            $CheckBoxCreateDeployment.IsEnabled = $false
            $CheckBoxCreateDeployment.IsChecked = $false
        }
    }
    else {
        $CheckBoxDistributeContent.IsEnabled = $false
        $CheckBoxDistributeContent.IsChecked = $false
        $CheckBoxCreateDeployment.IsEnabled = $false
        $CheckBoxCreateDeployment.IsChecked = $false
        #$ButtonVB.IsEnabled = $false
    }
}


# Function with actions when the AD group textbox is changed
function ADGroup-Changed {
    ErrorHandler -Object $TextBoxADGroup -Handling 'Clear'
    if ($TextBoxADGroup.Text -eq "") {
        # If no name for the AD group has been specified, we disable the control to create a new group
        $CheckBoxCreateADGroup.IsChecked = $false
        $CheckBoxCreateADGroup.IsEnabled = $false
    }
    else {
        # If a name for the AD group has been specified, we enable the control to create a new group
        #$CheckBoxCreateADGroup.IsChecked = $true
        $CheckBoxCreateADGroup.IsEnabled = $true
    }
}

# Function with actions when the source path textbox is changed
function SourcePath-Changed {

    
    ErrorHandler -Object $TextBoxSourcePath -Handling 'Clear'
    # If both a source path and an installation program has been specified, then we can enable the controls for distributing content and creating deployment and the PADT-button
    if ($TextBoxSourcePath.Text.Length -gt 0 -and $TextBoxInstallProgram.Text.Length -gt 0) {
        If ($ConfigMgrConnect) {
            $CheckBoxDistributeContent.IsEnabled = $true
            $CheckBoxDistributeContent.IsChecked = $true
        
        }
        <# # If a collection is also specified, we enable the control to create a deployment
            if ($TextBoxCollection.Text.Length -gt 0)
            {
                $CheckBoxCreateDeployment.IsEnabled = $true
                $CheckBoxCreateDeployment.IsChecked = $true
            }

        
            else
            {
                $CheckBoxCreateDeployment.IsEnabled = $false
                $CheckBoxCreateDeployment.IsChecked = $false
            }#>
    }
    else {
        $CheckBoxDistributeContent.IsEnabled = $false
        $CheckBoxDistributeContent.IsChecked = $false
        #Disabled because of Intune
        #$CheckBoxCreateDeployment.IsEnabled = $false
        #$CheckBoxCreateDeployment.IsChecked = $false
    }
}


function Replace-InFile {
    <#
    .SYNOPSIS
        Performs a find (or replace) on a string in a text file or files.
    .EXAMPLE
        PS> Find-InTextFile -FilePath 'C:\MyFile.txt' -Find 'water' -Replace 'wine'
    
        Replaces all instances of the string 'water' into the string 'wine' in
        'C:\MyFile.txt'.
    .EXAMPLE
        PS> Find-InTextFile -FilePath 'C:\MyFile.txt' -Find 'water'
    
        Finds all instances of the string 'water' in the file 'C:\MyFile.txt'.
    .PARAMETER FilePath
        The file path of the text file you'd like to perform a find/replace on.
    .PARAMETER Find
        The string you'd like to replace.
    .PARAMETER Replace
        The string you'd like to replace your 'Find' string with.
    .PARAMETER NewFilePath
        If a new file with the replaced the string needs to be created instead of replacing
        the contents of the existing file use this param to create a new file.
    .PARAMETER Force
        If the NewFilePath param is used using this param will overwrite any file that
        exists in NewFilePath.
    #>
    [CmdletBinding(DefaultParameterSetName = 'NewFile')]
    [OutputType()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -Path $_ -PathType 'Leaf' })]
        [string[]]$FilePath,
        [Parameter(Mandatory = $true)]
        [string]$Find,
        [Parameter()]
        [string]$Replace,
        [Parameter(ParameterSetName = 'NewFile')]
        [ValidateScript({ Test-Path -Path ($_ | Split-Path -Parent) -PathType 'Container' })]
        [string]$NewFilePath,
        [Parameter(ParameterSetName = 'NewFile')]
        [switch]$Force
    )
    begin {
        $Find = [regex]::Escape($Find)
    }
    process {
        try {
            foreach ($File in $FilePath) {
                if ($Replace) {
                    if ($NewFilePath) {
                        if ((Test-Path -Path $NewFilePath -PathType 'Leaf') -and $Force.IsPresent) {
                            Remove-Item -Path $NewFilePath -Force
                            (Get-Content $File) -replace $Find, $Replace | Add-Content -Path $NewFilePath -Force
                        }
                        elseif ((Test-Path -Path $NewFilePath -PathType 'Leaf') -and !$Force.IsPresent) {
                            Write-Warning "The file at '$NewFilePath' already exists and the -Force param was not used"
                        }
                        else {
                            (Get-Content $File) -replace $Find, $Replace | Add-Content -Path $NewFilePath -Force
                        }
                    }
                    else {
                        (Get-Content $File) -replace $Find, $Replace | Add-Content -Path "$File.tmp" -Force
                        Remove-Item -Path $File
                        Move-Item -Path "$File.tmp" -Destination $File
                    }
                }
                else {
                    Select-String -Path $File -Pattern $Find
                }
            }
        }
        catch {
            Write-Error $_.Exception.Message
        }
    }
}

function CreateCMApplication {

    param (
        [Parameter(Mandatory = $true)]
        $ApplicationName,
        [Parameter(Mandatory = $true)]
        [string]$Publisher,
        [Parameter(Mandatory = $true)]
        [string]$Appfullname,
        [Parameter()]
        [string]$ApplicationVersion,
        [Parameter()]
        [string]$ApplicationFolderName,
        [Parameter()]
        [string]$Description
            
    )
    # Create the application

    if (!(Get-CMApplication -Name $ApplicationName)) {
        $params = @{
            Name             = $ApplicationName
            Publisher        = $Publisher
            AutoInstall      = $true
            SoftwareVersion  = $ApplicationVersion
            LocalizedName    = $Appfullname
        }

        if ($Description) {
            $params.Description = $Description
        }

        New-CMApplication @params
        <# if ($Securityscope -ne "") {try{
                                            Add-CMObjectSecurityScope -Name $Securityscope -InputObject (Get-CMApplication -Name $ApplicationName)
                                            Write-Host "Added Securityscope $Securityscope to Application " -NoNewline; Write-Host $ApplicationName -ForegroundColor Green
		                                    write-log -message "Added Securityscope $Securityscope to Application $ApplicationName" -type 1
                                            } catch {
                                            Write-Host "Failed to Add Securityscope $Securityscope to Application " -NoNewline; Write-Host $ApplicationName -ForegroundColor red
		                                    write-log -message "Failed to add Securityscope $Securityscope to Application $ApplicationName" -type 3 }
                                            }
            #>
        $Newapplication = 1
        $message = "Created application $ApplicationName"
        Write-Host "Created application " -NoNewline; Write-Host $ApplicationName -ForegroundColor Green
        write-log -message $message -type 1
        Write-ProgressBar `
            -ProgressBar $ProgressBar `
            -Activity "Create Application" `
            -PercentComplete "20" `
            -CurrentOperation $message
    }
    else {
        $message = "Application $ApplicationName already exist: Take the existing one"	    
        Write-Host $message -ForegroundColor yellow
        write-log -message $message -type 1
        Write-ProgressBar `
            -ProgressBar $ProgressBar `
            -Activity "Create Application" `
            -PercentComplete "20" `
            -CurrentOperation $message
    }        
		
    # Check if a name for the application folder has been specified in the script parameters, if not a folder will not be created
    if ($ApplicationFolderName -eq "") {
        $FolderName = ""
    }
    else {
        $FolderName = $ApplicationFolderName
    }
		    
		
    # Check if an application folder should be created and if it already exists, if not create it
    if ($FolderName -ne "") {
        $ApplicationFolderPath = $SiteCode + ":" + "\Application\$FolderName"
        if (-not (Test-Path $ApplicationFolderPath)) {
            New-Item $ApplicationFolderPath
            $message = "Created application $Foldername"
            Write-Host $message -ForegroundColor Green
            write-log -message $message -type 1
            Write-ProgressBar `
                -ProgressBar $ProgressBar `
                -Activity "Create Folder" `
                -PercentComplete "25" `
                -CurrentOperation $message
        }
        else {
            $message = "Applicationfolder $Foldername already exists, no further action."
            write-log -message $message -type 1
            Write-ProgressBar `
                -ProgressBar $ProgressBar `
                -Activity "Create Folder" `
                -PercentComplete "25" `
                -CurrentOperation $message
        }
        # Move the application to folder
        $ApplicationObject = Get-CMApplication -Name $ApplicationName
        Move-CMObject -FolderPath $ApplicationFolderPath -InputObject $ApplicationObject

        $message = "Moved application $ApplicationName to Folder $Foldername"
        Write-Host $message -ForegroundColor Green
        write-log -message $message -type 1
        Write-ProgressBar `
            -ProgressBar $ProgressBar `
            -Activity "Move Application" `
            -PercentComplete "30" `
            -CurrentOperation $message
    }
		
}






function CreateCMCollection {

    param (
        [Parameter(Mandatory = $true)]
        [string]$CollectionName,
        [Parameter(Mandatory = $true)]
        [string]$CollectionNameUninstall,
        [Parameter(Mandatory = $true)]
        [string]$CollectionType,
        [string]$CollectionFolderName,
        [string]$CollectionUninstallFolderName,
        [string]$ApplicationtestCollectionname,
        [switch]$ADGroup
            
    )
    
    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Start to create Collection..."

    if (!(Get-CMCollection -Name $CollectionName)) {

        # Create the user/device-collection folder, if one has been specified in the parameters and if it does not exist
        # Set path to OU and collection folder depending on selected target
        if ($CollectionType -eq "User") {
            $CollectionFolderPath = $SiteCode + ":" + "\UserCollection\$CollectionFolderName"
            $CollectionUninstallFolderPath = $SiteCode + ":" + "\UserCollection\$CollectionUninstallFolderName"
        }
        else {
        
            $CollectionFolderPath = $SiteCode + ":" + "\DeviceCollection\$CollectionFolderName"
            $CollectionUninstallFolderPath = $SiteCode + ":" + "\DeviceCollection\$CollectionUninstallFolderName"
        }
        if ($CollectionFolderName -ne "") {
            if (-not (Test-Path $CollectionFolderPath)) {
                New-Item $CollectionFolderPath
                $message = "Created collection folder $CollectionFolderName"

                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message

                Write-Host $message -ForegroundColor Green
                write-log $message -type 1
            }
            else {
                $message = "Collection folder $CollectionFolderName already exists, no further action"

                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message

                write-log $message -type 1
            }
            if (-not (Test-Path $CollectionUninstallFolderPath)) {
                New-Item $CollectionUninstallFolderPath
                $message = "Created collection folder $CollectionUninstallFolderName"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Green
                write-log $message -type 1

            }
            else {
                $message = "Collection folder $CollectionUninstallFolderName already exists, no further action"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                write-log $message -type 1
            }
        }


		
        #Create the collection if check box is selected, and move it a collection folder if one is specified in the parameters
	
        $Schedule = New-CMSchedule -Start(Random-StartTime) –RecurInterval Days –RecurCount 3
        if ($CollectionType -eq "Device") {
				
            if (!(Get-CMDeviceCollection -Name $CollectionName)) {
                $AppCollection = New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $DeviceLimitingCollection -RefreshType Periodic -RefreshSchedule $Schedule
                $message = "Created device collection $CollectionName"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Green
                write-log $message -type 1
            }
            else {
                $AppCollection = Get-CMDeviceCollection -Name $CollectionName
                $message = "Device Collection $CollectionName already exists: Take existing one"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Yellow
                write-log $message -type 1
            }
   
            if (!(Get-CMDeviceCollection -Name $CollectionNameUninstall)) {
                $AppUninstallCollection = New-CMDeviceCollection -Name $CollectionNameUninstall -LimitingCollectionName $DeviceLimitingCollection -RefreshType Periodic -RefreshSchedule $Schedule
                $message = "Created device Uninstall collection $CollectionNameUninstall"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Green
                write-log $message -type 1
            }
            else {
                $AppUninstallCollection = Get-CMDeviceCollection -Name $CollectionNameUninstall
                $message = "Uninstall Collection $CollectionNameUninstall already exists: Take existing one"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Yellow
                write-log $message -type 1
            }

            #Create Test Collection if not exists
            if ($ApplicationtestCollectionname.Length -eq 0) {
                $message = "Evaluate Test-Collection $ApplicationtestCollectionname"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message

                if (!(Get-CMDeviceCollection -Name $ApplicationtestCollectionname)) {
                    $ApplicationtestCollection = New-CMDeviceCollection -Name $ApplicationtestCollectionname -LimitingCollectionName $DeviceLimitingCollection -RefreshType Periodic -RefreshSchedule $Schedule
                }
            }

			
            # If an AD group was specified, add a query membership rule based on that group
            if ($ADGroup) {
         
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Create AD-Group is selected, add Query Rule..."
                Add-CMDeviceCollectionQueryMembershipRule -Collection $AppCollection -QueryExpression "select *  from  SMS_R_System where SMS_R_System.SystemGroupName = ""$DomainNetbiosName\\$ADGroupName""" -RuleName "Members of AD group $ADGroupName"
                Add-CMDeviceCollectionQueryMembershipRule -Collection $AppUninstallCollection -QueryExpression "select *  from  SMS_R_System where SMS_R_System.SystemGroupName = ""$DomainNetbiosName\\$ADGroupUninstallName""" -RuleName "Members of AD group $ADGroupUninstallName"
            }
			
        }
	
        if ($CollectionType -eq "User") {
            $AppCollection = New-CMUserCollection -Name $CollectionName -LimitingCollectionName $UserLimitingCollection -RefreshType Both -RefreshSchedule $Schedule
            $message = "Created user collection $CollectionName"

            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor Green
            write-log $message -type 1
		
            $AppUninstallCollection = New-CMUserCollection -Name $CollectionNameUninstall -LimitingCollectionName $UserLimitingCollection -RefreshType Both -RefreshSchedule $Schedule
            $message = "Created user Uninstall collection $CollectionNameUninstall"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor Green
            write-log $message -type 1				
                
            if ($ADGroup) {
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Create AD-Group is selected, add Query Rule..."
                Add-CMUserCollectionQueryMembershipRule -Collection $AppCollection -QueryExpression "select * from SMS_R_User where SMS_R_User.SecurityGroupName = ""$DomainNetbiosName\\$ADGroupName""" -RuleName "Members of AD group $ADGroupName"
                Add-CMUserCollectionQueryMembershipRule -Collection $AppUninstallCollection -QueryExpression "select * from SMS_R_User where SMS_R_User.SecurityGroupName = ""$DomainNetbiosName\\$ADGroupUninstallName""" -RuleName "Members of AD group $ADGroupName"
            }

            
        }

        # Check if a collection folder name has been specified, then move the collection there
        if ($CollectionFolderName -ne "") {
            Move-CMObject -FolderPath $CollectionFolderPath -InputObject $AppCollection
            $message = "Moved collection $CollectionName to folder $CollectionFolderName"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor Green
            write-log $message -type 1
			
            Move-CMObject -FolderPath $CollectionUninstallFolderPath -InputObject $AppUninstallCollection
            $message = "Moved collection $CollectionNameUninstall to folder $CollectionUninstallFolderName"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message					
            Write-Host $message -ForegroundColor Green
            write-log $message -type 1
        }
    }
    else {
        $AppCollection = Get-CMCollection -Name $CollectionName
    }		            
}

function CreateCMDeploymentType {

    param (
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName,
        [Parameter(Mandatory = $true)]
        [string]$InstallationProgram,
        [string]$UnInstallationProgram,
        [string]$ContentSourcePath,
        [string]$DetectionType,
        [string]$DetectionMethod,
        [string]$ProductCode,
        [string]$ProductVersion,
        [switch]$Userinteraction,
        [switch]$RunInstallAs32Bit,
        [switch]$AllowFallbackSourceLocation,
        [switch]$DownloadOnSlowNetwork
            
    )


    # CREATE Script DEPLOYMENT TYPE
    
    # Create the deployment type
    $message = "Create Deployment Type..."
    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message

    if ($DetectionType -eq 'Script') {
        
        try { 
            if ($Userinteraction -eq $true) {         
                Add-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $ContentSourcePath -InstallCommand $InstallationProgram -ScriptLanguage PowerShell -ScriptText $DetectionMethod -InstallationBehaviorType InstallForSystem -RequireUserInteraction -UserInteractionMode Normal -LogonRequirementType WhetherOrNotUserLoggedOn
            }
            else {
                Add-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $ContentSourcePath -InstallCommand $InstallationProgram -ScriptLanguage PowerShell -ScriptText $DetectionMethod -InstallationBehaviorType InstallForSystem -UserInteractionMode Normal -LogonRequirementType WhetherOrNotUserLoggedOn
            }
                        
            $message = "Created a manual deployment type with a dummy detection method for $ApplicationName"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor yellow
            write-log $message -type 1
        }
        catch {
            $message = "Error while creating a manual deployment type with a dummy detection method"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message -Severity 3
            Write-Host $message -ForegroundColor Red
            Write-Host "$message for " -NoNewline; Write-Host $ApplicationName -ForegroundColor red
            Write-Host "Error: " $_.exception.message
            $logmessage = "$message for " + $ApplicationName + ". Error: " + $_.exception.message
            write-log $logmessage -type 3

            ### INSERT SCRIPT STOP ###
        } 
    }           
        
    if ($DetectionType -eq 'MSI') {
        $detection = New-CMDetectionClauseWindowsInstaller -ProductCode ($ProductCode) -Value:$true -ExpectedValue $ProductVersion -PropertyType ProductVersion -ExpressionOperator GreaterEquals
        #to get the correct format...
        $prodcode = "{" + $ProductCode + "}"
        $detection.Setting.ProductCode = $prodcode.ToUpper()
        try { 
            #Add-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $ContentSourcePath -InstallCommand $InstallationProgram -ScriptLanguage PowerShell -ScriptText 'if (Test-Path C:\DummyDetectionMethod) {Write-Host "IMPORTANT! This detection method does not work. You must manually change it."}' -InstallationBehaviorType InstallForSystem -UserInteractionMode Normal -LogonRequirementType WhereOrNotUserLoggedOn -AddDetectionClause $detection
            # Anpassung 20180515 Derploymenttype der auch mit V1802 kompatibel ist
            #Add-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $ContentSourcePath -InstallCommand $InstallationProgram -ScriptLanguage PowerShell -ScriptText 'if (Test-Path C:\DummyDetectionMethod) {Write-Host "IMPORTANT! This detection method does not work. You must manually change it."}' -InstallationBehaviorType InstallForSystem -RequireUserInteraction -UserInteractionMode Normal -LogonRequirementType WhetherOrNotUserLoggedOn -AddDetectionClause $detection
                            
            if ($Userinteraction -eq $true) {
                Add-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $ContentSourcePath -InstallCommand $InstallationProgram -InstallationBehaviorType InstallForSystem -RequireUserInteraction -UserInteractionMode Normal -LogonRequirementType WhetherOrNotUserLoggedOn -AddDetectionClause $detection
                $message = "Interaction for Application $ApplicationName was set to true"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Green
                write-log $message -type 1
            }
            else {
                Add-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName "Install $ApplicationName" -ContentLocation $ContentSourcePath -InstallCommand $InstallationProgram -InstallationBehaviorType InstallForSystem -LogonRequirementType WhetherOrNotUserLoggedOn -AddDetectionClause $detection
            }
            $message = "Created a MSI deployment type"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host "$message for " -NoNewline; Write-Host $ApplicationName -ForegroundColor yellow
            write-log "Created a MSI deployment type with Product Code {$ProductCode} and Version $ProductVersion for $ApplicationName" -type 1
        }
        catch {
 
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Error while creating a MSI deployment type with Product Code {$ProductCode} and Version $ProductVersion" -Severity 3
            Write-Host "Error while creating a MSI deployment type with Product Code {$ProductCode} and Version $ProductVersion for " -NoNewline; Write-Host $ApplicationName -ForegroundColor red
            Write-Host "Error: " $_.exception.message
            $logmessage = "Error while creating a MSI deployment type with Product Code {" + $ProductCode + "} and Version " + $ProductVersion + " for " + $ApplicationName + ". Error: " + $_.exception.message
            write-log $logmessage -type 3

            ### INSERT SCRIPT STOP ###
        }
    }
                
    Write-Host "Installation program set to: " -NoNewline; Write-Host $InstallationProgram -ForegroundColor Green
    $message = "Installation program set to $($InstallationProgram)"
    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
    write-log "$message for $ApplicationName" -type 1
    write-log "Installation program set to: $InstallationProgram" -type 1
                
    # set interaction if checked
    # excluded 2020-04-14 since CMDLet Option work again
    <#  if ($CheckboxInteraction.IsChecked -eq $true) 
            {
                try {    
                    $app = Get-CMApplication -Name $ApplicationName | Convert-CMApplication
                    $app.DeploymentTypes[0].Installer.RequiresUserInteraction = $true
                    $app1 = Convert-CMApplication -InputObject $app
                    $app1.Put()
                    Write-Host "Interaction for Application $ApplicationName is set to true sucessfully" -ForegroundColor Green
			        write-log "Interaction for Application $ApplicationName is set to true sucessfully" -type 1
     
                    } catch
                        {
                            Write-Host "Error while setting Interaction for Application " -NoNewline; Write-Host $ApplicationName -ForegroundColor red
                            Write-Host "Error: " $_.exception.message
                            $logmessage = "Error while setting Interaction for Application for "+$ApplicationName +". Error: " + $_.exception.message
                            write-log $logmessage -type 3
               
                        }
            }
    #>

    # Update the deployment type
    $NewDeploymentType = Get-CMDeploymentType -ApplicationName $ApplicationName

    # Set the uninstallation program
    if ($UnInstallationProgram.Length -gt 0) {
        Set-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -UninstallCommand $UnInstallationProgram
        $message = "Uninstallation program set to: $UnInstallationProgram"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
        Write-Host $message -ForegroundColor Green
        write-log $message -type 1
    }

    # Set behavior for running installation as 32-bit process on 64-bit systems
    if ($RunInstallAs32Bit -eq $true) {
        Set-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -Force32Bit $true
        Set-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -UninstallCommand $UninstallationProgram
        $message = "Run the installation and uninstall programs as 32-bit process on 64-bit clients is set"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
        Write-Host $message -ForegroundColor Green
        write-log $message -type 1
                
    }
    else {
        $message = "Run the installation and uninstall programs as 32-bit process on 64-bit clients is not set"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
        Write-Host $message -ForegroundColor Green
        write-log $message -type 1                
    }
            
			
    $message = "Content source path set to: $ContentSourcePath"
    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
    Write-Host $message -ForegroundColor Green
    write-log $message -type 1

    # Set the option for fallback source location
    if ($AllowFallbackSourceLocation -eq $true) {
        Set-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -ContentFallback $true
        $message = "Set Allow clients to use a fallback source location for content"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
        Write-Host $message -ForegroundColor Green
        write-log $message -type 1  
    }
    else {
        $message = "Allow clients to use a fallback source location for content is disabled"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
        Write-Host $message -ForegroundColor Green
        write-log $message -type 1  
    }

    # Set the behavior for clients on slow networks
    if ($DownloadOnSlowNetwork -eq $true) {
        Set-CMScriptDeploymentType -ApplicationName $ApplicationName -DeploymentTypeName $NewDeploymentType.LocalizedDisplayName -SlowNetworkDeploymentMode Download
        $message = "The behavior for clients on fast networks is set to: DOWNLOAD CONTENT FROM DISTRIBUTION POINT AND RUN LOCALLY"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
        Write-Host "The behavior for clients on slow networks is set to " -NoNewline; Write-Host "Download content from distribution point and run locally" -ForegroundColor Green
        write-log $message -type 1
    }
    else {
        $message = "The behavior for clients on fast networks is set to: DO NOT DOWNLOAD CONTENT"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
        Write-Host "The behavior for clients on slow networks is set to " -NoNewline; Write-Host "Do not download content" -ForegroundColor Green
        write-log $message -type 1
    }

    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Successfully created Deployment Type"

}




function DistributeContent {

    param (
        $SelectedDPGroup,
        [Parameter(Mandatory = $true)]
        [string]$ApplicationName
    )

    Set-Location $sccmloc
    # Distribute content to DP group
	        
    try {

        #Check if DPs selected from GridView
        if ($SelectedDPGroup) {
            
            ForEach ($Item in $SelectedDPGroup) {
                [string]$DPGroup = $item.name
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Distribute Content to $($DPGroup)"
                Start-CMContentDistribution -ApplicationName $ApplicationName -DistributionPointGroupName $DPGroup
                            
            }
        }
        else {
            $DPGroup = $iniFile.ConfigMgr.DPGroup
            $message = "Distribute Content to $($DPGroup)"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host "$message"
            Start-CMContentDistribution -ApplicationName $ApplicationName -DistributionPointGroupName $iniFile.ConfigMgr.DPGroup
        }
                
        write-log "$message to $DPGroup" -type 1
				
    }
    catch {
        $message = "Distributed content failed for $ApplicationName"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message -Severity 3
        Write-Host "$message" -NoNewline; Write-Host $DPGroup -ForegroundColor Red
        write-log "$message to $DPGroup" -type 3


        ### INSERT SCRIPT STOP ###
    }
}



function CreateADGroup {

    param (
        [Parameter(Mandatory = $true)]
        [string]$ADGroupName,
        [Parameter(Mandatory = $true)]
        [string]$OUPath,
        [string]$ADGroupDescription,
        [string]$ADGroupNamePrefix,
        [string]$ADUninstallGroupNamePrefix
    )

    #Create the AD group, if check box is selected
    set-location $driveloc
    try {
        try {
            Get-ADGroup $ADGroupName | Out-Null
            $message = "AD Group $ADGroupName already exist: Take the existing one"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host  $message -ForegroundColor Yellow
            write-log  $message -type 1
        }
        catch {
            $admindesc = "APP:" + $ApplicationName   # used for Webservice checking this attribute to install during TS
            New-ADGroup -Name $ADGroupName -Path $OUPath -Description $ADGroupDescription -GroupScope Global 
            get-adgroup $ADGroupName | Set-ADGroup -Add @{adminDescription = $admindesc }
			            
            $message = "Created AD group $ADGroupName in $OUPath"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor Green
            write-log $message -type 1
        }

        $ADGroupUninstallName = $ADGroupName.Replace($ADGroupNamePrefix, $ADUninstallGroupNamePrefix)
        try {        
            Get-ADGroup $ADGroupUninstallName
            $message = "AD Group $ADGroupUninstallName already exist: Take the existing one"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor Yellow
            write-log "AD Group $ADGroupUninstallName already exist: Take the existing one in $OUPath" -type 1
        }
        catch { 
            New-ADGroup -Name $ADGroupUninstallName -Path $DeviceUninstallOUPath -Description $ADUninstallGroupDescription -GroupScope Global
            $message = "Created AD Uninstall group $ADUninstallGroupName in $DeviceUninstallOUPath"
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor Green
            write-log $message -type 1
        }

    }
    catch {
        $message = "Creation of AD group $ADGroupName in $OUPath Creation of AD group failed. Error: $_.Exception.Message"
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message -Severity 3
        Write-Host $message -ForegroundColor Red
        write-log $message -type 3
    }
}


function Get-GlobalDeploymentSettings {
    return @{
        CreateInConfigMgr           = $checkboxCreateInConfigMgr.IsChecked
        CreateInIntune              = $checkboxCreateInIntune.IsChecked
        CreateCollection            = $CheckBoxCreateCollection.IsChecked
        CreateADGroup               = $CheckBoxCreateADGroup.IsChecked
        CollectionType              = if ($RadioButtonUser.IsChecked) { "User" } else { "Device" }
        CollectionFolderName        = $CollectionFolderName
        CollectionUninstallFolderName        = $CollectionUninstallFolderName
        UserInteraction             = $CheckboxInteraction.IsChecked
        DistributeContent           = $CheckBoxDistributeContent.IsChecked
        CreateDeployment            = $CheckBoxCreateDeployment.IsChecked
        ApplicationtestCollectionname = $ApplicationtestCollectionname
        RunInstallAs32Bit           = $RunInstallAs32Bit
        AllowFallbackSourceLocation = $AllowFallbackSourceLocation
        DownloadOnSlowNetwork       = $DownloadOnSlowNetwork
        EnableSWMapping             = $SWMappingenabled -eq $true
        SWProductTable              = $SWProductTable
        SWMapDBName                 = $SWMapDBName
        CatalogID                   = $TextboxSWCatalogID.Text
        Owner                       = $env:USERNAME
        LogFile                     = $logfile
        ShowSummary                 = $showsummaryfile -eq $true
        MailServer                  = $mailserver
        MailFrom                    = $mailfrom
        MailRecipients              = $mailrecipients
        AADGroupNamePrefix          = $AADGroupNamePrefix
        AADUninstallGroupNamePrefix = $AADUninstallGroupNamePrefix
        PilotAADGroup               = $PilotAADGroup
        DeviceOUPath                = $DeviceOUPath
        OUPath                      = $DeviceOUPath
        
    }
}

function Create-ApplicationObjects {
    param (
        [string]$ApplicationName,
        [string]$Appfullname,
        [string]$ApplicationVersion,
        [string]$ApplicationDescription,
        [string]$Publisher,
        [string]$InstallationProgram,
        [string]$UninstallationProgram,
        [string]$ContentSourcePath,
        [string]$CollectionName,
        [string]$CollectionNameUninstall,
        [string]$CollectionFolderName = $null,
        [string]$CollectionUninstallFolderName = $null,
        [string]$ADGroupName,
        [string]$ADGroupDescription,
        [string]$OUPath = $null,
        [string]$ProductCode,
        [string]$ProductVersion,
        [string]$DetectionType = "Script",
        [string]$DetectionMethod = 'if (Test-Path C:\\DummyDetectionMethod) {Write-Host "IMPORTANT! This detection method does not work. You must manually change it."}',

        [bool]$CreateInConfigMgr = $true,
        [bool]$CreateInIntune = $true,
        [bool]$CreateCollection = $false,
        [bool]$CreateADGroup = $false,
        [string]$CollectionType = "Device",
        [bool]$UserInteraction = $false,
        [bool]$DistributeContent = $false,
        [bool]$CreateDeployment = $false,
        [string]$DeployPurpose = "Available",
        [string]$ApplicationtestCollectionname = $null,
        [bool]$RunInstallAs32Bit = $false,
        [bool]$AllowFallbackSourceLocation = $false,
        [bool]$DownloadOnSlowNetwork = $false,
        [bool]$EnableSWMapping = $false,
        [string]$SWProductTable = $null,
        [string]$SWMapDBName = $null,
        [string]$CatalogID = $null,
        [string]$Owner = $env:USERNAME,
        [string]$LogFile = $null,
        [bool]$ShowSummary = $false,
        [string]$MailServer = $null,
        [string]$MailFrom = $null,
        [string]$MailRecipients = $null,
        [string]$AADGroupNamePrefix = $null,
        [string]$AADUninstallGroupNamePrefix = $null,
        [string]$PilotAADGroup = $null,
        [string]$DeviceOUPath = $null
    )

    # This function assumes all pre-validation is already done and starts object creation.
    $ProgressBar = New-ProgressBar
    Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create Application" -CurrentOperation "Starting Object creation..."

    if ($CreateInConfigMgr) {

        try {
            Import-Module (Join-Path $(Split-Path $env:SMS_ADMIN_UI_PATH) ConfigurationManager.psd1) | Out-Null
            Set-Location ($iniFile.ConfigMgr.SiteCode + ":")
            $sccmloc = Get-Location
            $ConfigMgrConnect = $true
        }
        catch {
            $Message = "No connection to ConfigMgr Provider established..."
            $ConfigMgrConnect = $false
            Write-Host $Message -ForegroundColor DarkYellow
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $Message -Severity 3
            return
        }

        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        # ++   Application
        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CreateCMApplication -ApplicationName $ApplicationName -Publisher $Publisher -Appfullname $Appfullname -ApplicationVersion $ApplicationVersion -ApplicationFolderName $ApplicationFolderName -Description $ApplicationDescription

        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        # ++   Collections
        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
        if ($CreateCollection) {
            CreateCMCollection -CollectionName $CollectionName -CollectionNameUninstall $CollectionNameUninstall -CollectionType $CollectionType -CollectionFolderName $CollectionFolderName -CollectionUninstallFolderName $CollectionUninstallFolderName -ADGroup:$CreateADGroup -ApplicationtestCollectionname $ApplicationtestCollectionname
        }

        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        # ++   DEPLOYMENT TYPE
        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        CreateCMDeploymentType -ApplicationName $ApplicationName -InstallationProgram $InstallationProgram -UnInstallationProgram $UninstallationProgram -ContentSourcePath $ContentSourcePath -DetectionType $DetectionType -DetectionMethod $DetectionMethod -ProductCode $ProductCode -ProductVersion $ProductVersion -Userinteraction:$UserInteraction -RunInstallAs32Bit:$RunInstallAs32Bit -AllowFallbackSourceLocation:$AllowFallbackSourceLocation -DownloadOnSlowNetwork:$DownloadOnSlowNetwork

        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        # ++   Distribute Content
        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if ($DistributeContent) {
            DistributeContent -ApplicationName $ApplicationName
        }

        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        # ++   Deployments
        # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        if ($CreateDeployment) {

            try {
                New-CMApplicationDeployment -CollectionName $CollectionName -Name $ApplicationName -DeployPurpose $DeployPurpose
                $message = "Created $DeployPurpose deployment for $ApplicationName to collection $CollectionName"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Green
                write-log $message -type 1
            } catch {
                $message = "⚠ Deployment already exists or failed for $ApplicationName to collection $CollectionName. Error: $($_.Exception.Message)"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Red
                write-log $message -type 2
            }


            try {
                New-CMApplicationDeployment -CollectionName $CollectionNameUninstall -Name $ApplicationName -DeployPurpose Required -DeployAction Uninstall
                $message = "Created UNINSTALL deployment for $ApplicationName to collection $CollectionNameUninstall"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Green
                write-log $message -type 1
            } catch {
                $message = "⚠ UNINSTALL deployment already exists or failed for $ApplicationName to collection $CollectionNameUninstall. Error: $($_.Exception.Message)"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Red
                write-log $message -type 2
            }
           

            if (-not [string]::IsNullOrWhiteSpace($ApplicationtestCollectionname)) {
                try {
                    New-CMApplicationDeployment -CollectionName $ApplicationtestCollectionname -Name $ApplicationName -DeployPurpose Available -DeployAction Install
                    $message = "Created Test deployment for $ApplicationName to collection $ApplicationtestCollectionname"
                    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                    Write-Host $message -ForegroundColor Green
                    write-log $message -type 1
                } catch {
                    $message = "⚠ Test deployment already exists or failed for $ApplicationName to collection $ApplicationtestCollectionname. Error: $($_.Exception.Message)"
                    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                    Write-Host $message -ForegroundColor Red
                    write-log $message -type 2
                }
            }
        }


        if (-not $ProductCode) {
            $message = "IMPORTANT! Remember to manually modify the detection method afterwards."
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
            Write-Host $message -ForegroundColor Yellow
        }

        Set-Location $driveloc
    } # End ConfigMgr


    # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    # ++   AD-Groups
    # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

    if ($CreateADGroup -and $OUPath) {
        CreateADGroup -ADGroupName $ADGroupName -OUPath $OUPath -ADGroupDescription $ADGroupDescription -ADGroupNamePrefix $AADGroupNamePrefix -ADUninstallGroupNamePrefix $AADUninstallGroupNamePrefix
    }


    # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    # ++   Intune
    # ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        
    # Create Application in Intune
    
    if ($CreateInIntune) {
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Start creating Intune Application"

        if (-not (Connect-Azure -User $AADUser)) {
            $Message = "Failed to connect to Azure or Intune."
            Write-Host $Message -ForegroundColor Red
            Start-ProgressBar -ProgressBar $ProgressBar -Activity "Failed to connect to Intune" -PercentComplete "0" -CurrentOperation $Message -Severity 3
            return
        }


        $IntuneApp = Create-IntuneApp `
            -Publisher $Publisher `
            -DisplayName $ApplicationName `
            -ContentSourcePath $ContentSourcePath `
            -InstallationProgram $InstallationProgram `
            -UninstallationProgram $UninstallationProgram `
            -Description $ApplicationDescription `
            -DetectionType $DetectionType `
            -DetectionMethod $DetectionMethod
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "$ApplicationName successfully created"

        if ($CreateDeployment) {
            $message = "Create AAD Group $($ApplicationName)-Install"
            Write-Host $message
            Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message

            $AADInstalllAppGroup = Create-AADGroup -DisplayName "$($AADGroupNamePrefix)$($ApplicationName)-Install" -Description "AutoGenerated Group by <Cancom Application Manager> for $($ApplicationName) deployment"
            [string]$AppID = $IntuneApp.id

            try {
                Add-IntuneWin32AppAssignmentGroup -Include -ID $AppID -GroupID $($AADInstalllAppGroup.Id) -Intent $DeployPurpose -Notification "showAll"
                $message = "Successfully created App assignment"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                Write-Host $message -ForegroundColor Green
            }
            catch {
                $message = "Error: Failed to create App assignment"
                Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message -Severity 3
                Write-Host $message -ForegroundColor Red
            }

            if ($UninstallationProgram) {
                $AADUninstalllAppGroup = Create-AADGroup -DisplayName "$($AADUninstallGroupNamePrefix)$($ApplicationName)-Uninstall" -Description "AutoGenerated Group by <Cancom Application Manager> for $($ApplicationName) Uninstall deployment"
                try {
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $AppID -GroupID $($AADUninstalllAppGroup.Id) -Intent "uninstall" -Notification "showAll"
                    $message = "Successfully created App-Uninstall assignment"
                    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                    Write-Host $message -ForegroundColor Green
                }
                catch {
                    $message = "Error: Failed to create App Uninstall assignment"
                    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message -Severity 3
                    Write-Host $message -ForegroundColor Red
                }
            }

            if ($PilotAADGroup) {
                $AADPilotAppGroup = Create-AADGroup -DisplayName "$($AADGroupNamePrefix)$PilotAADGroup" -Description "AutoGenerated Group by <Cancom Application Manager> for Application Deployment Tests"
                try {
                    Add-IntuneWin32AppAssignmentGroup -Include -ID $AppID -GroupID $($AADPilotAppGroup.Id) -Intent "available" -Notification "showAll"
                    $message = "Successfully created App assignment for TestGroup $($PilotAADGroup)"
                    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message
                    Write-Host $message -ForegroundColor Green
                }
                catch {
                    $message = "Error: Failed to create Test-Group assignment"
                    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $message -Severity 3
                    Write-Host $message -ForegroundColor Red
                }
            }
        }
    } # End Intune
    #END Create Intune / AAD Objects if Checkbox is selected

 
    if ($EnableSWMapping -and $SWProductTable -and $SWMapDBName -and $ADGroupName) {
        $query = "select * from $SWProductTable where ADGruppe = '$ADGroupName'"
        $ADgrs = query-SQLSWMAP -querytext $query -DBname $SWMapDBName

        if (-not $ADgrs) {
            try {
                Insert-SQLSWProduct -SWPName $ApplicationName -SWPVersion $ApplicationVersion -SWInstallmethod "SCCM" -SWPOwner $Owner -SWPADGroup $ADGroupName -CatalogID $CatalogID
                Write-Host "Successfully created entry in SoftwareProductCatalog-Table" -ForegroundColor Green
                write-log "Successfully created entry in SoftwareProductCatalog-Table" -type 1
            }
            catch {
                $message = "Error while creating Entry in SoftwareProductCatalog-Table: $($_.Exception.Message)"
                Write-Host $message -ForegroundColor Red
                write-log $message -type 3
            }
        }
        else {
            $message = "AdGroup already in SoftwareProductCatalog-Table, skipping insert"
            Write-Host $message -ForegroundColor Yellow
            write-log $message -type 1
        }
        $ProgressBar1.PerformStep()
    }

    if ($MailServer) {
        $subjecttext = "Paket $ApplicationName wurde importiert"
        $bodytext = "<b>Paket:</b> $ApplicationName<br><b>User:</b> $Owner"
        Send-MailMessage -SmtpServer $MailServer -From $MailFrom -Subject $subjecttext -Body $bodytext -BodyAsHtml -To $MailRecipients
    }


    if ($ShowSummary -and (Test-Path $LogFile)) {
        Invoke-Item $LogFile
    }

    Show-Toast -Message "Application '$ApplicationName' created successfully!"

    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "SUCCESS: Created Application!"
    Close-ProgressBar $ProgressBar
}




#############################################################################
# END Functions
#############################################################################


#Get Applicationfolders
$DDPackageArray = Start-Appimportform

[array]$DDPackageArray = foreach ($Packagefolder in $Packagefolders) { $Packagefolder.name } 

foreach ($Packagefolder in $Packagefolders) {
    $Packagefolder.name
} 


if ($Showonlynewpackages -eq $true) {
    
    ForEach ($Item in $DDPackageArray) {
                                  
        if ((($Item.ToCharArray() | where-object { $_ -eq $iniFile.Global.PackageDelimiter } | Measure-Object).Count) -ge $Delimitercount) {

            if ($item.Split($iniFile.Global.PackageDelimiter)[0]) {
                                        

                # $appname= $item.Split($iniFile.Global.PackageDelimiter)[0]+"-"+$item.Split($iniFile.Global.PackageDelimiter)[1]+" "+$item.Split($iniFile.Global.PackageDelimiter)[2]
                $appADGroup = $ADGroupNamePrefix + $item.Split($iniFile.Global.PackageDelimiter)[0] + "-" + $item.Split($iniFile.Global.PackageDelimiter)[1] + "_" + $item.Split($iniFile.Global.PackageDelimiter)[2]
                                                                                
                                                                                           
                                                                           
                $ErrorActionPreference = "SilentlyContinue"
                try { 
                    $filt = "*" + $Testpackagestring + "*"
                    if (!($appADGroup -like $filt)) {
                        #Get-ADGroup $appADGroup -filter {name -like $appADGroup} -SearchBase $DeviceOUPath -SearchScope OneLevel
                        get-adgroup $appADGroup    | out-null 
                        write-host "Ignoring: " $appADGroup -ForegroundColor Red
                    }
                    else {
                        write-host "Checking Testgroup: " $appADGroup -ForegroundColor Yellow
                        $DDPackages.Items.Add($Item) | out-null                                                                                            
                    }
                }
                catch {
                    write-host "Checking: " $appADGroup -ForegroundColor Green
                    $DDPackages.Items.Add($Item) | out-null                                                                                            
                }
                                                                  
            }
        }
                                           
    }
}
else {
    ForEach ($Item in $DDPackageArray) {
        if ((($Item.ToCharArray() | where-object { $_ -eq $iniFile.Global.PackageDelimiter } | Measure-Object).Count) -ge $Delimitercount) {
            $DDPackages.Items.Add($Item) | out-null
        }
    }
}
        
#Set-Location $CurrentLocation
#$DDPackages.sorted = $true;
#$FormAppImport.Controls.Add($DDPackages)


#Validate Config on change
$TextBoxIntuneOutputFolder.add_TextChanged(
    {
        if ($TextBoxIntuneOutputFolder.Text -match "^([a-zA-Z]+:)?(\\[a-zA-Z0-9_.-: :]+)*\\?$") {
            ErrorHandler -Object $TextBoxIntuneOutputFolder -Handling 'Clear'
        }
        else {
            ErrorHandler -Object $TextBoxIntuneOutputFolder -Handling 'Set'
        }
    })

$DDPackages.add_SelectionChanged({ ButtonLoadPackageInfo })

#If Form edited execute Function
$TextBoxAppName.add_TextChanged({ AppName-Changed })
$TextBoxPublisher.add_TextChanged({ AppName-Changed })
$TextBoxVersion.add_TextChanged({ AppName-Changed })
#No Evaluation if ConfigMgr connection not exist
If ($ConfigMgrConnect) {
    $TextBoxCollection.add_TextChanged({ CollectionName-Changed })
    $TextBoxInstallProgram.add_TextChanged({ InstallProgram-Changed })
}
$TextBoxSourcePath.add_TextChanged({ SourcePath-Changed })
$TextBoxADGroup.add_TextChanged({ ADGroup-Changed })

#Add Function to App Browse Button
$TextBoxDDPackages.Visibility = "Hidden"
$ButtonBrowseApp.Add_Click({
        $SelectedApp = OpenFolderDialog $Packagefolderpath
        write-host $SelectedApp
        if ($SelectedApp) {
            $TextBoxDDPackages.Text = $SelectedApp
            $TextBoxDDPackages.Visibility = "Visible"
            $DDPackages.Visibility = "Hidden"
            $AppFolderName = $SelectedApp | split-path -leaf
            ButtonLoadPackageInfo -Packagename $AppFolderName -Path $SelectedApp
        }
    })

#Add Function to MSI Forms
$ButtonMSIClear.Add_Click({ ButtonMSIClearClick })
$TextBoxMSIPackage.add_TextChanged({ MSIFile-Changed })
#$TextBoxAppName.Add_Leave({ write-host "Left the first textbox" })

#Add Function to MSI Browse Button
$ButtonMSI.Add_Click({ ButtonMSIClick $Packagefolderpath })


#Add Function to Distribute Checkboxes
$CheckBoxDistributeContent.Add_Checked({ CheckBoxDistributeContentClick })
$LabelSelectDP.Add_PreviewMouseDown({ LabelSelectDistributionPointClick })
$checkboxCreateInIntune.Add_Checked({ $ButtonCreate.IsEnabled = $True })
$checkboxCreateInIntune.Add_UnChecked({
        if (!$checkboxCreateInIntune.IsChecked -and !$checkboxCreateInConfigMgr.IsChecked) {
            $ButtonCreate.IsEnabled = $False
        }
        else {
            $ButtonCreate.IsEnabled = $True
        }
    })
$checkboxCreateInConfigMgr.Add_Checked({ $ButtonCreate.IsEnabled = $True })
$checkboxCreateInConfigMgr.Add_UnChecked({
        if (!$checkboxCreateInIntune.IsChecked -and !$checkboxCreateInConfigMgr.IsChecked) {
            $ButtonCreate.IsEnabled = $False
        }
        else {
            $ButtonCreate.IsEnabled = $True
        }
    })

#Set Checkbox Status
if ($AllowInteractionDefault -eq $True) {
    $CheckboxInteraction.IsChecked = $true
}
if ($CreateADGroup -eq $True) {
    $CheckBoxCreateADGroup.IsChecked = $true
}

if (!$checkboxCreateInIntune.IsChecked -and !$checkboxCreateInConfigMgr.IsChecked) {
    $ButtonCreate.IsEnabled = $False
}


$TabControlMain.Add_SelectionChanged({
    $selectedTab = $TabControlMain.SelectedItem

    switch ($selectedTab.Name) {
        "TabCreateApp" {
            $BoxCommonControls.Visibility = "Visible"
            $ButtonCreate.Visibility = "Visible"
            $ButtonCreateWinGet.Visibility = "Collapsed"
        }
        "TabCreateWinget" {
            $BoxCommonControls.Visibility = "Visible"
            $ButtonCreate.Visibility = "Collapsed"
            $ButtonCreateWinGet.Visibility = "Visible"
        }
        "TabConfig" {
            $BoxCommonControls.Visibility = "Collapsed"
            $ButtonCreate.Visibility = "Collapsed"
            $ButtonCreateWinGet.Visibility = "Collapsed"
        }
        default {
            $BoxCommonControls.Visibility = "Visible"
            $ButtonCreate.Visibility = "Visible"
            $ButtonCreateWinGet.Visibility = "Collapsed"
        }
    }
})




#############################################################################
# BEGIN Winget
#############################################################################

#Check Winget availability
<#
if(Test-Administrator){
    $AppInstaller = Get-AppxProvisionedPackage -Online | Where-Object DisplayName -eq Microsoft.DesktopAppInstaller

    $Winget = Get-ChildItem -Path "C:\Program Files\WindowsApps\" -Recurse -Filter "winget.exe"

    If($AppInstaller.Version -lt "2022.506.16.0" -or $Winget.VersionInfo.FileVersion -lt "1.4.3531") {

        Write-Host "Winget is not installed, trying to install latest version from Github" -ForegroundColor Yellow

        Try {
            
            Write-Host "Creating Winget Packages Folder" -ForegroundColor Yellow

            if (!(Test-Path -Path C:\ProgramData\WinGetPackages)) {
                New-Item -Path C:\ProgramData\WinGetPackages -Force -ItemType Directory
            }

            Set-Location C:\ProgramData\WinGetPackages

            #Downloading Packagefiles
            #Microsoft.UI.Xaml.2.7.0
            Invoke-WebRequest -Uri "https://www.nuget.org/api/v2/package/Microsoft.UI.Xaml/2.7.0" -OutFile "C:\ProgramData\WinGetPackages\microsoft.ui.xaml.2.7.0.zip"
            Expand-Archive C:\ProgramData\WinGetPackages\microsoft.ui.xaml.2.7.0.zip -Force
            #Microsoft.VCLibs.140.00.UWPDesktop
            Invoke-WebRequest -Uri "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx" -OutFile "C:\ProgramData\WinGetPackages\Microsoft.VCLibs.x64.14.00.Desktop.appx"
            #Winget
            Invoke-WebRequest -Uri "https://github.com/microsoft/winget-cli/releases/latest/download/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -OutFile "C:\ProgramData\WinGetPackages\Winget.msixbundle"
            #Installing dependencies + Winget
            Add-ProvisionedAppxPackage -online -PackagePath:.\Winget.msixbundle -DependencyPackagePath .\Microsoft.VCLibs.x64.14.00.Desktop.appx,.\microsoft.ui.xaml.2.7.0\tools\AppX\x64\Release\Microsoft.UI.Xaml.2.7.Appx -SkipLicense

            Write-Host "Starting sleep for Winget to initiate" -Foregroundcolor Yellow
            Start-Sleep 2
        }
        Catch {
            Throw "Failed to install Winget"
            Break
        }

        }
    Else {
        Write-Host "Winget already installed, moving on" -ForegroundColor Green
    }
} Else {
     Write-Host "Running without privileged permissions, skip WinGet availibility"
}

#>

#Get WinGet cmd
Get-WingetCmd










$ButtonWingetSearch.add_Click({
    $SearchTerm = $TextBoxWingetSearch.Text
    $Message = "Searching for $SearchTerm"


    $ProgressBar = New-ProgressBar
    Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation $Message

    Write-Host $Message
    $arrayItems = @()

    # Use PowerShell module if available, else fallback to native
    if (Get-Command -Name "Find-WinGetPackage" -ErrorAction SilentlyContinue) {
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Searching via WinGet PowerShell Module..."
        $List = Get-WingetAppInfo $SearchTerm
    } else {
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Searching via Native winget.exe..."
        $List = Get-WingetAppInfo-Native $SearchTerm
    }

    foreach ($App in $List) {
        Start-ProgressBar -ProgressBar $ProgressBar -CurrentOperation "Processing $($App.Name)..."

        $itemObject = New-Object PSObject

        $itemObject | Add-Member -Type NoteProperty -Name "ID"      -Value $App.Id
        $itemObject | Add-Member -Type NoteProperty -Name "Name"    -Value $App.Name

        $version = if ($App.Version) { $App.Version } else { "" }
        $itemObject | Add-Member -Type NoteProperty -Name "Version" -Value $version

        $source = if ($App.Source) { $App.Source } else { "" }
        $itemObject | Add-Member -Type NoteProperty -Name "Source"  -Value $source

        $arrayItems += $itemObject
    }

    $dataGridWinget.ItemsSource = $arrayItems
    Close-ProgressBar $ProgressBar
})




#Define WingetFunction for use in PSADT
$WingetFunction = '
function Get-WingetCmd {
            #WinGet Path (if User/Admin context)
            $UserWingetPath = Get-Command winget.exe -ErrorAction SilentlyContinue
            #WinGet Path (if system context)
            $SystemWingetPath = Resolve-Path "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe"

            #Get Winget Location in User/Admin context
            if ($UserWingetPath) {
                $Script:Winget = $UserWingetPath.Source
                Write-Log  -Message "WinGet path defined $Winget" -Severity 1 -Source $deployAppScriptFriendlyName
            }
            #Get Winget Location in System context
            elseif ($SystemWingetPath) {
                #If multiple version, pick last one
                $Script:Winget = $SystemWingetPath[-1].Path
                Write-Log  -Message "WinGet path defined $Winget" -Severity 1 -Source $deployAppScriptFriendlyName
            }
            else {
                Write-Log -Message "WinGet is not installed, mandatory to run WinGet integration" -Severity 3 -Source $deployAppScriptFriendlyName
                
            }
	}

    Get-WingetCmd
	#Install <AppName> by Winget
'
$WingetDefaultCMD = 'Execute-Process -Path "$winget" -Parameters "install --id <AppID> -h --accept-package-agreements --accept-source-agreements"'


    $dataGridWinget.add_SelectionChanged({
        foreach ($WingetApp in $dataGridWinget.SelectedItems) {
            $WingetCMD = $WingetDefaultCMD.replace('<AppID>', $($WingetApp.ID)).replace('<AppName>', $($WingetApp.Name))
            $TextBoxWingetPreview.Text = $WingetCMD
        }

    })


$ButtonCreateWinGet.add_Click({

    $ProgressBar = New-ProgressBar
    Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create WinGET App" -CurrentOperation "Integrate Winget Application into PSADT Template"

    # Check if PSADT Template path is valid
    if ([string]::IsNullOrWhiteSpace($PSADTTemplate) -or -not (Test-Path $PSADTTemplate)) {

        $result = [System.Windows.MessageBox]::Show("PSADT Template path is not valid. Do you want to download the v4 template from GitHub?", "PSADT Template Missing", "YesNo", "Warning")

        if ($result -eq "Yes") {
            $templateFolder = Download-PSADTTemplate -DestinationFolder $Packagefolderpath -ProgressBar $ProgressBar


            if ($templateFolder -ne $null) {
                $global:PSADTTemplate = $templateFolder.FullName
            }
            else {
                [System.Windows.MessageBox]::Show("Template could not be downloaded. Aborting.", "Error", "OK", "Error")
                return
            }
        }
        else {
            [System.Windows.MessageBox]::Show("You chose not to download the template. Aborting.", "Cancelled", "OK", "Warning")
            return
        }
    }

    foreach ($WingetApp in $dataGridWinget.SelectedItems) {
        $message = "Start creating PSADT App for $($WingetApp.Name)"
        Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create WinGET App" -CurrentOperation $message -Severity 1
        Write-Host $message

        $yaml = Get-WingetInstallContent -PackageId $WingetApp.ID -PackageVersion $WingetApp.Version
        $localeYaml = Get-WingetLocaleContent -PackageId $WingetApp.ID -PackageVersion $WingetApp.Version

        $publisher = if ($localeYaml) { $localeYaml.Publisher } else { "Unknown Publisher" }
        $shortDesc = if ($localeYaml) { $localeYaml.ShortDescription } else { "" }

        $description = "$shortDesc | WinGetID=$($WingetApp.ID) | Update=True | Created by ApplicationManager from WinGet"


        Write-Host $publisher


        if ($yaml -eq $null) {
            $message = "❌ YAML not found for $($WingetApp.ID)"
            Write-Host $message
            Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create WinGET App" -CurrentOperation $message -Severity 3
            continue
        }

        # Build folder path using version
        $WingetAppTMPPath = Join-Path $Packagefolderpath "$($WingetApp.Name)_$($WingetApp.ID)"
        Write-Host "Source: $PSADTTemplate"
        Write-Host "Destination: $WingetAppTMPPath"
        try {

            $itemsToCopy = Get-ChildItem -Path $PSADTTemplate -Force

            foreach ($item in $itemsToCopy) {
                $destination = Join-Path $WingetAppTMPPath $item.Name

                if (Test-Path $destination) {
                    Write-Host "Skipping existing: $destination"
                } else {
     
                    try {
                        Copy-Item -Path $item.FullName -Destination $destination -Recurse -Force
                    }
                    catch {
                        Write-Warning "❌ Failed to copy: $($item.FullName) → $destination"
                        Write-Warning $_
                    }
                }
            }

        }
        catch {
            $message = "❌ Failed to copy template to $WingetAppTMPPath"
            Write-Host $message
            Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create WinGET App" -CurrentOperation $message -Severity 3
            continue
        }

        # Determine which file to patch (v3 or v4)
        $fileV3 = Join-Path $WingetAppTMPPath "Deploy-Application.ps1"
        $fileV4 = Join-Path $WingetAppTMPPath "Invoke-AppDeployToolkit.ps1"
        $file = if (Test-Path $fileV4) { $fileV4 } elseif (Test-Path $fileV3) { $fileV3 } else { $null }

        if (-not $file) {
            $message = "❌ No PSADT script found in: $WingetAppTMPPath"
            Write-Host $message
            Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create WinGET App" -CurrentOperation $message -Severity 3
            continue
        }

        # Set executable name for install/uninstall commands
        if ($file -eq $fileV4) {
            $installerExe = "Invoke-AppDeployToolkit.exe"
        } elseif ($file -eq $fileV3) {
            $installerExe = "Deploy-Application.exe"
        } else {
            $installerExe = "Deploy-Application.exe"  # fallback just in case
        }

        # Default installer (from PSADT)
        $InstallationProgram = "$installerExe"
        $UninstallationProgram = "$installerExe UNINSTALL"

        # Check for local _install.bat / _uninstall.bat overrides
        $installBat = Join-Path $WingetAppTMPPath "_install.bat"
        $uninstallBat = Join-Path $WingetAppTMPPath "_uninstall.bat"

        if (Test-Path $installBat) {
            $InstallationProgram = "_install.bat"
            Write-Host "Found custom install script: $InstallationProgram"
        }

        if (Test-Path $uninstallBat) {
            $UninstallationProgram = "_uninstall.bat"
            Write-Host "Found custom uninstall script: $UninstallationProgram"
        }

        # Prepare content for insertion
        $DefaultWingetCall = $WingetFunction + $TextBoxWingetPreview.Text
        $Content = $DefaultWingetCall.Replace('<AppID>', $WingetApp.ID).Replace('<AppName>', $WingetApp.Name)

        try {
            Replace-InFile -FilePath $file -Find '## <Perform Installation tasks here>' -Replace $Content
            Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create WinGET App" -CurrentOperation "✅ Successfully created '$($WingetApp.Name)' in Sourcepath"
        }
        catch {
            Start-ProgressBar -ProgressBar $ProgressBar -Activity "Create WinGET App" -CurrentOperation "❌ Failed to modify script for $($WingetApp.Name)" -Severity 3
        }

        
        if ($checkboxCreateInConfigMgr.IsChecked -or $checkboxCreateInIntune.IsChecked) {
            # Define application name and other deployment info

            Write-Host $WingetApp
            Write-Host $yaml

            if ($publisher -and $publisher -ne "Unknown Publisher") {
                $ApplicationName = "$publisher $($WingetApp.Name) $($WingetApp.Version)"
            } else {
                $ApplicationName = "$($WingetApp.Name) $($WingetApp.Version)"
            }

            $DeployPurpose = if ($RadioButtonRequired.IsChecked) { "Required" } else { "Available" }

            <#

            $DetectionScript = @'
$app = winget list '<AppID>' -e --accept-source-agreements 2>&1

if ($app -notmatch 'No installed package found matching input criteria') {
    Write-Output 'Detected'
    exit 0
} else {
    Write-Output 'Not Detected'
    exit 1
}
'@

            # Replace placeholder <AppID> with actual ID
            $DetectionScript = $DetectionScript.Replace('<AppID>', $WingetApp.ID)


            # Define support path
            $supportFilesPath = Join-Path $WingetAppTMPPath "SupportFiles"
            if (-not (Test-Path $supportFilesPath)) {
                New-Item -Path $supportFilesPath -ItemType Directory -Force | Out-Null
            }

            # Save detection method to a file
            $detectScriptPath = Join-Path $supportFilesPath "Detect-$($WingetApp.ID).ps1"
            try {
                $DetectionScript | Out-File -FilePath $detectScriptPath -Encoding UTF8 -Force
                Write-Host "✅ Detection script written to: $detectScriptPath"
            } catch {
                Write-Warning "❌ Failed to write detection script: $_"
            }

            #>


            # Define support path
            $supportFilesPath = Join-Path $WingetAppTMPPath "SupportFiles"
            if (-not (Test-Path $supportFilesPath)) {
                New-Item -Path $supportFilesPath -ItemType Directory -Force | Out-Null
            }

            # Read detection script template
            $scriptTemplatePath = Join-Path $workingdir "Files\DetectionScriptWingetTemplate.ps1"

            if (Test-Path $scriptTemplatePath) {
                try {
                    $DetectionScript = Get-Content -Path $scriptTemplatePath -Raw

                    # Replace placeholders
                    $DetectionScript = $DetectionScript.Replace('<APPID>', $WingetApp.ID)
                    $DetectionScript = $DetectionScript.Replace('<APPVERSION>', $WingetApp.Version)

                    # Save processed script
                    $detectScriptPath = Join-Path $supportFilesPath "Detect-$($WingetApp.ID).ps1"
                    $DetectionScript | Out-File -FilePath $detectScriptPath -Encoding UTF8 -Force

                    Write-Host "✅ Detection script written to: $detectScriptPath"
                } catch {
                    Write-Warning "❌ Failed to read or write detection script: $_"
                }
            } else {
                Write-Warning "❌ Detection script template not found: $scriptTemplatePath"
            }





            $appParams = @{
                ApplicationName           = $ApplicationName
                Appfullname               = $ApplicationName
                ApplicationVersion        = $WingetApp.Version
                ApplicationDescription    = $description
                Publisher                 = $publisher
                InstallationProgram       = $InstallationProgram
                UninstallationProgram     = $UninstallationProgram
                ContentSourcePath         = (Join-Path $Packagefolderpath "$($WingetApp.Name)_$($WingetApp.ID)")
                CollectionName            = $ApplicationName
                CollectionNameUninstall   = "$ApplicationName - Uninstall"
                ADGroupName               = "$($ADGroupNamePrefix)_$($ApplicationName)"
                ADGroupDescription        = "Members of this group will be targeted for deployment of $ApplicationName in ConfigMgr"
                ProductCode               = $null
                ProductVersion            = $WingetApp.Version
                DetectionType             = "Script"
                DetectionMethod           = $DetectionScript
                DeployPurpose             = $DeployPurpose
            }

            $globalParams = Get-GlobalDeploymentSettings
            Create-ApplicationObjects @globalParams @appParams
        }

        Start-Sleep -Seconds 2
    }



    Close-ProgressBar $ProgressBar
})


# fill the combobox with some powershell objects
#$null = $DDPackages.Items.Add('sdf')

#Add Function to ButtonCreate
#$ButtonCreate.add_Click({ ButtonCreateClick })

$ButtonCreate.add_Click({
    if (-not (Validate-Form)) {
        return
    }

    $ProductCode = $null
    $ProductVersion = $null

    if ($TextBoxMSIPackage.Text) {
        try {
            $ProductCode = [GUID](Get-MsiFileInformation -Path $TextBoxMSIPackage.Text -Property ProductCode)[3]
        } catch { $ProductCode = $null }

        try {
            $ProductVersion = (Get-MsiFileInformation -Path $TextBoxMSIPackage.Text -Property ProductVersion).Trim()
        } catch { $ProductVersion = $null }
    }

    $DetectionType = if ($ProductCode -and $ProductVersion) { "MSI" } else { "Script" }

    $ApplicationName = if ($TextBoxVersion.Text) { "$($TextBoxAppName.Text) $($TextBoxVersion.Text)" } else { $TextBoxAppName.Text }
    $DeployPurpose = "Available"
    if ($RadioButtonRequired.IsChecked) { $DeployPurpose = "Required" }
    elseif ($RadioButtonAvailable.IsChecked) { $DeployPurpose = "Available" }

    if ($TextBoxMSIPackage.Text) {
        try {
            $ProductCode = [GUID](Get-MsiFileInformation -Path $TextBoxMSIPackage.Text -Property ProductCode)[3]
        } catch { $ProductCode = $null }

        try {
            $ProductVersion = (Get-MsiFileInformation -Path $TextBoxMSIPackage.Text -Property ProductVersion).Trim()
        } catch { $ProductVersion = $null }
    }

    $DetectionType = if ($ProductCode -and $ProductVersion) { "MSI" } else { "Script" }



    $appParams = @{
        ApplicationName           = $ApplicationName
        Appfullname               = "$($TextBoxAppName.Text) $($TextBoxVersion.Text)"
        ApplicationVersion        = $TextBoxVersion.Text
        ApplicationDescription    = "Created by ApplicationManager"
        Publisher                 = $TextBoxPublisher.Text
        InstallationProgram       = $TextBoxInstallProgram.Text
        UninstallationProgram     = $TextBoxUnInstallProgram.Text
        ContentSourcePath         = $TextBoxSourcePath.Text
        CollectionName            = $TextBoxCollection.Text
        CollectionNameUninstall   = "$($TextBoxCollection.Text) - Uninstall"
        ADGroupName               = $TextBoxADGroup.Text
        ADGroupDescription        = "Members of this group will be targeted for deployment of $($TextBoxAppName.Text) in ConfigMgr"
        ProductCode               = $ProductCode
        ProductVersion            = $ProductVersion
        DetectionType             = $DetectionType
        DetectionMethod           = 'if (Test-Path C:\\DummyDetectionMethod) {Write-Host "IMPORTANT! This detection method does not work. You must manually change it."}'
        DeployPurpose             = $DeployPurpose
        
    }

    $globalParams = Get-GlobalDeploymentSettings
    Create-ApplicationObjects @globalParams @appParams
})

$CheckBoxCreateDeployment.Add_Checked({
    CheckBoxCreateDeploymentChanged
})
$CheckBoxCreateDeployment.Add_Unchecked({
    CheckBoxCreateDeploymentChanged
})



#Show Form
$Form.ShowDialog() | out-null

Stop-Transcript