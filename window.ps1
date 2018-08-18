[xml]$xaml = @"
<Window 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DocManager"
        x:Name="titlewindow"
        Title="" Height="360" Width="650" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">

    <Grid Margin="0,0,0,0" Background="#FF404040">
        <TabControl Margin="0,0,0,0" TabStripPlacement="Left" Background="White" SelectedIndex="6">
            <TabItem Header="Automatic bundler" MinWidth="100" MinHeight="30">
                <Grid Background="White" Margin="0,0,0,0">
                    <GroupBox Header="Workspace location" Margin="0,0,0,0" Height="70" VerticalAlignment="Top">
                        <Grid Margin="0,0,0,0">
                            <TextBox x:Name="workspace_folder" IsReadOnly="True" ScrollViewer.HorizontalScrollBarVisibility="Visible" TextWrapping="NoWrap" AcceptsReturn="True" Margin="5,5,55,5" FontWeight="Light" />
                            <Button x:Name="browse_workspace" Content="Ì" Margin="0,5" HorizontalAlignment="Right" Width="50" FontFamily="Webdings" FontSize="25" Foreground="#FF0296C1" />
                        </Grid>
                    </GroupBox>
                    <GroupBox Header="Check list" Margin="0,70,0,40">
                        <Grid Margin="0,0,0,0">
                            <GridSplitter HorizontalAlignment="Stretch" Height="50" Margin="165,10,165,0" VerticalAlignment="Top"/>
                            <CheckBox x:Name="asm_number" Content="Assembly number" Height="25" Margin="10,10,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="150" IsEnabled="False"/>
                            <CheckBox x:Name="rev_number" Content="Dublicate revision" Height="25" Margin="10,35,0,0" VerticalAlignment="Top" Width="150" VerticalContentAlignment="Center" HorizontalAlignment="Left" IsEnabled="False"/>
                            <CheckBox x:Name="fitup_check" Content="Fitup check" Height="25" Margin="10,60,0,0" VerticalAlignment="Top" Width="150" VerticalContentAlignment="Center" HorizontalAlignment="Left" IsEnabled="False"/>
                            <CheckBox x:Name="weldcard" Content="Drawings" Height="25" Margin="10,85,0,0" VerticalAlignment="Top" Width="150" VerticalContentAlignment="Center" HorizontalAlignment="Left" IsEnabled="False"/>
                            <CheckBox x:Name="drawings" Content="Welding card" Height="25" Margin="10,110,0,0" VerticalAlignment="Top" Width="150" VerticalContentAlignment="Center" HorizontalAlignment="Left" IsEnabled="False"/>
                            <CheckBox x:Name="dim_raport" Content="Dimensional raport" Height="25" Margin="10,135,0,0" VerticalAlignment="Top" Width="150" VerticalContentAlignment="Center" HorizontalAlignment="Left" IsEnabled="False"/>
                            <Label x:Name="asm_number_text" Content="" Margin="165,10,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Height="25" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="rev_number_text" Content="" Margin="165,35,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Height="25" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <ProgressBar x:Name="checklist_progress" Height="10" Margin="10,0,10,5" VerticalAlignment="Bottom" Value="0"/>
                        </Grid>
                    </GroupBox>
                    <Button x:Name="check_button" Content="Check folders" HorizontalAlignment="Left" Height="25" Margin="5,0,0,5" VerticalAlignment="Bottom" Width="90" IsEnabled="False"/>
                    <Button x:Name="make_button" Content="Make bundle" Margin="0,0,5,5" Height="25" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="90" IsEnabled="False"/>
                </Grid>
            </TabItem>
            <TabItem Header="Split/Merge PDF" MinWidth="100" MinHeight="30">
                <Grid>
                    <GroupBox Header="Folder path (click button to select files)" Height="70" VerticalAlignment="Top" >
                        <Grid>
                            <TextBox x:Name="pathtext_tab2" IsReadOnly="True" HorizontalScrollBarVisibility="Visible" Text="" FontWeight="Light" Margin="5,5,55,5"/>
                            <Button x:Name="browse_tab2" Content="Ì" HorizontalAlignment="Right" Width="50" FontFamily="Webdings" FontSize="25" Foreground="#FF0296C1" Margin="0,5" />
                        </Grid>
                    </GroupBox>
                    <GroupBox Header="Files to process" Margin="0,70,0,40" >
                        <Grid>
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition/>
                            </Grid.ColumnDefinitions>
                            <TextBox x:Name="fileslist_tab2" IsReadOnly="True" Padding="5" FontWeight="Light" Text="" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Visible" AcceptsReturn="True" TextWrapping="Wrap"/>
                        </Grid>
                    </GroupBox>
                    <Button x:Name="location_tab2" Height="25" Margin="0,0,5,5" Content="Open result" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="90"/>
                    <Button x:Name="merge_button" Width="90" Height="25" Margin="110,0,0,5" Content="Merge" VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                    <Button x:Name="split_button" Width="90" Height="25" Margin="5,0,0,5" Content="Split" VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                    <ProgressBar x:Name="progress_tab2" Height="5" Margin="5,0,5,33" VerticalAlignment="Bottom" Value="0"/>
                </Grid>
            </TabItem>
            <TabItem Header="Stamp/Watermark PDF" MinWidth="100" MinHeight="30">
                <Grid>
                    <GroupBox Header="Folder path (click button to select files)" Height="70" VerticalAlignment="Top" >
                        <Grid>
                            <TextBox x:Name="pathtext_tab3" IsReadOnly="True" HorizontalScrollBarVisibility="Visible" Text="" FontWeight="Light" Margin="5,5,55,5"/>
                            <Button x:Name="browse_tab3" Content="Ì" HorizontalAlignment="Right" Width="50" FontFamily="Webdings" FontSize="25" Foreground="#FF0296C1" Margin="0,5" />
                        </Grid>
                    </GroupBox>
                    <GroupBox Header="Stamps" Margin="0,70,0,40">
                        <Grid Margin="0,0,0,0">
                            <CheckBox x:Name="fitup_stamp" Content="FITUP" Height="18" Margin="0,0,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="prepar_stamp" Content="ETVLM" Height="18" Margin="80,0,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="dupl1_stamp" Content="Dupl.1/5" Height="18" Margin="0,18,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80" IsEnabled="False"/>
                            <CheckBox x:Name="dupl2_stamp" Content="Dupl.2" Height="18" Margin="80,18,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="dupl3_stamp" Content="Dupl.3" Height="18" Margin="0,36,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="dupl4_stamp" Content="Dupl.4" Height="18" Margin="80,36,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="dupl6_stamp" Content="Dupl.6" Height="18" Margin="0,54,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="dupl7_stamp" Content="Dupl.7" Height="18" Margin="80,54,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="dupl8_stamp" Content="Dupl.8" Height="18" Margin="0,72,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="dupl9_stamp" Content="Dupl.9" Height="18" Margin="80,72,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="80"/>
                            <CheckBox x:Name="custom45_stamp" Content="Custom 45° center large:" Height="18" Margin="0,94,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="160"/>
                            <TextBox x:Name="custom45_stamp_text" Padding="1" Height="20" Text="" FontWeight="Light" Margin="0,116,0,0" HorizontalAlignment="Left" Width="160" VerticalAlignment="Top"/>
                            <CheckBox x:Name="customdupl_stamp" Content="Custom string at bottom:" Height="18" Margin="0,140,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="160"/>
                            <TextBox x:Name="custom_stamp_text" Padding="1" Height="20" Text="" FontWeight="Light" Margin="0,162,0,0" HorizontalAlignment="Left" Width="160" VerticalAlignment="Top"/>
                            <TextBox x:Name="fileslist_tab3" IsReadOnly="True" Padding="5" Margin="170,0,0,0" FontWeight="Light" Text="" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Visible" AcceptsReturn="True" TextWrapping="Wrap"/>
                        </Grid>
                    </GroupBox>
                    <Button x:Name="location_tab3" Height="25" Margin="0,0,5,5" Content="Open result" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="90"/>
                    <Button x:Name="stamp_button" Width="90" Height="25" Margin="5,0,0,5" Content="Stamp" VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                    <ProgressBar x:Name="progress_tab3" Height="5" Margin="5,0,5,33" VerticalAlignment="Bottom" Value="0"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="tab4" Header="WPS/KK XLS" MinWidth="100" MinHeight="30">
                <Grid>
                    <GroupBox Header="File path (click button to select file)" Height="90" VerticalAlignment="Top" >
                        <Grid>
                            <TextBox x:Name="pathtext_tab4"  IsReadOnly="True" HorizontalScrollBarVisibility="Visible" Height="37" Text="" FontWeight="Light" Margin="5,5,55,0" VerticalAlignment="Top"/>
                            <Button x:Name="browse_tab4" Content="Ì" HorizontalAlignment="Right" Width="50" Height="37" FontFamily="Webdings" FontSize="25" Foreground="#FF0296C1" Margin="0,5,0,0" VerticalAlignment="Top" />
                            <TextBox x:Name="fileslist_tab4" IsReadOnly="True" Margin="5,0,0,0" Height="21" FontWeight="Light" Text="" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Visible" AcceptsReturn="True" TextWrapping="Wrap" VerticalAlignment="Bottom"/>
                        </Grid>
                    </GroupBox>
                    <GroupBox Header="Check Weldcard" Margin="0,90,0,40">
                        <Grid Margin="0,0,0,0">
                            <Label Content="E-Profiil project No:" Margin="80,0,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Tagging:" Margin="80,18,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Quantity of Products:" Margin="80,36,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Classification Rules/Standard:" Margin="80,54,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Customer project No:" Margin="80,72,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Customer PO No:" Margin="80,90,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Customer drawing No:" Margin="80,108,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Welding engineer:" Margin="80,126,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <Label Content="Designer:" Margin="80,144,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Right" HorizontalAlignment="Left" Width="162"/>
                            <GridSplitter HorizontalAlignment="Stretch" Margin="246,0,0,0"/>
                            <Label x:Name="prj_number_text" Margin="290,0,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="tagging_text" Margin="290,18,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="qty_text" Margin="290,36,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="rules_text" Margin="290,54,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="cusprj_text" Margin="290,72,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="cuspo_text" Margin="290,90,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="cusdrw_text" Margin="290,108,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="wldeng_text" Margin="290,126,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                            <Label x:Name="designer_text" Margin="290,144,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="18" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="162"  FontWeight="Light" />
                        </Grid>
                    </GroupBox>
                    <Button x:Name="location_tab4" Height="25" Margin="0,0,5,5" Content="Open result" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="90"/>
                    <Button x:Name="wps_button" Width="90" Height="25" Margin="5,0,0,5" Content="Get WPS" VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                    <Button x:Name="printpdf_button" Width="90" Height="25" Margin="110,0,0,5" Content="XLS to PDF" VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                    <Button x:Name="checkweldcard_button" Width="90" Height="25" Margin="215,0,0,5" Content="Check Card" VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                    <ProgressBar x:Name="progress_tab4" Height="5" Margin="5,0,5,33" VerticalAlignment="Bottom" Value="0"/>
                </Grid>
            </TabItem>
            <TabItem Header="Resize PDF" MinWidth="100" MinHeight="30">
                <Grid>
                    <GroupBox Header="Folder path (click button to select files)" Height="70" VerticalAlignment="Top" >
                        <Grid>
                            <TextBox x:Name="pathtext_tab5"  IsReadOnly="True" HorizontalScrollBarVisibility="Visible" FontWeight="Light" Margin="5,5,55,5"/>
                            <Button x:Name="browse_tab5" Content="Ì" HorizontalAlignment="Right" Width="50" FontFamily="Webdings" FontSize="25" Foreground="#FF0296C1" Margin="0,5" />
                        </Grid>
                    </GroupBox>
                    <GroupBox Header="Formats and document screen pixels" Margin="0,70,0,40">
                        <Grid Margin="0,0,0,0">
                            <CheckBox x:Name="a4_resize" Content="A4" Height="25" Margin="10,10,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="100"/>
                            <CheckBox x:Name="a3_resize"  Content="A3" Height="25" Margin="10,35,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="100"/>
                            <CheckBox x:Name="a2_resize" Content="A2" Height="25" Margin="10,60,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="100"/>
                            <CheckBox x:Name="a1_resize" Content="A1" Height="25" Margin="10,85,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="100"/>
                            <CheckBox x:Name="custom_resize" Content="Custom:" Height="25" Margin="10,110,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" HorizontalAlignment="Left" Width="100"/>
                            <TextBox x:Name="width_resize" Padding="3" FontWeight="Light" Margin="80,110,0,0" HorizontalAlignment="Left" Width="40" Height="25" VerticalAlignment="Top" HorizontalContentAlignment="Right"/>
                            <TextBox x:Name="height_resize" Padding="3" FontWeight="Light" Margin="0,110,317,0" Height="25" VerticalAlignment="Top" HorizontalAlignment="Right" Width="40" HorizontalContentAlignment="Left"/>
                            <GridSplitter HorizontalAlignment="Stretch" Height="110" Margin="75,0,317,0" VerticalAlignment="Top"/>
                            <Label Content="595px x 842px" Margin="75,10,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="25" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="100" FontWeight="Light"/>
                            <Label Content="1191px x 842px" Margin="75,35,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="25" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="100" FontWeight="Light"/>
                            <Label Content="1648px x 1191px" Margin="75,60,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="25" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="100" FontWeight="Light"/>
                            <Label Content="2384px x 1648px" Margin="75,85,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="25" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="100" FontWeight="Light"/>
                            <Label Content="Width x Height" Margin="75,135,0,0" VerticalAlignment="Top" VerticalContentAlignment="Center" Padding="0" Height="25" HorizontalContentAlignment="Center" HorizontalAlignment="Left" Width="100" FontWeight="Light"/>
                            <TextBox x:Name="fileslist_tab5" IsReadOnly="True" Padding="5" FontWeight="Light" HorizontalScrollBarVisibility="Disabled" VerticalScrollBarVisibility="Visible" AcceptsReturn="True" TextWrapping="Wrap" HorizontalAlignment="Right" Width="312"/>
                        </Grid>
                    </GroupBox>
                    <Button x:Name="location_tab5" Height="25" Margin="0,0,5,5" Content="Open result" VerticalAlignment="Bottom" HorizontalAlignment="Right" Width="90"/>
                    <Button x:Name="resize_button" Width="90" Height="25" Margin="5,0,0,5" Content="Resize" VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                    <ProgressBar x:Name="progress_tab5" Height="5" Margin="5,0,5,33" VerticalAlignment="Bottom" Value="0"/>
                </Grid>
            </TabItem>
            <TabItem Header="Rotate PDF" MinWidth="100" MinHeight="30" IsEnabled="False"></TabItem>
            <TabItem Header="Tools" MinWidth="100" MinHeight="30">
                <Grid>
                    <GroupBox Header="Drawings management" Height="70" VerticalAlignment="Top" >
                        <Grid>
                            <Button x:Name="sort_drws" Width="90" Height="25" Margin="5,0,0,5" Content="Sort drawings..." VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                            <Button x:Name="find_dublicates" Width="90" Height="25" Margin="100,0,0,5" Content="Find dublicates..." VerticalAlignment="Bottom" HorizontalAlignment="Left"/>
                        </Grid>
                    </GroupBox>
                    <ProgressBar x:Name="progress_tab7" Height="5" Margin="0,0,0,0" VerticalAlignment="Bottom" Value="0"/>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Window>
"@
$syncHash = [hashtable]::Synchronized(@{})
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$syncHash.Window=[Windows.Markup.XamlReader]::Load($reader)
$reader.Dispose()
$xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | 
        
    foreach {
        
        $syncHash.Add( $_.name , $syncHash.Window.FindName( $_.name ) )
        
    }