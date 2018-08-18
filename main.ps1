#region Copyrights

  <# 'DocManager' is a powershell based application for documents organizing
     using iTextSharp assembly for PDFs manipulation.
     Copyright (C) 2017  Vladimir Mihhejenko
 
     This program is free software: you can also redistribute it and/or 
     modify it under the terms of the GNU Affero General Public License 
     as published by the Free Software Foundation, either version 3 of the 
     License, or any later version.

     This program is distributed in the hope that it will be useful,
     but WITHOUT ANY WARRANTY; without even the implied warranty of
     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
     GNU Affero General Public License for more details.
 
     You should have received a copy of the GNU Affero General Public License 
     along with this program.  If not, see <http://www.gnu.org/licenses/>. Or
     write to the Free Software Foundation, Inc., 51 Franklin Street, 
     Fifth Floor, Boston, MA 02110-1301 USA.
 
     For contacts use vovit@mail.ru #>

#endregion
$Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open();$PSinstance1=[powershell]::Create().AddScript({ # main

#region Globals
$scriptFiles = 'get-folder.ps1','doc-helper.ps1','get-files.ps1','window.ps1'
$global:defaultLocation = @( 

    'X:\YOUR PATH HERE\DocManager\bin'

) | foreach { if(Test-Path $_){$_} }

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms 

try {   Add-Type -Path (Join-Path $defaultLocation -ChildPath 'itextsharp.dll') 
        $scriptFiles | % { Import-Module (Join-Path $defaultLocation -ChildPath $_) } 
} catch { [Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error") }

$patchnotes = (Join-Path $defaultLocation -ChildPath 'patchnotes.ini')
Set-ItemProperty -Path $patchnotes -Name IsReadOnly -Value $true
$lastversion = (Get-Content $patchnotes | Select-String 'version\s+(\d+.*)' -AllMatches | % { $_.Matches.Groups[1].value }) | select -First 1
$lastverfile = $lastversion + '.txt'
if(Test-Path $home\$lastverfile){ }
else { 
        New-Item -Path $home -Name $lastverfile
        Start-Process notepad -WindowStyle Normal -ArgumentList $patchnotes
        (get-content $patchnotes) | Out-File $home\$lastverfile
     }
$syncHash.Window.Title = 'Technical documentation manager v.{0}' -f $lastversion

#endregion
#region Functions

function Fill-Files {
    
    Param($syncHash, $pathText, $fileList, $type)
        
    $Global:files = $null
    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.$pathText.Clear() })
    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.$fileList.Clear() })
    $Global:files = [GetFiles]::init().Search($type)

    if ($files.directory -match 'C:\.*') {

        $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.$pathText.Text = $files.directory })
        $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.$fileList.Text = $files.baseNames })

    } else {
            
        $Global:files = $null
        [Microsoft.VisualBasic.Interaction]::MsgBox("Not valid location, select files on disk 'C:'","Critical,OkOnly","Error")

    }

}

#endregion
#region Actions

    # tabindex 0
    $syncHash.browse_workspace.Add_Click({

        $syncHash.check_button.IsEnabled = $false
        $syncHash.make_button.IsEnabled = $false
    
        $global:workspace = $null
        $syncHash.make_button.IsEnabled = $false  
        $syncHash.workspace_folder.clear()
        $syncHash.make_button.IsEnabled = $false
        $syncHash.asm_number.IsChecked = $false
        $syncHash.rev_number.IsChecked = $false
        $syncHash.fitup_check.IsChecked = $false
        $syncHash.drawings.IsChecked = $false
        $syncHash.weldcard.IsChecked = $false
        $syncHash.dim_raport.IsChecked = $false
        $syncHash.asm_number_text.Content=$null
        $syncHash.rev_number_text.Content=$null

        $workspace = Get-Folder

        if($workspace.length -gt 3 -and (Test-Path $workspace) -and ($workspace -match 'C:\.*') ){

            $workElements = '01_Fitup_NOV','02_EP_BOM','03_WPS_KK','04_Drawings','05_Dim_Raport'
            $syncHash.workspace_folder.Clear()
        
            $newBundleFolder = (Get-Date -f yyMMddHHmmssss)
            $workspace = New-Item -Path ( Join-Path -Path $workspace -ChildPath $newBundleFolder ) -ItemType Directory
            $global:workspaces = $workElements | % { New-Item -Path ( Join-Path -Path $workspace.FullName -ChildPath $_ ) -ItemType Directory }
            $syncHash.workspace_folder.Text = $workspace

            if ([Microsoft.VisualBasic.Interaction]::MsgBox("Do you want to choose files manually","Question,YesNo","Question") -like 'yes'){
            
                [GetFiles]::init().setTitle('01_Fitup_NOV: Select drawings').Search('pdf')  | foreach {Copy-Item $_.fullPaths $workspaces[0]}
                [GetFiles]::init().setTitle('02_EP_BOM: Select BOM').Search('pdf')          | foreach {Copy-Item $_.fullPaths $workspaces[1]}
                [GetFiles]::init().setTitle('03_WPS_KK: Select welding card').Search('xls') | foreach {Copy-Item $_.fullPaths $workspaces[2]}
                [GetFiles]::init().setTitle('04_Drawings: Select drawings').Search('pdf')   | foreach {Copy-Item $_.fullPaths $workspaces[3]}
                [GetFiles]::init().setTitle('05_Dim_Raport: Select raport').Search('xls')   | foreach {Copy-Item $_.fullPaths $workspaces[4]}

            }
            else {

                [Microsoft.VisualBasic.Interaction]::MsgBox("Put files in folders before continue","Information,OkOnly","Information")
                Start-Process explorer -WindowStyle Maximized -ArgumentList $workspace
            
            }

            $syncHash.check_button.IsEnabled = $true
        }
        else {[Microsoft.VisualBasic.Interaction]::MsgBox("Not valid location, select folder on disk 'C:'","Critical,OkOnly","Error")}

        $global:workspace = $workspace 

    })
    $syncHash.check_button.Add_Click({ 
        
        $syncHash.checklist_progress.Value = 0
        $syncHash.check_button.IsEnabled = $false
        $syncHash.asm_number.IsChecked = $false
        $syncHash.rev_number.IsChecked = $false
        $syncHash.fitup_check.IsChecked = $false
        $syncHash.drawings.IsChecked = $false
        $syncHash.weldcard.IsChecked = $false
        $syncHash.dim_raport.IsChecked = $false

        $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
        $Runspace.SessionStateProxy.SetVariable("workspace" , $workspace )
        $Runspace.SessionStateProxy.SetVariable("workspaces" , $workspaces )
        $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )

        $code = {
            
            Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')
            
            $global:workFiles = @{}
            if ($workspace -ne $null) {

                for($i=0;$i -lt $workspaces.count;$i++){$global:workFiles.($workspaces[$i].basename)+=gci -Path $workspaces[$i] -Filter *}
                if ($workFiles.'01_Fitup_NOV'.Count -gt 0) {
                    
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $true })
                    $workFiles.'01_Fitup_NOV' | foreach {
                        
                        try{[void][DocHelper]::init($defaultLocation).read($_).stamp('angle','FIT UP CHECK').stamp('flatten').resizeTo(1191,842)}
                        catch{[Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error")}
                
                    }
                    
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $false })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.fitup_check.IsChecked = $true })
                
                }
                if ($workFiles.'02_EP_BOM'.Count -gt 0) {

                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $true })

                    try{$parseBom = [DocHelper]::init($defaultLocation).read( $workFiles.'02_EP_BOM'[0] ).parseBOM()} # get first file in folder
                    catch{[Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error")}
                   
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.asm_number_text.Content = $parseBom.asmNumber })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.asm_number.IsChecked = $true })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.rev_number_text.Content = $parseBom.revNumber })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.rev_number.IsChecked = $true })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $false })

                    $global:parseBom = $parseBom
                }
                if ($workFiles.'03_WPS_KK'.Count -gt 0) {

                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $true })

                    $file = ($workFiles.'03_WPS_KK') | % { if ($_.Extension -eq ".xls"){$_} }
                    [void][DocHelper]::init($defaultLocation).wpsFrom( $file )
                    
                    $file = ($workFiles.'03_WPS_KK') | foreach { if ($_.Extension -eq ".xls"){$_} }
                    [void][DocHelper]::init($defaultLocation).excelToPdf( $file )
                    
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.weldcard.IsChecked = $true })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $false })

                }
                if ($workFiles.'04_Drawings'.Count -gt 0) { 
                
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $true })

                    Sort-Drawings $workFiles.'04_Drawings'

                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.drawings.IsChecked = $true })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $false })

                }
                if ($workFiles.'05_Dim_Raport'.Count -gt 0) {

                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $true })

                    $file = ($workFiles.'05_Dim_Raport') | foreach { if ($_.Extension -eq ".xls"){$_} }
                    [void][DocHelper]::init($defaultLocation).excelToPdf( $file )
                    
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.dim_raport.IsChecked = $true })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $false })

            
                }
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.make_button.IsEnabled = $true })
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.Value = 100 })

            }
            else { [Microsoft.VisualBasic.Interaction]::MsgBox("No folder selected",'OKOnly,Information',"Information") }

        }
        $PSinstance1=[powershell]::Create().AddScript($code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()
        
    })
    $syncHash.make_button.Add_Click({
        
        $syncHash.make_button.IsEnabled = $false

        $syncHash.checklist_progress.Value = 0
        $syncHash.checklist_progress.IsIndeterminate = $true
         
        $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
        $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
        $Runspace.SessionStateProxy.SetVariable("workspace" , $workspace )
        $Runspace.SessionStateProxy.SetVariable("asm_number" , $syncHash.asm_number_text.Content )
        $Runspace.SessionStateProxy.SetVariable("rev_number" , $syncHash.rev_number_text.Content )
        $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )
        
        $code = {

            Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')

            $finalFolders = gci $workspace * -Directory
            
            $fitupFiles = $finalFolders[0].GetFiles() | where { $_.Extension -match '.pdf' } | sort { [regex]::Replace($_, '\d+',{$args[0].Value.Padleft(20)})}
            $bomFile    = $finalFolders[1].GetFiles() | where { $_.Extension -match '.pdf' } | select -First 1
            $wpsFiles   = $finalFolders[2].GetFiles() | where { $_.Extension -match '.pdf' } | sort -Descending
            $wldFile    = $finalFolders[2].GetFiles() | where { $_.Extension -match '.pdf' } | sort | select -First 1
            $drwFiles   = $finalFolders[3].GetFiles() | where { $_.Extension -match '.pdf' } | sort { [regex]::Replace($_, '\d+',{$args[0].Value.Padleft(20)})}
            $dimRaport  = $finalFolders[4].GetFiles() | where { $_.Extension -match '.pdf' }

            if ($asm_number.length -lt 1){$asm_number=[Microsoft.VisualBasic.Interaction]::InputBox("Enter assembly number","User input","XX-XXXX/X/X",100,100)}
            if ($rev_number.length -lt 1){$rev_number=[Microsoft.VisualBasic.Interaction]::InputBox("Enter document revision","User input","X",100,100)}

            $filename = "{0}_V.{1}" -f ($asm_number -replace "(\b\d\b)",'0$1' -replace "\/",'_'), $rev_number

            $dublicates = ($fitupFiles, $bomFile, $wpsFiles, $drwFiles), # Dublicate 1  + w/o stamps
                          ($fitupFiles, $bomFile, $wpsFiles, $drwFiles), # Dublicate 2  +
                          ($bomFile, $wpsFiles, $drwFiles),              # Dublicate 3  + w/o fitup
                          (0),                                           # Dublicate 4  (?)
                          (0),                                           # Dublicate 5  -
                          (0),                                           # Dublicate 6  -
                          ($wpsFiles, $drwFiles),                        # Dublicate 7  +
                          (0),                                           # Dublicate 8  (?)
                          (0),                                           # Dublicate 9  (?)
                          (0)                                            # Dublicate 10 (?)
            
            if ([int16]$rev_number -gt 0) { $dublicates[1]=$dublicates[2] = ($wldFile, $drwFiles) }

            for ($i = 0 ; $i -lt $dublicates.Count ; $i++){ 
                
                $newFilename = Join-Path $workspace -ChildPath ("{0}_Dupl.{1}.pdf" -f $filename, ($i + 1))
                [void][DocHelper]::init($defaultLocation).mergeFilesTo((Get-Item $dublicates[$i].fullname),$newFilename)
                if ($i -eq 0){ continue }
                [void][DocHelper]::init($defaultLocation).read((Get-Item $newFilename)).stamp('str',("{0}_Dupl.{1}" -f $filename, ($i + 1)))
            
            }
            try {   (Get-Item $workspace).GetFiles() | foreach { [void][DocHelper]::init($defaultLocation).read( $_ ).splitByFormats() } } 
            catch { [Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error") }
            
            $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.Value = 100 })
            $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.checklist_progress.IsIndeterminate = $false })
            
        }
        $PSinstance1=[powershell]::Create().AddScript($code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()
    
    })
    # tabindex 1
    $syncHash.browse_tab2.Add_Click({ Fill-Files -syncHash $syncHash -pathText "pathtext_tab2" -fileList "fileslist_tab2" -type 'pdf' })
    $syncHash.merge_button.Add_Click({ 

        if ($files -ne $null -and $syncHash.fileslist_tab2.Text.Length -gt 0) {
            
            $syncHash.progress_tab2.Value = 0
            $syncHash.progress_tab2.IsIndeterminate = $true
            $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.fileslist_tab2.Clear()})

            $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
            $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
            $Runspace.SessionStateProxy.SetVariable("files" , $files )
            $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )

            $code = {

                Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')

                if ( $files.directory -match ' - Directory' ) {

                    $root = Get-Item (Join-Path $files.directory -ChildPath '\..\')
                    $filename = ($files.directory.Trim(' - Directory')).split('\') | select -Last 1

                    [void][DocHelper]::init($defaultLocation).mergeFilesTo((Get-Item $files.fullPaths),(Join-Path $root -ChildPath $filename))

                    if ([Microsoft.VisualBasic.Interaction]::MsgBox("Do you wish to remove merged files","Question,YesNo","Merging") -like 'yes') {
                    
                        try { Remove-Item $files.directory -Recurse -ErrorAction Stop } catch { 
                            [Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error") }
                        $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.pathtext_tab2.Text = $root.FullName })

                    }
                    
                }
                else {
                
                    [void][DocHelper]::init($defaultLocation).mergeFilesTo(
                    
                        (Get-Item $files.fullPaths),
                        (Join-Path $files.directory -ChildPath (
                        
                            '_merged_'+(Get-Date -f HHmmss-dd-MM-yyyy)+'.pdf')
                    
                        )
                
                    )
                    if ([Microsoft.VisualBasic.Interaction]::MsgBox("Do you wish to remove merged files","Question,YesNo","Merging") -like 'yes') {
                    
                        try { Remove-Item $files.fullPaths -ErrorAction Stop } catch { 
                        [Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error") }

                    }

                }

                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab2.IsIndeterminate = $false })
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab2.Value = 100 })
                $Global:files = $null
                [Microsoft.VisualBasic.Interaction]::MsgBox("Done",'OKOnly,Information',"The operation completed successfully")

            }

            $PSinstance1=[powershell]::Create().AddScript($Code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()

        }
        else { [Microsoft.VisualBasic.Interaction]::MsgBox("No files selected",'OKOnly,Information',"Information") }

    })
    $syncHash.split_button.Add_Click({

        if ($files -ne $null -and $syncHash.fileslist_tab2.Text.Length -gt 0) {
            
            $syncHash.progress_tab2.Value = 0
            $syncHash.progress_tab2.IsIndeterminate = $true
            $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.fileslist_tab2.Clear()})
            $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
            $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
            $Runspace.SessionStateProxy.SetVariable("files" , $files )
            $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )

            $code = {
                
                Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')
                    
                $files.fullPaths | foreach { [void][DocHelper]::init($defaultLocation).read( (Get-Item $_) ).split() }
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab2.IsIndeterminate = $false })
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab2.Value = 100 })
                $Global:files = $null
                [Microsoft.VisualBasic.Interaction]::MsgBox("Done",'OKOnly,Information',"The operation completed successfully")

            }
            
            $PSinstance1=[powershell]::Create().AddScript($Code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()

        }
        else { [Microsoft.VisualBasic.Interaction]::MsgBox("No files selected",'OKOnly,Information',"Information") }

    })
    $syncHash.location_tab2.Add_Click({ Start-Process explorer -WindowStyle Maximized -ArgumentList $syncHash.pathtext_tab2.Text })
    # tabindex 2
    $syncHash.browse_tab3.Add_Click({ Fill-Files -syncHash $syncHash -pathText "pathtext_tab3" -fileList "fileslist_tab3" -type 'pdf' })
    $syncHash.stamp_button.Add_Click({

        $checkBox = @{ 
            
            'Fitup' = $syncHash.fitup_stamp.IsChecked 
            'Prep' = $syncHash.prepar_stamp.IsChecked
            'Dupl2' = $syncHash.dupl2_stamp.IsChecked
            'Dupl3' = $syncHash.dupl3_stamp.IsChecked
            'Dupl4' = $syncHash.dupl4_stamp.IsChecked
            'Dupl6' = $syncHash.dupl6_stamp.IsChecked
            'Dupl7' = $syncHash.dupl7_stamp.IsChecked
            'Dupl8' = $syncHash.dupl8_stamp.IsChecked
            'Dupl9' = $syncHash.dupl9_stamp.IsChecked
            'Custom45' = $syncHash.custom45_stamp.IsChecked
            'CustDupl' = $syncHash.customdupl_stamp.IsChecked 
        
        }

        $checkBox.Values | foreach { if ($_) { $counter++ } }

        if ($counter -lt 1) { [Microsoft.VisualBasic.Interaction]::MsgBox("No stamps selected",'OKOnly,Question',"Warning") }
        elseif ($counter -gt 1) { [Microsoft.VisualBasic.Interaction]::MsgBox("Multiple stamps selected",'OKOnly,Question',"Warning") }
        else {

            if ($files -ne $null -and $syncHash.fileslist_tab3.Text.Length -gt 0) {
            
                $syncHash.progress_tab3.Value = 0
                $syncHash.progress_tab3.IsIndeterminate = $true
                #$syncHash.Window.Dispatcher.invoke([action]{ $syncHash.fileslist_tab3.Clear()})
                $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
                $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
                $Runspace.SessionStateProxy.SetVariable("files" , $files )
                $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )
                $Runspace.SessionStateProxy.SetVariable("checkBox" , $checkBox )
                $Runspace.SessionStateProxy.SetVariable("custom45", $syncHash.custom45_stamp_text.Text) 
                $Runspace.SessionStateProxy.SetVariable("custString", $syncHash.custom_stamp_text.Text)
                
                $code = {
                    
                    Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')

                    $files.fullPaths | foreach { 
                
                        $file = [DocHelper]::init($defaultLocation).read((Get-Item $_))
                        $name = ($_.split('\') | select -Last 1).trim('.pdf')
                        
                        if ( $checkBox.Fitup    ) { $file.stamp('angle', 'FIT UP CHECK') }
                        if ( $checkBox.Prep     ) { $file.stamp('angle', 'ETTEVALMISTUS') }
                        if ( $checkBox.Dupl2    ) { $file.stamp('str', "{0}_Dupl.2" -f $name) }
                        if ( $checkBox.Dupl3    ) { $file.stamp('str', "{0}_Dupl.3" -f $name) }
                        if ( $checkBox.Dupl4    ) { $file.stamp('str', "{0}_Dupl.4" -f $name) }
                        if ( $checkBox.Dupl6    ) { $file.stamp('str', "{0}_Dupl.6" -f $name) }
                        if ( $checkBox.Dupl7    ) { $file.stamp('str', "{0}_Dupl.7" -f $name) }
                        if ( $checkBox.Dupl8    ) { $file.stamp('str', "{0}_Dupl.8" -f $name) }
                        if ( $checkBox.Dupl9    ) { $file.stamp('str', "{0}_Dupl.9" -f $name) }
                        if ( $checkBox.Custom45 ) { $file.stamp('angle', $custom45 ) }
                        if ( $checkBox.CustDupl ) { $file.stamp('str', $custString ) } 

                    }

                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab3.IsIndeterminate = $false })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab3.Value = 100 })
                    [Microsoft.VisualBasic.Interaction]::MsgBox("Done",'OKOnly,Information',"The operation completed successfully")

                }

                $PSinstance1=[powershell]::Create().AddScript($Code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()
            
            }
            else { [Microsoft.VisualBasic.Interaction]::MsgBox("No files selected",'OKOnly,Question',"Warning") }
        }
    })
    $syncHash.location_tab3.Add_Click({ Start-Process explorer -WindowStyle Maximized -ArgumentList $syncHash.pathtext_tab3.Text })
    # tabindex 3
    $syncHash.browse_tab4.Add_Click({ 

        $Global:textElement = 'prj_number_text','tagging_text','qty_text','rules_text','cusprj_text','cuspo_text','cusdrw_text','wldeng_text','designer_text'
        $textElement | foreach { $syncHash.$_.Content = $null }
        Fill-Files -syncHash $syncHash -pathText "pathtext_tab4" -fileList "fileslist_tab4" -type 'xls' 
    
    })
    $syncHash.location_tab4.Add_Click({ Start-Process explorer -WindowStyle Maximized -ArgumentList $syncHash.pathtext_tab4.Text })
    $syncHash.checkweldcard_button.Add_Click({ 

        $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.Value = 0 })

        if ($files -ne $null -and $syncHash.fileslist_tab4.Text.Length -gt 0) {

            $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()

            $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
            $Runspace.SessionStateProxy.SetVariable("textElement" , $textElement )
            $Runspace.SessionStateProxy.SetVariable("files" , $files )
                
            $code = {

                $counter = 0
                $grid = (4,7),(5,7),(6,7),(7,7),(4,22),(5,22),(6,22),(7,22),(8,22)
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.IsIndeterminate = $true })
                $objExcel = New-Object -ComObject Excel.Application
                $objExcel.Visible = $true
                $workBooks = $objExcel.Workbooks
                $workBook = $workBooks.Open( $files.fullPaths, 3)
                $objExcel.Visible = $false
                $workBook.Saved = $true
                $sheet = $workBook.Worksheets.Item(1)
                for ( $i = 0 ; $i -lt $textElement.Count ; $i++ ){

                        #$syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.Value += 100/$textElement.Count })
                        $cell = $sheet.Cells( $grid[$i][0] , $grid[$i][1] )
                        $value = $sheet.Range( $cell , $cell ).MergeArea.Cells( 1 , 1 ).value2
                        if ($value.Length -gt 0) { 
                        
                            $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.($textElement[$i]).Content = $value }) 
                     
                        }
                        else{ $counter++ } 
            
                }

                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.IsIndeterminate = $false })
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.Value = 100 })
                
                if ($counter -eq 0) {

                       [Microsoft.VisualBasic.Interaction]::MsgBox("Done",'OKOnly,Information',"Wps checking result")        }
                else { [Microsoft.VisualBasic.Interaction]::MsgBox("{0} of 9 cells are empty" -f $counter,'OKOnly,Critical',"Wps checking result") }

                $workBook.Close($false)
                $objExcel.Quit()

                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workBook)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workBooks)
                [System.Runtime.InteropServices.Marshal]::ReleaseComObject($objExcel)
        
                Remove-Variable -Name objExcel  
            }

            $PSinstance1=[powershell]::Create().AddScript($Code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()

        }
        else { [Microsoft.VisualBasic.Interaction]::MsgBox("Select file",'OKOnly,Information',"Wps checking result") }

    })
    $syncHash.printpdf_button.Add_Click({ 
        
        $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.Value = 0 })

        if ($files -ne $null -and $syncHash.fileslist_tab4.Text.Length -gt 0) {

            $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
        
            $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
            $Runspace.SessionStateProxy.SetVariable("files" , $files )
            $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )
                
            $code = {

                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.IsIndeterminate = $true })

                Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')
                
                [void][DocHelper]::init($defaultLocation).excelToPdf( (Get-Item $files.fullPaths) )
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.IsIndeterminate = $false })
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.Value = 100 })
                [Microsoft.VisualBasic.Interaction]::MsgBox("Done",'OKOnly,Information',"Welding card pdf export")

            }

            $PSinstance=[powershell]::Create().AddScript($code);$PSinstance.Runspace=$Runspace;$job=$PSinstance.BeginInvoke()
        }
        else { [Microsoft.VisualBasic.Interaction]::MsgBox("Select file",'OKOnly,Information',"Welding card pdf export") }
        
    })
    $syncHash.wps_button.Add_Click({
    
        $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.Value = 0 })

        if ($files -ne $null -and $syncHash.fileslist_tab4.Text.Length -gt 0) {

            $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
        
            $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
            $Runspace.SessionStateProxy.SetVariable("files" , $files )
            $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )

            $code = {
                    
                Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')

                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.IsIndeterminate = $true })
                [void][DocHelper]::init($defaultLocation).wpsFrom( (Get-Item $files.fullPaths) )
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.IsIndeterminate = $false })
                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab4.Value = 100 }) 

            }

            $PSinstance=[powershell]::Create().AddScript($code);$PSinstance.Runspace=$Runspace;$job=$PSinstance.BeginInvoke()

        }
        else { [Microsoft.VisualBasic.Interaction]::MsgBox("Select file",'OKOnly,Information',"Wps from welding card") }

    })
    # tabindex 4
    $syncHash.browse_tab5.Add_Click({ Fill-Files -syncHash $syncHash -pathText "pathtext_tab5" -fileList "fileslist_tab5" -type 'pdf' })
    $syncHash.location_tab5.Add_Click({ Start-Process explorer -WindowStyle Maximized -ArgumentList $syncHash.pathtext_tab5.Text })
    $syncHash.resize_button.Add_Click({
    
        $checkBox = @{ 
            
            'A4' = $syncHash.a4_resize.IsChecked 
            'A3' = $syncHash.a3_resize.IsChecked
            'A2' = $syncHash.a2_resize.IsChecked
            'A1' = $syncHash.a1_resize.IsChecked
            'Custom' = $syncHash.custom_resize.IsChecked
        
        }

        $checkBox.Values | foreach { if ($_) { $counter++ } }
        if ($counter -lt 1) { [Microsoft.VisualBasic.Interaction]::MsgBox("No format selected",'OKOnly,Question',"Warning") }
        elseif ($counter -gt 1) { [Microsoft.VisualBasic.Interaction]::MsgBox("Multiple formats selected",'OKOnly,Question',"Warning") }
        else {

            if ($files -ne $null -and $syncHash.fileslist_tab5.Text.Length -gt 0) {
                
                $syncHash.progress_tab5.Value = 0
                $syncHash.progress_tab5.IsIndeterminate = $true
                #$syncHash.Window.Dispatcher.invoke([action]{ $syncHash.fileslist_tab5.Clear()})
                $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
                $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
                $Runspace.SessionStateProxy.SetVariable("files" , $files )
                $Runspace.SessionStateProxy.SetVariable("defaultLocation" , $defaultLocation )
                $Runspace.SessionStateProxy.SetVariable("checkBox" , $checkBox )
                $Runspace.SessionStateProxy.SetVariable("cusWidth", $syncHash.width_resize.Text -as [int16] ) 
                $Runspace.SessionStateProxy.SetVariable("cusHeight", $syncHash.height_resize.Text -as [int16] )

                $code = {
                
                    Import-Module (Join-Path $defaultLocation -ChildPath 'doc-helper.ps1')

                    $files.fullPaths | foreach { 
                
                        $file = [DocHelper]::init($defaultLocation).read((Get-Item $_))
                        $file.stamp('flatten')
                        
                        if ( $checkBox.A4     ) { $file.resizeTo( 595 ,  842  ) }
                        if ( $checkBox.A3     ) { $file.resizeTo( 1191 , 842  ) }
                        if ( $checkBox.A2     ) { $file.resizeTo( 1648 , 1191 ) }
                        if ( $checkBox.A1     ) { $file.resizeTo( 2384 , 1648 ) }
                        if ( $checkBox.Custom ) { $file.resizeTo($cusWidth, $cusHeight) }

                    }

                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab5.IsIndeterminate = $false })
                    $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab5.Value = 100 })
                    [Microsoft.VisualBasic.Interaction]::MsgBox("Done",'OKOnly,Information',"The operation completed successfully")

                }

                $PSinstance1=[powershell]::Create().AddScript($Code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()

            }
            else { [Microsoft.VisualBasic.Interaction]::MsgBox("No files selected",'OKOnly,Question',"Warning") }
        }
    })
    # tabindex 5
    # tabindex 6
    # tabindex 7
    $syncHash.sort_drws.Add_Click({
    
        $files = [GetFiles]::init().Search('pdf')

        if ($files.directory -match 'C:\.*') {

            Sort-Drawings -Workfiles (Get-Item $files.fullPaths)
            Start-Process explorer -WindowStyle Maximized -ArgumentList $files.directory
            $files = $null 

        } else {
            
            $files = $null
            [Microsoft.VisualBasic.Interaction]::MsgBox("Not valid location, select files on disk 'C:'","Critical,OkOnly","Error")

        }

    })
    $syncHash.find_dublicates.Add_Click({
         
        $folder = Get-Folder

        if($folder.length -gt 3 -and (Test-Path $folder) ){

            $syncHash.progress_tab7.IsIndeterminate = $true

            $Runspace=[runspacefactory]::CreateRunspace();$Runspace.ApartmentState="STA";$Runspace.ThreadOptions="ReuseThread";$Runspace.Open()
            $Runspace.SessionStateProxy.SetVariable("syncHash" , $syncHash )
            $Runspace.SessionStateProxy.SetVariable("folder" , $folder )

            $code = {

                [System.Collections.ArrayList] $strings = @()
                [System.Object[]] $folders = @()

                $folders = Get-ChildItem $folder -Recurse -Directory 
                $folders += Get-Item $folder
            
                $folders | foreach {   

                    $path = $_
                    if ( -not [regex]::Match($path.Name,'(?i)ar[ch]+i+v[e]?').success ) {

                        [System.Collections.ArrayList] $a = @()
                        Get-ChildItem -Path $path.FullName -Filter *.pdf |
        
                            foreach { [void]$a.Add( ($_.BaseName -replace '_Rev.\d+','') ) }
                            $t = $a | Select-Object -Unique 
                            if($t -ne $null) { 
                
                                Compare-Object $a $t |
                
                                foreach { 
                                    $hash = [ordered]@{ 

                                                Match = $_.InputObject
                                                Folder = $path.Name
                                                FullPath = $path.FullName
                                                                
                                    }
                                    $strings += New-Object psobject -Property $hash
        
                                }

                            }
                    }
                }
                if ($strings.Count -ge 1) {
                
                    $strings | Out-GridView -Title 'Duplicate search results'

                }
                else {[Microsoft.VisualBasic.Interaction]::MsgBox("No dublicates found","Information,OkOnly","Info")}

                $syncHash.Window.Dispatcher.invoke([action]{ $syncHash.progress_tab7.IsIndeterminate = $false })

            }

            $PSinstance1=[powershell]::Create().AddScript($Code);$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke()
        }
        else {[Microsoft.VisualBasic.Interaction]::MsgBox("Not valid location","Critical,OkOnly","Error")}
    })

#endregion
$syncHash.Window.ShowDialog();$Runspace.Close();$Runspace.Dispose();Stop-Process -processname powershell # close console from other thread

});$PSinstance1.Runspace=$Runspace;$job=$PSinstance1.BeginInvoke() # main