<# TODO's 

    1. Add check, if wps not filled

#>
class DocHelper {
    
    # PROPERTIES
    
    hidden [System.IO.FileSystemInfo] $inputFile
    hidden [String] $output
    hidden [String] $customStamp
    hidden [String] $archive 
    hidden [String] $wpsRegister
    [String] $revNumber
    [String] $asmNumber
    [String] $asmFileName
    [System.Collections.ArrayList] $bomFiles = @()
    
    # INITS
    
    hidden DocHelper( $defaultLocation ) {
        
        $section = $null
        $settings = $ini = @{} ; switch -Regex -File $defaultLocation\settings.ini {
            "^\[(.+)\]" {$section=$matches[1];$ini[$section]=@{}}
            "(.+?)\s*=>(.*)"{$name,$value=$matches[1..2];$ini[$section][$name]=$value}}
        $this.archive = $settings['archive'].Values | % { if (Test-Path $_){$_}}
        $this.wpsRegister = $settings['wpsregister'].Values | % { if (Test-Path $_){$_}}

    } # Default, Parameterless Constructor
    static [DocHelper] init([String] $defaultLocation) { 
        
        return [DocHelper]::New($defaultLocation) 
    
    } # Named Constructor    
    hidden [iTextSharp.text.pdf.PdfReader] initreader() { 
        
        return [iTextSharp.text.pdf.PdfReader]::new($this.inputFile.FullName) ;  
        
    } # reader 

    # METHODS CHAINING
    
    [DocHelper] read([System.IO.FileSystemInfo] $inputFile) {
        
        $this.inputFile = $inputFile
        Set-ItemProperty $inputFile IsReadOnly -Value $false
        return $this # chaining
    
    } # End
    [DocHelper] to([String] $output){
    
        $this.output = $output ; return $this #chaining

    } # End
    [DocHelper] resizeTo([int16] $width, [int16] $height) {

        $reader = $this.initreader()

        for ($pNum = 1 ; $pNum -le $reader.NumberOfPages ; $pNum++){
            if ($reader.GetPageSizeWithRotation($pNum).Rotation -ne 0){
                [Microsoft.VisualBasic.Interaction]::MsgBox("Please normilize rotation before resizing","Critical,OkOnly","Warning")
                $reader.Close()
                return $this
            }
        }

        [iTextSharp.text.pdf.PdfReader]::unethicalreading = $true
        $rect = [iTextSharp.text.Rectangle]::new(

            0, 0, $width, $height 
        
        ) # define page size
        $doc = [iTextSharp.text.Document]::new($rect)
        [iTextSharp.text.Document]::Compress = $true
        $stream = [System.IO.MemoryStream]::new()
        $writer = [iTextSharp.text.pdf.PdfWriter]::GetInstance(
            
            $doc, $stream 
        
        )
        $doc.Open()
        $cb = $writer.DirectContent
        
        for ($pNum = 1 ; $pNum -le $reader.NumberOfPages ; $pNum++) {

            $page = $writer.GetImportedPage($reader, $pNum) # get content
            $pH = $rect.Height / $reader.GetPageSizeWithRotation($pNum).Height
            $pW = $rect.Width / $reader.GetPageSizeWithRotation($pNum).Width
            $cb.AddTemplate( $page, $pW, 0, 0, $pH, 0, 0 ) 
            $doc.NewPage()
        
        }
        $doc.Close() # must be closed before getting bytes from memory
        $reader.Close()
        $this.streamTofile($stream)
        
        return $this # chaining
    
    } # End
    [DocHelper] stamp([String] $text, [String] $customStamp) {
        
        $this.customStamp = $customStamp
        $this.stamp($text)

        return $this # chaining

    } # End
    [DocHelper] stamp([String] $text) {
        try {
        $reader = $this.initreader()
        [iTextSharp.text.pdf.PdfReader]::unethicalreading = $true
        $stream = [System.IO.MemoryStream]::new()
        $stamper = [iTextSharp.text.pdf.PdfStamper]::new( $reader , $stream )
        $stamper.AnnotationFlattening = ($true,$false)[!($text -eq 'flatten')]
        $layer=[iTextSharp.text.pdf.PdfLayer]::new("Watermark",$stamper.Writer)
        for ($p = 1 ; $p -le $reader.NumberOfPages ; $p++) {
            
            $cb = $stamper.GetOverContent($p)
            $rect = $reader.GetPageSize($p)
            $t = @{} ; switch ($text) {
           
                'angle'   { $t = @{ 
                        
                        "size" = if($rect.Height -le $rect.Width ) {$rect.Height/8} else {$rect.Width/8}
                        "text" = $this.customStamp
                        "x" = if($rect.Height -le $rect.Width -or $rect.Width -le 600 ) {$rect.Width/2} else {$rect.Height/2}
                        "y" = if($rect.Height -le $rect.Width -or $rect.Width -le 600 ) {$rect.Height/2} else {$rect.Width/2}
                        "angle" = -45
                        "opacity" = 0.2
                        "align" = [iTextSharp.text.pdf.PdfContentByte]::ALIGN_CENTER
                        "font" = [iTextSharp.text.pdf.BaseFont]::HELVETICA
                        
                 }}
                'str'  { $t = @{ 
                        
                        "size" = 10 
                        "text" = "{0}   Page {1} of {2}" -f 
                            
                            $this.customStamp, $p, $reader.NumberOfPages

                        "x" =  if($rect.Height -le $rect.Width -or $rect.Width -le 600 ) {$rect.Width} else {$rect.Height}
                              
                        "y" = if($rect.Height -le $rect.Width -or $rect.Width -le 600 ) {$rect.Height/300} else {$rect.Width/300}
                        "angle" = 0
                        "opacity" = 1
                        "align" = [iTextSharp.text.pdf.PdfContentByte]::ALIGN_RIGHT
                        "font" = [iTextSharp.text.pdf.BaseFont]::HELVETICA_BOLD
                        
                 }}
                'flatten'  { $t = @{ 
                        
                        "size" = 1 
                        "text" = ""
                        "x" = 1
                        "y" = 1
                        "angle" = 1
                        "opacity" = 0
                        "align" = [iTextSharp.text.pdf.PdfContentByte]::ALIGN_RIGHT
                        "font" = [iTextSharp.text.pdf.BaseFont]::HELVETICA
                        
                 }}
        
            }
            $cb.BeginLayer($layer) # next commands should be "bound" to this new layer
            $cb.SetFontAndSize(
                
                [iTextSharp.text.pdf.BaseFont]::CreateFont(
            
                    $t.font, 
                    [iTextSharp.text.pdf.BaseFont]::CP1252, 
                    [iTextSharp.text.pdf.BaseFont]::NOT_EMBEDDED
            
                ), 
                $t.size
            
            )            
            $gState = [iTextSharp.text.pdf.PdfGState]::new( )
            $gState.FillOpacity = $t.opacity
            $cb.SetGState($gState)
            $cb.SetColorFill([iTextSharp.text.BaseColor]::BLACK)
           #$cb.Rectangle(200,200,200,200) Fill area with color and opacity before placing text
           #$cb.Fill()
            $cb.BeginText()
            $cb.ShowTextAligned( $t.align , $t.text , $t.x , $t.y , $t.angle )
            $cb.EndText()
            $cb.EndLayer()

        }
        $stamper.Close()
        $reader.Close()
        $this.streamTofile($stream)
        } catch { [Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error") }
        return $this

    } # End
    hidden [DocHelper] normRotation([iTextSharp.text.pdf.PdfReader] $reader) {

        $reader = $this.initreader()
        [iTextSharp.text.pdf.PdfReader]::unethicalreading = $true
        $doc = [iTextSharp.text.Document]::new()
        [iTextSharp.text.Document]::Compress = $true
        $stream = [System.IO.MemoryStream]::new()
        $writer = [iTextSharp.text.pdf.PdfWriter]::GetInstance(
            
            $doc, $stream 
        
        )
        $doc.Open()
        $cb = $writer.DirectContent
        for ($pNum = 1 ; $pNum -le $reader.NumberOfPages ; $pNum++) {

            $pageSize = $reader.GetPageSizeWithRotation($pNum)
            $page = $writer.GetImportedPage($reader, $pNum)
            $pRotation = $pageSize.Rotation
            [iTextSharp.text.Rectangle] $newPageSize = $null
            $rotateFlip = [System.Drawing.RotateFlipType]
            switch($rotateFlip) { 

                $rotateFlip::RotateNoneFlipNone {
                
                    $newPageSize = [iTextSharp.text.Rectangle]::new($pageSize) ; break 
                
                }
                $rotateFlip::Rotate90FlipNone {
                    
                    $pRotation += 90
                    $newPageSize = [iTextSharp.text.Rectangle]::new($pageSize.Height, $pageSize.Width, $pRotation) ; break 
                
                }
                $rotateFlip::Rotate180FlipNone {
                    
                    $pRotation += 180
                    $newPageSize = [iTextSharp.text.Rectangle]::new($pageSize.Height, $pageSize.Width, $pRotation) ; break 
                
                }
                $rotateFlip::Rotate270FlipNone {
                    
                    $pRotation += 270
                    $newPageSize = [iTextSharp.text.Rectangle]::new($pageSize.Height, $pageSize.Width, $pRotation) ; break 
                
                }
            }
            $pRotation += 90
            $newPageSize = [iTextSharp.text.Rectangle]::new($pageSize.Height, $pageSize.Width, $pRotation)
            $doc.SetPageSize($newPageSize)
            $doc.NewPage()

            switch ($pRotation)
            {

                0 {   $cb.AddTemplate($page, 0, 0) ; break }
                90 {  $cb.AddTemplate($page, 0, -1, 1, 0, 0, $newPageSize.Height) ; break }
                180 { $cb.AddTemplate($page, -1, 0, 0, -1, $newPageSize.Width, $newPageSize.Height) ; break }
                270 { $cb.AddTemplate($page, 0, 1, -1, 0, $newPageSize.Width, 0) ; break }
                
            }

        }
        $doc.Close() # must be closed before getting bytes from memory
        $writer.Close()
        $reader.Close()
        $this.streamTofile($stream)


        <#
        iTextSharp.text.pdf.PdfContentByte cb = writer.DirectContent;
        iTextSharp.text.pdf.PdfImportedPage page;
        int rotation;
        int i = 0;
        while (i < pageCount)
        {
            i++;
            var pageSize = reader.GetPageSizeWithRotation(i);

            // Pull in the page from the reader
            page = writer.GetImportedPage(reader, i);

            // Get current page rotation in degrees
            rotation = pageSize.Rotation;

            // Default to the current page size
            iTextSharp.text.Rectangle newPageSize = null;

            // Apply our additional requested rotation (switch height and width as required)
            switch (rotateFlipType)
            {
                case RotateFlipType.RotateNoneFlipNone:
                    newPageSize = new iTextSharp.text.Rectangle(pageSize);
                    break;
                case RotateFlipType.Rotate90FlipNone:
                    rotation += 90;
                    newPageSize = new iTextSharp.text.Rectangle(pageSize.Height, pageSize.Width, rotation);
                    break;
                case RotateFlipType.Rotate180FlipNone:
                    rotation += 180;
                    newPageSize = new iTextSharp.text.Rectangle(pageSize.Width, pageSize.Height, rotation);
                    break;
                case RotateFlipType.Rotate270FlipNone:
                    rotation += 270;
                    newPageSize = new iTextSharp.text.Rectangle(pageSize.Height, pageSize.Width, rotation);
                    break;
            }

            // Cap rotation into the 0-359 range for subsequent check
            rotation %= 360;

            document.SetPageSize(newPageSize);
            document.NewPage();

            // based on the rotation write out the page dimensions
            switch (rotation)
            {
                case 0:
                    cb.AddTemplate(page, 0, 0);
                    break;
                case 90:
                    cb.AddTemplate(page, 0, -1f, 1f, 0, 0, newPageSize.Height);
                    break;
                case 180:
                    cb.AddTemplate(page, -1f, 0, 0, -1f, newPageSize.Width, newPageSize.Height);
                    break;
                case 270:
                    cb.AddTemplate(page, 0, 1f, -1f, 0, newPageSize.Width, 0);
                    break;
                default:
                    throw new System.Exception(string.Format("Unexpected rotation of {0} degrees", rotation));
                    break;
            }
        }
        #>
    
    return $this
    
    } # End
    [DocHelper] split() {
    
        $reader = $this.initreader()
        [iTextSharp.text.pdf.PdfReader]::unethicalreading = $true

        if ($reader.NumberOfPages -ne 1 ) {

            $newFolder = $this.inputFile.FullName + ' - Directory'
            
            if (-not (Test-Path $newFolder) ) {
              
                New-Item $newFolder -Type Directory     
                
                for($p = 1 ; $p -le $reader.NumberOfPages ; $p++) {

                    $newFileName = Join-Path -Path $newFolder -ChildPath (
                        
                        $this.inputFile.BaseName + '( ' + $p + ' )' +  $this.inputFile.Extension

                    )
                    $document = New-Object iTextSharp.text.Document
                    $fileStream = New-Object System.IO.FileStream($newFileName, [System.IO.FileMode]::Create)
                    $pdfCopy = New-Object iTextSharp.text.pdf.PdfCopy($document, $fileStream)
                    $document.Open()
                    $pdfCopy.AddPage( $pdfCopy.GetImportedPage($reader, $p) )

                    $pdfCopy.Dispose()
                    $fileStream.Dispose()
                    $document.Dispose() 
                    
                }
            }  

            $reader.Dispose()

            if ([Microsoft.VisualBasic.Interaction]::MsgBox("Do you wish to remove splitted files","Question,YesNo","Spliting") -like 'yes') {
                    
                try { Remove-Item $this.inputFile.FullName -ErrorAction Stop } catch { 
                        
                [Microsoft.VisualBasic.Interaction]::MsgBox(($_.Exception.Message).tostring(),"Critical","Error") }

            }

        }

        return $this
    
    } # End
    [DocHelper] wpsFrom([System.IO.FileSystemInfo] $excelFile) {
        
        $excel = $this.openXls($excelFile) # Sheet, Workbook, Workbooks, XlsObject
        $wpsArray = New-Object System.Collections.ArrayList
        $files = New-Object System.Collections.ArrayList  
        $lastRow = $excel[0].UsedRange.rows.count + 1 # lastrow of activesheet
        $excel[0].Range( # get every cell from 4 to 6 column each row(wps)

            $excel[0].Cells(1,4), $excel[0].Cells($lastRow,6)

        ).value2 | 
            
            foreach { if ($_ -match '\d{3}\.\w+\.?\w+' ) { [void]$wpsArray.Add($_) }}

        $wpsArray = $wpsArray | select -Unique | Sort-Object
        # pipeline find latest revision of files mached wpsArray elements and copy to folder
        $filename = $excelFile.Name -replace '.xls', '.pdf' 
        $filename = $filename -replace '_KK_', '_WPS_'
        $newFilename = Join-Path -Path  $excelFile.DirectoryName -ChildPath $filename
        $files = $wpsArray | ForEach-Object {
         
            Get-ChildItem -Path $this.wpsRegister -Recurse -Filter *$_*.pdf | 
            Sort-Object { [regex]::Replace( $_, '\d+',{ $args[0].Value.Padleft(20) } ) } |
            Select-Object -Last 1
        
        } # output to $files results of copying

        $this.mergeFilesTo($files, $newFilename)
        
        if ($wpsArray.Count -ne $files.Count) { # check if not all proccessed
            
            [Microsoft.VisualBasic.Interaction]::MsgBox(
                
                "Processed: " + $files.Count + " wps from total: " + $wpsArray.Count, "Critical", "Alert"
            ) 
        }
        #else { [Microsoft.VisualBasic.Interaction]::MsgBox("Done",'OKOnly,Information',"Wps from welding card") }
        
        $this.releaseXls($excel[0], $excel[1], $excel[2], $excel[3]) # Sheet, Workbook, Workbooks, XlsObject
            
        return $this # chaining

    } # End
    [DocHelper] splitByFormats() {

        $reader = $this.initreader()

        [iTextSharp.text.pdf.PdfReader]::unethicalreading = $true
            
        $rdFile = $this.inputFile
        $newFolder = New-Item ( Join-Path $rdFile.Directory -ChildPath 'FOR_PRINT' ) -ItemType Directory -Force
        $documents   = [System.Collections.ArrayList]::new()
        $fileStreams = [System.Collections.ArrayList]::new()
        $pdfCopies   = [System.Collections.ArrayList]::new()
            
        'A0','A1','A2','A3' | foreach { 
                
            $newFilename = "{0}\{1}_{2}" -f $newFolder, $_, $rdFile.Name
            $documents += New-Object iTextSharp.text.Document
            $fileStreams += New-Object System.IO.FileStream($newFilename, [System.IO.FileMode]::Create)

        }
        for($i = 0; $i -lt 4;$i++){ 

            $pdfCopies += New-Object iTextSharp.text.pdf.PdfCopy($documents[$i], $fileStreams[$i]) 
            $documents[$i].Open()

        }
        for($p = 1 ; $p -le $reader.NumberOfPages ; $p++) {
                
            $size = $reader.GetPageSizeWithRotation($p)
            $w = ($size.Height,$size.Width)[($size.Width -gt $size.Height)]
            $h = ($size.Width,$size.Height)[($size.Width -gt $size.Height)]

            if (($w-$h) -gt 800) {$pdfCopies[0].AddPage($pdfCopies[0].GetImportedPage($reader,$p))}
            if (($w-$h) -lt 800 -and ($w-$h) -gt 600) {$pdfCopies[1].AddPage($pdfCopies[1].GetImportedPage($reader,$p))}
            if (($w-$h) -lt 600 -and ($w-$h) -gt 400) {$pdfCopies[2].AddPage($pdfCopies[2].GetImportedPage($reader,$p))}
            if (($w-$h) -lt 400)        {$pdfCopies[3].AddPage($pdfCopies[3].GetImportedPage($reader,$p))}

        }
        for($i = 0; $i -lt 4;$i++){ 
                
            try{$pdfCopies[$i].Dispose();$fileStreams[$i].Dispose();$documents[$i].Dispose()}
            catch{$fileStreams[$i].Dispose();$documents[$i].Dispose();Remove-Item $fileStreams[$i].Name -Force}
                

        }
        $reader.Close()
        

        return $this

    }
    [DocHelper] mergeFilesTo([Object[]] $filesCollection, [String] $outputFile) {
    
        $doc = New-Object iTextSharp.text.Document
        $stream = [System.IO.MemoryStream]::new()
        $pdfCopy = New-Object iTextSharp.text.pdf.PdfCopy($doc, $stream)
        $doc.Open()

        $filesCollection | sort { [regex]::Replace($_, '\d+',{$args[0].Value.Padleft(20)})} | ForEach-Object {
 
            $reader = New-Object iTextSharp.text.pdf.PdfReader($_.FullName)
            [iTextSharp.text.pdf.PdfReader]::unethicalreading = $true 
            $pdfCopy.AddDocument($reader)
            $reader.Dispose() 
        
        }

        $pdfCopy.Dispose()
        $doc.Dispose()
        $this.streamTofile($stream, $outputFile)
    
        return $this # chaining

    } # End
    [DocHelper] excelToPdf([System.IO.FileSystemInfo] $excelFile) {
        
        $excel = $this.openXls($excelFile) # Sheet, Workbook, Workbooks, XlsObject
        $xlFixedFormat = 'Microsoft.Office.Interop.Excel.xlFixedFormatType' -as [type] 
        $filepath = Join-Path -Path $excelFile.DirectoryName -ChildPath ($excelFile.BaseName + '.pdf') 
        $excel[0].ExportAsFixedFormat($xlFixedFormat::xlTypePDF, $filepath)

        $this.releaseXls($excel[0], $excel[1], $excel[2], $excel[3]) # Sheet, Workbook, Workbooks, XlsObject

        return $this

    } # End
    [DocHelper] parseBOM(){
    
        $reader = $this.initreader()
        $text = New-Object System.Text.StringBuilder

        for($i = 1 ; $i -le $reader.NumberOfPages ; $i++) {
    
            [iTextSharp.text.pdf.parser.ITextExtractionStrategy]$strategy = New-object iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy
            [string]$currentText = [iTextSharp.text.pdf.parser.PdfTextExtractor]::GetTextFromPage($reader, $i, $strategy)
            [void]$text.Append($currentText) 
    
        }
        $this.asmNumber = $text.ToString() | Select-String '(\d{2}-\d+\/.*?(?=\s))\s+Koostu markeering' | foreach { $_.Matches.Groups[1].Value }
        # $text.ToString() | Select-String ('{0}\/.*?(?=\s+)' -f $asmNumber.ToString() ) -AllMatches | foreach { $_.Matches.Groups.Value }
        $fileName = $this.inputFile.BaseName | Select-String '(?i)\b[ap]\d{7}' | foreach { $_.Matches.Groups.Value }
        $filesInBom = $text.ToString() | Select-String '(?i)\b[ap]\d{7}\b' -AllMatches | foreach { $_.Matches.Groups.Value } | select -Unique
        $filesInBom | foreach { $this.bomFiles += ($_,$null)[$_ -eq $fileName] }
        
        if ($this.asmNumber.Length -ge 7) {
            $this.bomFiles = $this.bomFiles | select -Skip 1
            $this.asmFileName = $this.asmNumber -replace "(\b\d\b)",'0$1' -replace "\/",'_'
            $filter = $this.asmFileName + '_V*'
            try {
            $this.revNumber = ((gci -Path $this.archive -Include $filter -Recurse) | 
                                foreach { if ($_ -match 'V.\b\d+\b.pdf') {$_}} | 
                                sort { [regex]::Replace( $_, '_V.\b\d+\b',{ $args.value.Padleft(20) } ) } | 
                                select -Last 1 
                              ).basename.split('_V.') | select -Last 1 
            [int16]$this.revNumber += 1
            }
            catch { $this.revNumber = 0 }
            
        }

        $reader.Close()

        return $this

    }
    
    # HELPERS

    hidden [void] streamTofile([System.IO.MemoryStream] $stream, [String] $outputFile) {
        
        [byte[]] $content = $stream.ToArray()
        $streamTofile = [System.IO.FileStream]::new(
         
            $outputFile, 
            [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write
        
        ) # get instance for writing into same inputFile
        $streamTofile.Write($content, 0, $content.Length) # write bytes
        $streamTofile.Dispose() # remove all resources from instance
        $stream.Flush() # clear memory stream

    } # End
    hidden [void] streamTofile([System.IO.MemoryStream] $stream) {

        $this.streamTofile($stream, $this.inputFile.FullName)

    } # End
    hidden [array] openXls([System.IO.FileSystemInfo] $excelFile){
        
        $objExcel = New-Object -ComObject Excel.Application
        # (Exception from HRESULT: 0x80010105 (RPC_E_SERVERFAULT)
        # if excel object not shown (temp. workaround)
        $objExcel.Visible = $true
        $workBooks = $objExcel.Workbooks
        #$state = $false
        #$workBook = while($true){ 
        #     
        #    try{$workBook=$workBooks.Open($excelFile.FullName, 3);$state=$true;$workBook;break} 
        #    catch{Write-Warning $_.Exception}
        # 
        #} # with this block EXCEL.EXE clones (in try with exception) every run
        $workBook = $workBooks.Open( $excelFile.FullName,2,$true)
        $objExcel.Visible = $false
        $workBook.Saved = $true
        $sheet = $workBook.Worksheets.Item(1)

        return $sheet, $workBook, $workBooks, $objExcel

    } # End
    hidden [void] releaseXls($sheet, $workBook, $workBooks, $objExcel) {

        $workBook.Close($false)
        $objExcel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workBook)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workBooks)
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($objExcel)
        
        Remove-Variable -Name objExcel   

    } # End

} # DocHelper class
function Sort-Drawings {

    Param ([Object[]] $workfiles)

    if ($workFiles.Count -gt 0) {
                $asmDrawings = $workFiles | where { $_.Name -match '(?i)\ba\d{7}.*.pdf\b' }
                $prtDrawings = $workFiles | where { $_.Name -match '(?i)\bp\d{7}.*.pdf\b' }
                $restFiles   = $workFiles | where { $_.Name -notmatch '(?i)(\bp\d{7}.*.pdf\b)|(\ba\d{7}.*.pdf\b)' } | sort
            }

    [System.Collections.ArrayList] $boms = @()
    $asmDrawings | foreach { $hash = @{ BOM = ([DocHelper]::init($defaultLocation).read($_).parseBOM()).bomFiles
                                        FileName = $_.baseName }
                             $boms += New-Object psobject -Property $hash }
    $mainAssembly = $workFiles | where { $_.Name -match ($boms | where { $_.BOM -ne $null }).FileName }
    $bomsFiles = ($boms | where { $_.FileName -eq $mainAssembly.BaseName }).BOM

    $increment = 1
    if ( $mainAssembly.fullname -is [string] ) {
        Rename-Item -NewName ("_ {0} _{1}" -f $increment,$mainAssembly.name) -Path $mainAssembly.fullname
    }
    for ( $c=0 ; $c -lt $bomsFiles.count){
    
        $file = $workFiles | where { $_.basename -match $bomsFiles[$c] } 
        $c++
        $increment++
        if ( $file.fullname -is [string] ) {
            Rename-Item -NewName ("_ {0} _{1}" -f $increment,$file.name) -Path $file.FullName 
        }

    }
    $restFiles | foreach { 

        $increment++
        Rename-Item -NewName ("_ {0} _{1}" -f $increment,$_.name) -Path $_.fullname 

    }

}
function Copy-byHP {

    Param (
    
        [String] $BOM,
        [String] $dest

    )

    for ($i = 0;$i -lt $list.Count;$i++){ $files | where { $_.BaseName -match $list[$i] } | Copy-Item -Destination $dest }

}
function CollectDrawings {

    Param (
        
        [String] $prjNumber
    
    )

    $project   = ($prjNumber -split '/')[0]
    $prdNumber = ($prjNumber -split '/')[1]
    $asmNumber = ($prjNumber -split '/')[2]

    $prjFolder = gci 'w:\' | where { $_.Name -match $project }
    $asmFolder = gci -Path (Join-Path $prjFolder.fullname -ChildPath 'Tööjoonised') | where { $_.Name -match "{0}.+{1}.+{2}" -f $project, $prdNumber, $asmNumber }

}