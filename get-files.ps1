class GetFiles{

    [int16] $count
    [System.Array] $baseNames
    [string] $directory
    [System.Array] $fullPaths
    [String] $title

    hidden Getfiles() {}
    static [GetFiles] init() { 

        return [GetFiles]::New() 
    
    } # Named Constructor
    [GetFiles] setTitle($title) {
    
        $this.title = $title
            
    return $this
    }
    [GetFiles] Search($filter) {
        
        ($ofd = New-Object System.Windows.Forms.OpenFileDialog -Property @{
            
            Filter = switch($filter){
            
                 'pdf' {"PDF Files|*.pdf"}
                 'xls' {"XLS Files|*.xls;*.xlsx"}
            }
            Multiselect = switch($filter){
                
                'pdf' { $true }
                'xls' { $false }

            }
            Title = $this.title

        }).ShowDialog()

        $this.directory = Split-Path -Path $ofd.FileName
        $this.count = $ofd.FileNames.Count

        if ($this.count -gt 1) {

            $this.baseNames = (Split-Path -Path $ofd.FileNames -Leaf).foreach({ $_ + "`n" })
            $this.fullPaths = $ofd.FileNames 
            
        }
        elseif ($this.count -eq 1){

            $this.baseNames = Split-Path -Path $ofd.FileName -Leaf -Resolve
            $this.fullPaths = $ofd.FileName 

        }

        return $this

    }
}