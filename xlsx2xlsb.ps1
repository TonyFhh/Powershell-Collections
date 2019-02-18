#xlsx to xlsb converter
# THis script will convert xlsx files within a folder into xlsb. Xlsb files take less space and open faster if users don't use data analysis tools.
# The new xlsb file will inherit the file times of the initial xlsx file

# Process Parameters
param(
#[Parameter(Mandatory=$true)] #This sets the parameter to be mandatory, if not provided, powershell will ask for it.
[string]$path
)

function main () {

#    Import-Module "D:\balloontip.ps1"
    . D:\balloontip.ps1 #Import the function from the script so i don't have to rewrite

    if ( [string]::IsNullOrEmpty($path) ) { #If string is empty
#        echo "path is empty!"
        $path=Read-FolderBrowserDialog -InitialDirectory "D:\Users\tmph72\Documents\Work eCompare"
#        echo "new path is $path"
    } else {
#        echo "path is $path"
        if (!(test-path -Path "$path") ) {
            echo "Invalid path `"$path`". Perhaps you forgot to quote the path?"
            exit
        }
    }
    
    $converted=0
    $total=0
    
    $ExcelWB = new-object -comobject excel.application
    foreach($file in (Get-ChildItem -path "$path" -recurse -Include *.xlsx)) {
        $total++
        $newname = $file.FullName -replace '\.xlsx$', '.xlsb'
        if ( [System.IO.File]::Exists($newname) ) {
            if ( (Get-ChildItem $newname).LastWriteTime -ge (Get-ChildItem $file.FullName).LastWriteTime ) {
            echo "Found existing $($file.BaseName).xlsb file, skipping"
            Try {
                rm $file -ErrorAction Stop
            } Catch {
                echo "Unable to delete file $($file.Basename), it could in use by another process"
            }
            continue
            } else {
            echo "Removing $($file.BaseName).xlsb"
            rm $newname
            }
        }
        
        Try {
            $Workbook = $ExcelWB.Workbooks.Open($file.FullName)
        } Catch {
            echo "Unable to open file $($file.BaseName), skipping"
#            $Workbook.Close($false)
            continue
        }
        $Workbook.SaveAs($newname,50) #Save as xlsb
        $Workbook.Close($false)
        
        (Get-ChildItem $newname).CreationTime = (Get-ChildItem $file.FullName).CreationTime
        (Get-ChildItem $newname).LastWriteTime = (Get-ChildItem $file.FullName).LastWriteTime
        (Get-ChildItem $newname).LastAccessTime  = (Get-ChildItem $file.FullName).LastAccessTime
        $converted++
#        echo "Converted"
        Try {
            rm $file -ErrorAction Stop
        } Catch {
            echo "Unable to delete file $($file.Basename), it could in use by another process"
        }
    }
    $ExcelWB.quit()
    
    
#    if ( $total -ge 5 ) {
    Show-BalloonTip(3) -Title "xlsx2xlsb.ps1" -MessageType Info -Message "Script has finished converting files."
#    }
    echo "Converted $converted out of $total xlsx files to xlsb"
#    [System.GC]::Collect()
#    [System.GC]::WaitForPendingFinalizers()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB) | out-null #Release the object and suppress output (else it will print 0)
    #    Remove-Variable ExcelWB
    [System.GC]::Collect() #Forces immediate Garbage Collection for all objects

}

function Read-FolderBrowserDialog([string]$InitialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    $OpenFileDialog = New-Object System.Windows.Forms.FolderBrowserDialog #Opens the gui interface for selecting folder.
    $OpenFileDialog.SelectedPath = $InitialDirectory
    if ( $OpenFileDialog.ShowDialog() -eq "OK" ) { #Why does writing it this way work and not directly selected path?
        $folder += $OpenFileDialog.SelectedPath
    } else {
        exit
    }
    return $folder
}

main