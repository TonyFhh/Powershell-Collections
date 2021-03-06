# Sorter.ps1
# Intended to move ecompare results to designated assigned to folders
# How to expand to include the src/tgt folders?
    # Differences: missing assigned to: instead of leaving it, move to under stream



param(
[Parameter(Mandatory=$true)]
[ValidateScript({Test-Path $_ })][string]$result_path,
[ValidateScript({Test-Path $_ })][string]$tracker,
[switch]$base #works like option tag without arguments
)

function main () {

    # echo "results_path is $result_path"
    # echo "lookupfile is $tracker"
    
    # Python Portion
        # Ideally the file will be generated by pandas reading from tracker
        # The python will generate a csv file containing reportinfo
        # The python will generate cfg with destination path (eCompare_results\...)

    # Read the report data into stream, source, and assigned hashes.
    echo "Reading tracker information..."
    $tracker_data = python parsetracker.py $tracker
    if ( ! $? ) {
        exit 1 
    }
        

    if (Test-Path $tracker_data ) {
        
        $report_data = Import-Csv -Path $tracker_data -Delimiter ~
        $source_hash = @{}
        $stream_hash = @{}
        $assigned_hash = @{}
        foreach( $r in $report_data ) {
            $source_hash[$r.'Result Name']=$r.'Source'
            $stream_hash[$r.'Result Name']=$r.'Stream'
            $assigned_hash[$r.'Result Name']=$r.'Assigned to'
        }
    } else {
        Throw "Invalid tracker path `"${_}`", Perhaps you forgot to quote the path?"
    }
    
    $stream_values = $( $stream_hash.Keys | % { $stream_hash.Item($_) } ) #Hold all values of stream_hash for quick checks in main loop later
    
#    echo "source_hash is $( $source_hash | Out-String)" #prints full hash, #to retrieve individual key value, use $hash.$key
#    echo "stream_hash is $( $stream_hash | Out-String)"
#    echo "assigned_hash is $( $assigned_hash | Out-String)"


    #Check the hashes for emptiness
#    $assigned_hash.Count


    #Make the directories as needed
    # Expand functionality to include check of GMP/CCRS in result_path
#    $a_hashvals = $( $assigned_hash.Keys | % { $assigned_hash.Item($_) } ) #Reads the assigned_hash
    echo "Checking and creating needed directories..."
    $a_hashvals = % { 'Dev', 'BA', 'User' } 
    $stream_hash.Keys | %{ $stream_hash.Item($_) } | Get-Unique | Foreach-Object { 
        New-Item -Itemtype Directory -Force -Path $result_path\$_ | Out-Null
        $parent_dir = $_

        if ( ! $base ) {
            $a_hashvals | ForEach-Object {
                New-Item -Itemtype Directory -Force -Path $result_path\$parent_dir\$_ | Out-Null
            }
        }
    }
    
    echo "Moving files..."
    ls -path $result_path\* | Foreach-Object {
        # if iterated ls object is one of the stream folders, skip it
        if ( $stream_values -contains $_.Name ) {
            return #return is used instead of continue due to the weird design of each Foreach being a method
        }

        $full_path = $_
        $result_file_base = $_.BaseName
        $result_file = $_.Name
#        echo "result_file_base is $result_file_base, result_file is $result_file"
        # if there are no already existing results, move folder straight in
        <#if ( $assigned_hash.Item($result_file) -eq "Dev" ) {
            echo "result_file is Dev"
        } #>

        # Determine the destination path of the file
        if ( $base ) {
            # Check file for its stream and shift to associated folder
            if ( $stream_hash.ContainsKey($result_file_base) ) {
                $dest_path = "$result_path\$($stream_hash.Item($result_file_base))"
            } else {
                echo "Unable to determine stream for $result_file"
                return
            }
        } else { #Standard move for moving eCompare results
            # assigned destination paths depending on its 'assigned to' hash as per tracker
            switch($assigned_hash.Item($result_file_base)) {
                "BA" {
                    $dest_path = "$result_path\$($stream_hash.Item($result_file_base))\BA"
                    #$dest_path = "$result_path\$($source_hash.Item($result_file_base))\$($stream_hash.Item($result_file_base))\BA"
    #                echo "BA: destpath is $dest_path"
                }
                "Dev" {
                    $dest_path = "$result_path\$($stream_hash.Item($result_file_base))\Dev"
                    #$dest_path = "$result_path\$($source_hash.Item($result_file_base))\$($stream_hash.Item($result_file_base))\Dev"
    #                echo "Dev: destpath is $dest_path"
                }
                default {
                    # in event if source or stream of result_file cannot be determined
                    echo "Non Dev/BA assignment for $result_file, skipping"
                    return
                
                }
            } #end switch case (determine dest dir)
        }

        # Check result path for existing result file
        if ( [System.IO.File]::Exists("$dest_path\$result_file") ) {
            if ( (Get-ChildItem $dest_path\$result_file).LastWriteTime -ge (Get-ChildItem $result_path\$result_file).LastWriteTime ) {
                echo "found same or newer result in destination folder for $result_file, hence not moving"
                return
            } else {
#                echo "else: src file newer"
                if ( ! $base ) { # Handle results file
                    # Rename src path file, appending timestamps in name
                    $result_file_rename = "$($result_file_base)_$($full_path.LastWriteTime.ToString("ddMMyy_HHmm"))$($full_path.Extension)"
                    move_Result $result_path $result_file $dest_path $result_file_rename
                } else {
                    move_Result $result_path $result_file $dest_path $result_file
                }
            }
        } else {
#            echo "else: file not exist"
            move_Result $result_path $result_file $dest_path $result_file
        }


    } #End of ls loop
    
    rm $tracker_data
    
    # Check if $results_path is a single file or folder
    # $ExcelWB = new-object -comobject excel.application
    # if ( test-path $result_path -PathType Container ) {
       # echo "$result_path is folder"
        # ls -path $result_path\* -Recurse -Include *.xlsb,*.xlsx | Foreach-Object {
           # echo "starting $_"
            # file_check $_ #Wish i can make this work in background, but all the tedious stuff with Start-Job makes it unworthwhile
            # sleep 1
        # }
    # } elseif ( test-path $result_path -Include *.xlsx,*.xlsb ) {
       # echo "$result_path is file"
        # file_check $result_path
    # } else {
        # throw "Invalid path $result_path. Specify the path of a folder, xlsx or xlsb file"
    # }
    # $ExcelWB.quit()
    
    # Show-BalloonTip(3) -Title "orphans.ps1" -MessageType Info -Message "Script has checking file orphans."
    
    # [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelWB) | out-null #Release the object and suppress output (else it will print 0)
    # [System.GC]::Collect() #Forces immediate Garbage Collection for all objects
}

function move_Result($sourcedir, $sfile, $targetdir, $tfile) {
#    echo "result_file is $sourcedir, dest_path is $targetdir"
    try {
        mv "$sourcedir/$sfile" "$targetdir/$tfile" -Force
    } Catch {
        echo "Unable to move file $sfile, src path file or destination path file may already be in use by another process"
    }    
}


# Extra function to provide system tray notifications (copied here to improve portability to users)
function Show-BalloonTip {
    [cmdletbinding()]            
    param(            
     [parameter(Mandatory=$true)]            
     [string]$Title,            
     [ValidateSet("Info","Warning","Error")]             
     [string]$MessageType = "Info",            
     [parameter(Mandatory=$true)]            
     [string]$Message,            
     [string]$Duration=1000  
    )            

    [system.Reflection.Assembly]::LoadWithPartialName('System.Windows.Forms') | Out-Null            
    $balloon = New-Object System.Windows.Forms.NotifyIcon            
    $path = Get-Process -id $pid | Select-Object -ExpandProperty Path
    $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($MyInvocation.PSCommandPath ) #Get application name of running shell and extract its icon (to show in system tray)           
    $balloon.Icon = $icon            
    $balloon.BalloonTipIcon = $MessageType            
    $balloon.BalloonTipText = $Message            
    $balloon.BalloonTipTitle = $Title            
    $balloon.Visible = $true 
    
    $balloon.ShowBalloonTip($Duration)
    sleep $Duration
#    $balloon.Visible = $false
    $balloon.Dispose()
}

main