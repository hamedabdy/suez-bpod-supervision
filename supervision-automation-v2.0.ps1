<#
    Documentation
    -------------

Required Files : 
                   + supervision-automation.ps1 (script)
                   + conf.psd1 (conf file)
                   + 2x .csv status correlation files
                   + IT4US incidents csv export file
                   + 2x BPOD (d-1 and d-2) csv export files

RUN :
           => Modify the config file (conf.psd1) with desired values
                -> csvFileIT4US = IT4US incidents csv export file
                -> csvFileBPOD  = BPOD incidents CSV export file concerning d-1
                -> csvFileBPOD2 = BPOD incidents CSV export file concerning d-2
                -> csvFileCorrespBpod and csvFileCorrespIt4us = concern the status correlations files
           => Choose the desired destination for your daily report, weekly report and finally the log file
           => Open the root directory of the script using your Shell and run the following command :
           => powershell.exe -noexit .\supervision-automation-vx.y.ps1


This script is used to automate the following task:
 - Verify if all of the transfered records from BPOD into IT4US (via Web service)
  have added properly or not.

This script consists of:
1- Reading the generated CSV files from BPOD and IT4US
2- Normalize the data (replace states with their correspondance)
3- Extract unique IDs
4- Compare IDs (correlation_id) from D-1
5- Generate statistics
6- Write these statistics into an Excel file
7- Generate weekly log of statistics
8- Time the script execution

The script is run using 'run()' which calls the 'main()' function with necesssay
params.

TO DO : 
1- DONE -- Logging
2- Auto download CSV files from BPOD (mail) and IT4US (PROD)
3- Test the script
4- Wrap script into an scheduled job inside a VM
5- Generate daily Chart
6- Generate weekly Charts and Diagrams

Voici les points pour la supervision :
1)	DONE -- Afficher les incidents qui sont en désynchro de statut.
2)	DONE -- Afficher les incidents en trop dans IT4US.
3)	DONE -- Afficher les incidents en trop dans BPOD.
4)	DONE -- Afficher les doublons de CorrelationID dans IT4US.
5)	DONE -- Afficher les doublons d'Incident Number dans BPOD.
6)	DONE -- Dans le fichier BPOD, corriger les incidents BPOD qui n’ont pas d’Incident Number IT4US.

#>



## Global Variables
## Loading Config data (should be in the same path as the script itself)
$conf = Import-LocalizedData -BaseDirectory $PSScriptRoot -FileName conf.psd1    
#$root = $conf.root
$root = "$PSScriptRoot\"
$logFile = $( $root + $conf.logfile )

# Trace Log 
$verbose = $false


###### Functions ######

# Date Comparison : $date1 is compared to $date2 and in case of equality :
# $true is returned; $false otherwise
function compareDate ( $date1, $date2 ) {
    $d1 = Get-Date -Date $date1 -UFormat "%Y-%m-%d"
    $d2 = Get-Date -Date $date2 -UFormat "%Y-%m-%d"
    if( ( $d1.GetType() -eq $d2.GetType() ) -and ( $d1.GetType() -eq "".GetType() ) )
    {
        if ( $d1 -eq $d2 ) { return $true }
        else { return $false }
    }
    else
    {
        throw "[compareDate] Please verify that date types are correct"
    }
}

## Select incidents by date
# Return : Array
function selectByDate ( [Array]$array, $date ) {
    [System.Collections.ArrayList]$tempA = @()
    for ( $i=0; $i -lt $array.number.Count; $i++ ) {
        if ( compareDate -date1 $array[$i].sys_created_on -date2 $date ) {
            $tempA.Add( $array[$i] )
        }
    }
    return $tempA
}

## Remove IT4US latest incidents (D day)
# Return IT4US array
function latestIncidents ( $array ) {
    $temp = @()
    $todayDate = ( get-date (Get-Date).AddDays(-1) -UFormat "%Y-%m-%d" )
    $en = $array.GetEnumerator()
    while ( $en.MoveNext() ) {
        $d2 = Get-Date ( $en.Current.sys_created_on ) -UFormat "%Y-%m-%d"
        if ( $d2 -lt $todayDate ) {
            $temp += $en.Current
        }
    }
    return $temp
}

# Comparing and replacing Initial state table 
# with the its own conversion table values
function convertState ( [System.Object]$csvArray, [array]$conversionTable ) {
    Write-Verbose "Input Count = $($csvArray.Count)" -Verbose:$verbose
    if ( $csvArray.state -ne $null ) {
        foreach( $e in $conversionTable ) {
            $i = $csvArray.state.IndexOf( $e.current )
            while ( $i -gt -1 ) {
                $csvArray[$i].state =  $csvArray[$i].state -replace( $e.current, $e.new )
                $i = $csvArray.state.IndexOf( $e.current )
            }
        }
    }
    if ( $csvArray.Etat -ne $null ) {
        foreach( $e in $conversionTable ) {
            $i = $csvArray.Etat.IndexOf( $e.current )
            if ( $e.current -eq $e.new ) {
                # Skip states that are the same on othe original and the conversion table
                continue
            }
            while ( $i -gt -1 ) {
                $csvArray[$i].Etat =  $csvArray[$i].Etat -replace( $csvArray[$i].Etat, $e.new )
                $i = $csvArray.Etat.IndexOf( $e.current )
            }
        }
    }
    Write-Verbose "-- Output Count = $($csvArray.Count)" -Verbose:$verbose
    return $csvArray
}

## Compare and count states from IT4US and BPOD
# Return : int counter
function matchStates ( $stateIt4us, $stateBPOD ) {
    $c = 0
    if( ( $stateIt4us.Count -ne 0 ) -and ( $stateBPOD.Count -ne 0 ) ) {
        $d = Compare-Object -IncludeEqual -ExcludeDifferent $stateIt4us $stateBPOD
    }
    $c = $d.Count
    return $c
}

# Counting number of each state
function countStates ( $stateArray, $newStateArray ) {
    $c = @( 0, 0, 0 )
    # Unique States
    $stateUnique = $newStateArray | select -Unique
    if ( $stateArray.state -ne $null ) {
        for ( $i=0; $i -lt $stateArray.Count; $i++ ) {
            # 0 => En-Attente, 1 => Clôturé, 2 => Résolu
            if ( $stateArray[$i].state -eq $stateUnique[0] ) { $c[0]++ }
            if ( $stateArray[$i].state -eq $stateUnique[1] ) { $c[1]++ }
            if ( $stateArray[$i].state -eq $stateUnique[2] ) { $c[2]++ }
        }
    }
    if ( $stateArray.Etat -ne $null ) {
        for ( $i=0; $i -lt $stateArray.Count; $i++ ) {
            # 0 => En-Attente, 1 => Clôturé, 2 => Résolu
            if ( $stateArray[$i].Etat -eq $stateUnique[0] ) { $c[0]++ }
            if ( $stateArray[$i].Etat -eq $stateUnique[1] ) { $c[1]++ }
            if ( $stateArray[$i].Etat -eq $stateUnique[2] ) { $c[2]++ }
        }
    }
    return $c
}

## Compare Correlation ids and find unmatching states!
function deltaStateByCorrId ( $array1, $array2 ) {
    Write-Host "Diff(State (IT4US), Etat (BPOD))..."
    $a = @(), @()
    Write-Host "Correlation_id --- State (IT4US) -- Etat (BPOD) -- BPOD ID" -ForegroundColor Cyan
    foreach ( $e in $array1 ) {
        $j = $array2."Numéro du dossier".IndexOf( $e.correlation_id )
        #if ( $e.correlation_id -eq $array2[$j]."Numéro du dossier" ) {
            if ( $e.state -ne $array2[$j].Etat ) {
                $a[0] += $e
                $a[1] += $array2[$j]
                Write-host "`t`t$($e.correlation_id) --- $($e.state) -- $($array2[$j].Etat) -- $($array2[$j]."Numéro du dossier")" -ForegroundColor Cyan
            }
        #}
    }
    return $a
}

function it4usDups ( $array1 ) {
    Write-Host "Finding duplicate Correlation-ids from IT4US..."
    $a = $array1 | Group-Object -Property correlation_id | Where-Object { $_.Count -gt 1 }
    return $a
}

function bpodDups ( $array2 ) {
    Write-Host "Finding duplicate Incident-numbers from BPOD..."
    $a = $array2 | Group-Object -Property ID_externe_servicenow | Where-Object { $_.Count -gt 1 }
    return $a
}

function diffIncNumber ( $array1, $array2 ) {
    Write-Host "Diff( Incident-number(IT4US), ID_externe_servicenow(BPOD) ) ..."
    $a = $array1 | Where-Object { $_.number -notin $array2.ID_externe_servicenow }
    return $a
}

function diffCorrId ( $array1, $array2 ) {
    Write-Host "Diff( Numero(BPOD), corr_id(IT4US) ) ..."
    $a = $array2 | Where-Object { $_."Numéro du dossier" -notin $array1.correlation_id }
    return $a
}

function fillBpodIncNbr ( $array1, $array2, $emptyIncObj ) {
    $a = foreach ( $e in $emptyIncObj.Group ) {
        $j = $array1.correlation_id.IndexOf( $e."Numéro du dossier" )
        $e.ID_externe_servicenow = $array1[$j].number
        $e."Référence externe" = $array1[$j].number
        $e
    }
    return $a
}

function deltaStats ( $array1, $array2 ) {
    Write-Host "`n"
    $d = @{}
    
    $d.desyncStates = deltaStateByCorrId $array1 $array2
    Write-Host "`t\->Writing desync states to file..." -ForegroundColor DarkYellow
    $d.desyncStates[0] | Export-Csv -NoTypeInformation -Encoding Default -Path ($root+"desync-IT4US.csv")
    $d.desyncStates[1] | Export-Csv -NoTypeInformation -Encoding Default -Path ($root+"desync-BPOD.csv")
    
    $d.it4usDups = it4usDups $array1
    Write-Host "`t\->Writing IT4US duplicate correlation-ids to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.it4usDups | Select-Object -Property Count, Name, @{name="number"; Expression={$_.Group.number} } | Out-String -Width 500 )
    $d.it4usDups | Select-Object -Property Count, Name, @{name="number"; Expression={$_.Group.number} } | Export-Csv $($root+"dups-IT4US.csv") -Encoding Default
    
    $d.bpodDups = bpodDups $array2
    Write-Host "`t\->Writing BPOD duplicate incident-numbers to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.bpodDups | Select-Object -Property Count, Name, @{name="Numéro du dossier"; Expression={$_.Group."Numéro du dossier"} } | Out-String -Width 500 )
    $d.bpodDups | Select-Object -Property Count, Name, @{name="Numéro du dossier"; Expression={$_.Group."Numéro du dossier"} } | Export-Csv $($root+"dups-BPOD.csv") -Encoding Default -NoTypeInformation
    
    $d.diffIncNumber = diffIncNumber $array1 $array2
    Write-Host "`t\->Writing IT4US incident-numbers that are NOT IN BPOD to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.diffIncNumber | Out-String )
    $d.diffIncNumber | Export-Csv $($root+"diff-IT4US.csv") -Encoding Default -NoTypeInformation
    
    $d.diffCorrId = diffCorrId $array1 $array2
    Write-Host "`t\->Writing BPOD correlation-ids that are NOT IN TI4US to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.diffCorrId | Out-String )
    $d.diffCorrId | Out-File $($root+"diff-BPOD.csv") -Encoding Default
    
    $d.fillBpodIncNbr = fillBpodIncNbr $array1 $array2 $d.bpodDups[0]
    Write-Host "`t\->Writing new BPOD file having no IT4US incident number to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.fillBpodIncNbr | Out-String )
    $d.fillBpodIncNbr | Export-Csv $($root+"new-BPOD.csv") -Encoding Default -NoTypeInformation
    return $d
}

<# Extracting new incidents from BPOD file
    Compare BPOD File1 with BPOD File2 (extract the difference by ID)
    finally extract the whole record by ID
    Return an array of new elements
#>
function getNewBpodIncidents ( $csvBpod1, $csvBpod2 ) {
    $b = @()
    $a = Compare-Object -ReferenceObject $csvBpod1."Numéro du dossier" -DifferenceObject $csvBpod2."Numéro du dossier"
    foreach ( $elem in $csvBpod1 ) {
        if ( $a -match $elem."Numéro du dossier" ) {
            $b += $elem
        }
    }
    return $b
}

function getBpodAttribs ( $bpodArray, $uniqueIds ) {
    $tempA = @()
    foreach ( $e in $bpodArray ) {
        if ( $uniqueIds -contains $e."Numéro du dossier" ) {
            $tempA += $e
        }
    }
    return $tempA
}

## Append into Excel File (weekly report)
function makeWeeklyExcelReport ( $xl, $weeklyReportOutFilePath, $dailyExcelReport ) {
    
    ## Start of week
    $s = get-date -hour 0 -minute 0 -second 0
    $s = $s.AddDays(-($s).DayOfWeek.value__)
    $s = $s.AddDays(1)
    $s = Get-Date $s -UFormat "%d-%m-%Y"

    ## Create new Excel file if inexistent
    if( -not ( Test-Path $($weeklyReportOutFilePath+".xlsx") ) ) {
        $wb = $xl.WorkBooks.Add()
        $xl.Cells.Item(1, 1).Font.Bold = $true
        $xl.Cells.Item(1, 1).Font.Size = 18
        $xl.Cells.Item(1, 1).Font.ThemeFont = 1
        $xl.Cells.Item(1, 1).Font.ThemeColor = 4
        $xl.Cells.Item(1, 1) = "Weekly Report For ${s}"
        $wb.SaveAs($weeklyReportOutFilePath)
        $wb.close()
        Write-Host "Preparing Weekly report.`nSleeping for 10 seconds."
        Start-Sleep -Seconds 10
    }
    ## Open Report and copy UsedRange
    $xlWorkBook = $xl.WorkBooks.Open( $dailyExcelReport, $null, $true )
    $xlWorkSheet = $xl.WorkSheets.Item(1)
    $xlWorkSheet.activate()
    $copyRange = $xlWorkSheet.UsedRange
    $copyRange.Copy()
    
    ## Open weekly report and append to the end
    $xlWorkBook2 = $xl.WorkBooks.Open($weeklyReportOutFilePath)
    $xlWorkSheet2 = $xl.WorkSheets.item(1)
    $xl.Columns.Item( 1 ).columnWidth = 40
    $lastRow = $xlWorkSheet2.UsedRange.rows.count + 5
    $range2 = $xl.Range( "A" + $lastRow ).Activate()
    $xlWorkSheet2.Paste()
    $xlWorkBook2.Save()
    $xlWorkBook2.close()
    $xlWorkBook.Close()
}

## Excel report file make
function makeExcelReport {

    Param(
            $xl,
            $outFilePath,
            $row=1,
            $date,
            $newIt4usIncidents,
            $globalMatchStates,
            $newIncidentsMatchStates,
            $it4usStateCounters,
            $bpodStateCounters,
            $it4usUnikIds,
            $bpodUnikIds,
            $writeToFile
          )

    # Writing results into Final Report Excel File
    # Create new workbook
    $wbReport = $xl.WorkBooks.add()

    $xl.Columns.Item( 1 ).columnWidth = 40

    # Bold heading text
    $xl.Rows.Item( $row ).Font.Bold = $true
    # Centre ( vertically ) heading
    $xl.Rows.Item( $row ).VerticalAlignment = -4108
    # Centre ( horizontally ) heading
    $xl.Rows.Item( $row ).HorizontalAlignment = -4108

    $xl.Cells.Item( $row, 1 ) = "SYNTHESE DU"
    $xl.Cells.Item( $row, 2 ) = $date

    $row++
    $row++
    # Bold heading text
    $xl.Rows.Item( $row ).Font.Bold = $true
    # Centre ( vertically ) heading
    $xl.Rows.Item( $row ).VerticalAlignment = -4108
    # Centre ( horizontally ) heading
    $xl.Rows.Item( $row ).HorizontalAlignment = -4108

    $xl.Cells.Item( $row, 1 ) = "NOMBRE D'INCIDENTS CREES LE"
    $xl.Cells.Item( $row, 2 ) = $date

    $row++
    # Bold heading text
    $xl.Cells.Item( $row, 1 ).Font.Bold = $true
    $xl.Cells.Item( $row, 1 ) = "NOMBRE D'INCIDENTS"
    $xl.Cells.Item( $row, 2 ) = "IT4US"
    $xl.Cells.Item( $row, 3 ) = "B'POD"
    $xl.Cells.Item( $row, 4 ) = "DELTA (Soustration)"
    $xl.Cells.Item( $row, 5 ) = "COMPARAISON ID"

    $row++

    $xl.Cells.Item( $row, 1 ) = "Incidents"
    $xl.Cells.Item( $row, 2 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 2 ) = $it4usUnikIds.Count
    $xl.Cells.Item( $row, 3 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 3 ) = $bpodUnikIds.Count
    $xl.Cells.Item( $row, 4 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 4 ) = [math]::Abs( $it4usUnikIds.Count - $bpodUnikIds.Count )
    $xl.Cells.Item( $row, 5 ) = [math]::Abs( $it4usUnikIds.Count - $newIncidentsMatchStates )

    $row++
    $xl.Cells.Item( $row, 1 ).Font.Bold = $true
    $xl.Cells.Item( $row, 1 ) = "VERIFICATION GLOBAL DES STATUS"

    $row++
    $xl.Cells.Item( $row, 1 ) = "En attente / Traitement"
    $xl.Cells.Item( $row, 2 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 2 ) = $it4usStateCounters[0]
    $xl.Cells.Item( $row, 3 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 3 ) = $bpodStateCounters[0]
    $xl.Cells.Item( $row, 4 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 4 ) = [math]::Abs( $it4usStateCounters[0] - $bpodStateCounters[0] )

    $row++
    $xl.Cells.Item( $row, 1 ) = "Résolu"
    $xl.Cells.Item( $row, 2 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 2 ) = $it4usStateCounters[2]
    $xl.Cells.Item( $row, 3 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 3 ) = $bpodStateCounters[2]
    $xl.Cells.Item( $row, 4 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 4 ) = [math]::Abs( $it4usStateCounters[2] - $bpodStateCounters[2] )

    $row++
    $xl.Cells.Item( $row, 1 ) = "Clôturé"
    $xl.Cells.Item( $row, 2 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 2 ) = $it4usStateCounters[1]
    $xl.Cells.Item( $row, 3 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 3 ) = $bpodStateCounters[1]
    $xl.Cells.Item( $row, 4 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 4 ) = [math]::Abs( $it4usStateCounters[1] - $bpodStateCounters[1] )

    $row++
    $xl.Cells.Item( $row, 1 ) = "Total"
    $xl.Cells.Item( $row, 2 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 2 ) = $stateIT4US.Count
    $xl.Cells.Item( $row, 3 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 3 ) = $stateBPOD.Count
    $xl.Cells.Item( $row, 4 ).Interior.ColorIndex = 43
    $xl.Cells.Item( $row, 4 ) = [Math]::Abs( $stateIT4US.Count - $stateBPOD.Count )

    $wbReport.SaveAs( $outFilePath)      
}

##### MAIN #####
## Magic starts here
function main {

    Param (
            $csvFileIT4US,
            $csvFileBPOD,
            $csvFileBpod2,
            $csvFileConversionBpod,
            $csvFileConversionIt4us,
            $outFilePath,
            $writeToFile,
            $weeklyReportOutFilePath,
            $writeToWeekReport
          )

    ## Normalize (remove signature (last 5 lines) from) BPOD CSV files
    $csvFileBPOD = $csvFileBPOD[0..($csvFileBPOD.Count-6)]
    $csvFileBpod2 = $csvFileBpod2[0..($csvFileBpod2.Count-6)]

    # Yesterday Date
    $dateYesterday = ( get-date ( get-date ).addDays( -1 ) -UFormat "%d/%m/%Y" )
    Write-Host "Yesterday Date => " $dateYesterday

    ## Convert states
    # IT4US
    Write-Host "`nConverting IT4US States... "
    $stateIT4US = convertState -csvArray $csvFileIT4US -conversionTable $csvFileConversionIt4us
    #$stateIT4US = latestIncidents -array $stateIT4US
    Write-Host "IT4US records count = " $stateIT4US.Count
    if ( $csvFileIT4US.Count -ne $stateIT4US.Count ) {
        write-host "Unmatching number of records in CSV files: original import Count = $($csvFileIT4US.Count)  After conversion Count = $($stateIT4US.Count)" -ForegroundColor Magenta
    }
    # BPOD
    Write-Host "`nConverting BPOD States..."
    $stateBPOD = convertState -csvArray $csvFileBPOD -conversionTable $csvFileConversionBpod
    Write-Host "B'POD records count = " $stateBPOD.Count

    ## Right Here compare states by their corr_ids
    deltaStats -array1 $stateIT4US -array2 $stateBPOD

    throw ""

    <# NOT NEEDED for the time being. Don't need statistical numbers
    ## BPOD : Extracting new incidents by applying delta on (D-1 and D-2) by IDs
    $newBpodIncidents = getNewBpodIncidents -csvBpod1 $csvFileBPOD -csvBpod2 $csvFileBpod2
    # Convert states
    $stateBpod2 = convertState -csvArray $newBpodIncidents -conversionTable $csvFileConversionBpod
    Write-Host "B'POD new incidents count = " $stateBpod2.Count
    #>
    ## Number of new incidents in IT4US
    [System.Collections.ArrayList]$newIt4usIncidents = selectByDate -array $csvFileIT4US -date $dateYesterday
    Write-Host "IT4US new incidents count = " $newIt4usIncidents.number.Count

    ## Count incidents by state
    Write-Host "`n`t`t`t`t`t`t  En-Attente   `tClôturé     `tRésolu"
    # IT4US
    $it4usStateCounters = countStates -stateArray $stateIT4US -newStateArray $csvFileConversionBpod.new
    Write-Host "IT4US each state counts = " $($it4usStateCounters -join "`t`t`t")
    # BPOD
    $bpodStateCounters = countStates -stateArray $stateBPOD -newStateArray $csvFileConversionBpod.new
    Write-Host "BPOD each state counts =  " $($bpodStateCounters -join "`t`t`t")

    ## Extracting Unique IDs
    # IT4US
    $it4usUniqueIds = ( $newIt4usIncidents.number | select -Unique )
    Write-Host "`nIT4US unique IDs count : "  $it4usUniqueIds.Count
    # BPOD
    $bpodUniqueIds = ( $newBpodIncidents."Numéro du dossier" | select -Unique )
    Write-Host "BPOD unique IDs count : "  $bpodUniqueIds.Count
    $bpodUnikIds = getBpodAttribs -bpodArray $stateBpod2 -uniqueIds $bpodUniqueIds

    ## Comparing tables from IT4US and BPOD and counting similarities
    # Global count
    #$globalMatchStates = matchStates -stateIt4us $stateIT4US.state -stateBPOD $stateBPOD.Etat
    # New incidents
    $newIncidentsMatchStates = matchStates -stateIt4us $it4usUniqueIds -stateBPOD $bpodUnikIds."ID_externe_servicenow"
    Write-Host "`nNew incidents BPOD IT4US Delta : " ([math]::Abs( $it4usUniqueIds.Count - $bpodUnikIds.Count ) )
    Write-Host "`nNew incidents (unique ids) (BPOD IT4US Delta) : " ( [math]::Abs( $it4usUnikIds.Count - $newIncidentsMatchStates ) )

    ## Writing data into Excel Report
    if( $writeToFile ) {
        # Create new Excel object
        $xl = New-Object -ComObject Excel.Application
        # Show interaction with Excel
        $xl.Visible = $false
        # Show Excel message alerts and dialogues
        $xl.DisplayAlerts = $false
        makeExcelReport -xl $xl -outFilePath $outFilePath `
        -date $dateYesterday -newIt4usIncidents $newIt4usIncidents.number.Count `
        -newIncidentsMatchStates $newIncidentsMatchStates `
        -it4usStateCounters $it4usStateCounters `
        -bpodStateCounters $bpodStateCounters -it4usUnikIds $it4usUniqueIds `
        -bpodUnikIds $bpodUniqueIds -writeToFile $writeToFile
        Write-Host "`nDone ==> Report written into : $outFilePath.xlsx" -ForegroundColor Gray

        if( $writeToWeekReport ) {
            makeWeeklyExcelReport -xl $xl `
            -weeklyReportOutFilePath $weeklyReportOutFilePath `
            -dailyExcelReport $outFilePath
            Write-Host "`nDone ==> Weekly report updated! $weeklyReportOutFilePath.xlsx" -ForegroundColor Gray
        }

        $xl.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
        Stop-Process -Name Excel -Force
        # Garbage Collection : Quitting and eliminating object
        Remove-Variable -Name xl
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers() 

    } 
    else {
        Write-Host "`n--writeToFile-- paramater is disabled!"
    }
}


##### RUN #####
## Magic is run from here
function run () {
    $csvFileIT4US = Import-Csv ( $root + $conf.csvFileIT4US ) -Delimiter ","
    $csvFileBPOD = Import-Csv ( $root + $conf.csvFileBPOD ) -Encoding Default -Delimiter ";"
    $csvFileBpod2 = Import-Csv ( $root + $conf.csvFileBpod2 ) -Encoding Default -Delimiter ";"
    $csvFileConversionBpod = Import-Csv ( $root + $conf.csvFileCorrespBpod ) -Encoding Default -Delimiter ";"
    $csvFileConversionIt4us = Import-Csv ( $root + $conf.csvFileCorrespIt4us ) -Encoding Default -Delimiter ";"

    # Excel Report output file Path
    $outFilePath = ( $root + $conf.outFilePath )
    $writeToFile = ( $conf.writeToFile )

    # Excel Week Report log
    $weeklyReportOutFilePath = ( $root + $conf.weeklyReportOutFilePath )
    $writeToWeekReport = ( $conf.writeToWeekReport )

    main -csvFileIT4US $csvFileIT4US -csvFileBPOD $csvFileBPOD `
    -csvFileBpod2 $csvFileBpod2 -csvFileConversionBpod $csvFileConversionBpod `
    -csvFileConversionIt4us $csvFileConversionIt4us -outFilePath $outFilePath `
    -writeToFile $writeToFile -weeklyReportOutFilePath $weeklyReportOutFilePath `    -writeToWeekReport $writeToWeekReport

    Write-Host "`n`n>>> Script is done. You may close the this window <<<"
}

## Logging console out stream into logfile
<#$ErrorActionPreference="SilentlyContinue"
Stop-Transcript | out-null
$ErrorActionPreference = "Continue"
Start-Transcript -path $logFile -Append
#>

## Running script and calculating elapsed time
$elapesed = Measure-Command { run }
Write-Host "Script runtime: `$($elapesed.Hours):`$($elapesed.Minutes):`$($elapesed.Seconds)" -ForegroundColor Yellow<## Stop loggingStop-Transcript#>