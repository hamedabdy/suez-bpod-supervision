<#
    Documentation
    -------------

Required Files : 
                   + supervision-automation.ps1 (script)
                   + conf.psd1 (conf file)
                   + 2x .csv status conversion files
                   + IT4US incidents csv export file
                   + 2x BPOD (d-1 and d-2) csv export files

RUN :
           => Modify the config file (conf.psd1) with desired values
                -> csvFileIT4US = IT4US incidents csv export file
                -> csvFileBPOD  = BPOD incidents CSV export file concerning d-1
                -> csvFileCorrespBpod and csvFileCorrespIt4us = concern the status correlations files


This script is used to automate the following task:
 - Verify if all of the transfered records from BPOD into IT4US (via Web service)
  have added properly or not.

This script consists of:
1- Reading the generated CSV files from BPOD and IT4US
2- Normalize the data (replace states with their correspondance)
4- Generate Stat files that show the differences between the 2 files
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
# Conversion States
$Global:it4usConvStates = @( 
    @{current='New';new='En attente'},
    @{current='Closed';new='Clôturé'},
    @{current='Resolved';new='Résolu'},
    @{current='To assign';new='En attente'},
    @{current='Pending';new='En attente'},
    @{current='Assigned';new='En attente'},
    @{current='In Progress';new='En attente'}
 )

$Global:bpodConvStates = @(
    @{current='Clôturé';new='Clôturé'},
    @{current='Résolu';new='Résolu'},
    @{current='Attente';new='En attente'},
    @{current='Traitement';new='En attente'},
    @{current='Suspendu';new='En attente'}
);

## Loading Config data (should be in the same path as the script itself)
$conf = Import-LocalizedData -BaseDirectory $PSScriptRoot -FileName conf.psd1    
#$root = $conf.root
$root = "$PSScriptRoot\"
$Global:dataRoot = "$PSScriptRoot\INC"
$logFile = "$PSScriptRoot\bpod-it4us-supervision.log"
$files = ( Get-ChildItem $Global:dataRoot | sort LastWriteTime | select -Last 2 )
$it4us = $files.name.IndexOf( ($files.name -like 'Supervision Flux BPOD*')[0] )
$bpod = $files.name.IndexOf( ($files.name -like 'Dossiers BPOD--IT4US*')[0] )
$Global:it4us = $files[$it4us].name
$Global:bpod = $files[$bpod].name
if ( -not ( (Test-Path $Global:dataRoot\$Global:it4us) -and (Test-Path $Global:dataRoot\$Global:bpod)) ) {
    Throw "Could not find the input CSV files"
}

# Trace Log
$verbose = $false


###### Functions ######

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

## Compare Correlation ids and find unmatching states!
function deltaStateByCorrId ( $array1, $array2 ) {
    Write-Host "`nDiff(State (IT4US), Etat (BPOD))..."
    $a = @(), @(), @()
    Write-Host "Correlation_id --- State (IT4US) -- Etat (BPOD) -- BPOD ID" -ForegroundColor Cyan
    foreach ( $e in $array1 ) {
        $j = $array2."Numéro du dossier".IndexOf( $e.correlation_id )
        if ( $e.correlation_id -contains $array2[$j]."Numéro du dossier" ) {
            if ( $e.state -notmatch $array2[$j].Etat ) {
            if ( ($e.state -imatch 'Résolu') -and ($array2[$j].Etat -imatch 'Clôturé') ) { Continue }
                $a[0] += $e
                $a[1] += $array2[$j]
                $a[2] += ( $e | Select-Object *, @{name="Numéro_BPOD"; Expression={$array2[$j]."Numéro du dossier"}}, @{name="Etat_BPOD"; Expression={$array2[$j].Etat}} )
                Write-host "`t`t$($e.correlation_id) --- $($e.state) -- $($array2[$j].Etat) -- $($array2[$j]."Numéro du dossier")" -ForegroundColor Cyan
            }
        }
    }
    return $a
}

function it4usDups ( $array1 ) {
    Write-Host "`nFinding duplicate Correlation-ids from IT4US..."
    $a = $array1 | Group-Object -Property correlation_id | Where-Object { $_.Count -gt 1 }
    return $a
}

function bpodDups ( $array2 ) {
    Write-Host "`nFinding duplicate Incident-numbers from BPOD..."
    $a = $array2 | Group-Object -Property ID_externe_servicenow | Where-Object { $_.Count -gt 1 }
    return $a
}

function diffIncNumber ( $array1, $array2 ) {
    Write-Host "`nDiff( incident-number(IT4US), ID_externe_servicenow(BPOD) )..."
    $a = $array1 | Where-Object { $_.number -notin $array2.ID_externe_servicenow }
    $b = $a | Where-Object { $_.correlation_id -notin $array2."Numéro du dossier" }
    return @{a=$a; b=$b}
}

function diffCorrId ( $array1, $array2 ) {
    Write-Host "`nDiff( Numéro(BPOD), corr_id(IT4US) )..."
    $a = $array2 | Where-Object { $_."Numéro du dossier" -notin $array1.correlation_id }
    return $a
}

function fillBpodIncNbr ( $array1, $array2, $emptyIncObj ) {
    Write-Host "`nFilling missing INC numbers in BPOD file..."
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
    $d.desyncStates[2] | Export-Csv -NoTypeInformation -Encoding Default -Path ($root+"desync-IT4US.csv")
    #$d.desyncStates[1] | Export-Csv -NoTypeInformation -Encoding Default -Path ($root+"desync-BPOD.csv")
    
    $d.it4usDups = it4usDups $array1
    Write-Host "`t\->Writing IT4US duplicate correlation-ids to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.it4usDups | Select-Object -Property Count, Name, @{name="number"; Expression={$_.Group.number} } | Out-String )
    $d.it4usDups | Select-Object -Property Count, Name, @{name="number"; Expression={$_.Group.number} } | Export-Csv $($root+"dups-IT4US.csv") -Encoding Default -NoTypeInformation
    
    $d.bpodDups = bpodDups $array2
    Write-Host "`t\->Writing BPOD duplicate incident-numbers to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.bpodDups | Select-Object -Property Count, Name, @{name="Numéro du dossier"; Expression={$_.Group."Numéro du dossier"} } | Out-String )
    $d.bpodDups | Select-Object -Property Count, Name, @{name="Numéro du dossier"; Expression={$_.Group."Numéro du dossier"} } | Export-Csv $($root+"dups-BPOD.csv") -Encoding Default -NoTypeInformation
    
    $d.diffIncNumber = diffIncNumber $array1 $array2
    Write-Host "`t\->Writing IT4US incident-numbers that are NOT IN BPOD to file..." -ForegroundColor DarkYellow
    #Write-Host ( $d.diffIncNumber[0] | Out-String )
    $d.diffIncNumber.a | Export-Csv $($root+"diff-IT4US-byNumber.csv") -Encoding Default -NoTypeInformation
    $d.diffIncNumber.b | Export-Csv $($root+"diff-IT4US-byCorrId.csv") -Encoding Default -NoTypeInformation
    
    $d.diffCorrId = diffCorrId $array1 $array2
    Write-Host "`t\->Writing BPOD correlation-ids that are NOT IN TI4US to file..." -ForegroundColor DarkYellow
    if( $d.diffCorrId.Count -eq 0 ) { Write-Host "Found None!" }
    else {
        Write-Host ( $d.diffCorrId | Out-String )
        $d.diffCorrId | Select-Object -Property "Numéro du dossier", "ID_externe_servicenow", "Etat"  | Export-Csv $($root+"diff-BPOD.csv") -Encoding Default -NoTypeInformation
    }
    
    $d.fillBpodIncNbr = fillBpodIncNbr $array1 $array2 $d.bpodDups[0]
    Write-Host "`t\->Writing new BPOD file having no IT4US incident number to file..." -ForegroundColor DarkYellow
    Write-Host ( $d.fillBpodIncNbr | Out-String )
    #$d.fillBpodIncNbr | Export-Csv $($root+"new-BPOD.csv") -Encoding Default -NoTypeInformation
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


##### MAIN #####
## Magic starts here
function main {

    Param (
            $csvFileIT4US,
            $csvFileBPOD,
            $csvFileConversionBpod,
            $csvFileConversionIt4us
          )

    ## Normalize (remove signature (last 5 lines) from) BPOD CSV files
    $csvFileBPOD = $csvFileBPOD[0..($csvFileBPOD.Count-6)]

    ## Convert states
    # IT4US
    Write-Host "`nConverting IT4US States... "
    $stateIT4US = convertState -csvArray $csvFileIT4US -conversionTable $csvFileConversionIt4us
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
}


##### RUN #####
## Magic is run from here
function run () {
    $csvFileIT4US = Import-Csv ( "$Global:dataRoot\$Global:it4us" ) -Delimiter ","
    $csvFileBPOD = Import-Csv ( "$Global:dataRoot\$Global:bpod" ) -Encoding Default -Delimiter ";"

    main -csvFileIT4US $csvFileIT4US -csvFileBPOD $csvFileBPOD `
    -csvFileConversionBpod $Global:bpodConvStates `
    -csvFileConversionIt4us $Global:it4usConvStates

    Write-Host "`n`n>>> Script is done. <<<"
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