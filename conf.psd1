﻿## Make sure to have a uniform CSV delimiter (";" "," etc) 
## in all your CSV files

@{
    root = "C:\Users\VC5622\Desktop\script\"
    csvFileIT4US = "INC\Supervision Flux BPOD - Incidents IT4US.csv"
    csvFileBPOD = "INC\Dossiers BPOD--IT4US_31-01-2018 08-00-01.csv"
    csvFileBpod2 = "INC\Dossiers BPOD--IT4US_30-01-2018 08-00-05.csv"
    csvFileCorrespBpod = "correspondance-bpod.csv"
    csvFileCorrespIt4us = "correspondance-it4us.csv"

    # Excel Report output file Path
    outFilePath = "Report"
    writeToFile = $false
    weeklyReportOutFilePath = "WeeklyReport"
    writeToWeekReport = $false

    # Logging
    logfile = "bpod-it4us-supervision.log"
}