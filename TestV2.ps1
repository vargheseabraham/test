
    #--------------------------------------------------------------------------------------------------------------#
    #------------------------------------------------ INITIALIZATION ----------------------------------------------# 
    #--------------------------------------------------------------------------------------------------------------#
   
    $sAMAccountName = "varghese.a@hcl.com"
    $hostName = "LP-5CD625589M"
    $generatedDate = (Get-Date).ToString("dd-MM-yyyy")
    
    #Collector Configurations
    $rootFolder = "$env:USERPROFILE\analyzeTest\"
    $basePath = "$env:USERPROFILE\analyzeTest\"+$generatedDate+"\"
    $reportFolder = "$env:USERPROFILE\analyzeTest\automatedReports\"
    $backupFolder = "$env:USERPROFILE\analyzeTest\Backup\"
    $uploadedFolder = "$env:USERPROFILE\analyzeTest\uploadedReports\"

    $applReportPath = $basePath+"$($env:COMPUTERNAME)_ApplicationUsage.csv"
    $browserReportPath = $basePath+"$($env:COMPUTERNAME)_browserUsage.csv"
    $meetingsReportPath = $basePath+"$($env:COMPUTERNAME)_scheduledMeetings.csv"  
  

    #Analyzer Configurations
    $currentDate = (Get-Date).ToString("dd-MM-yyyy")
    $folderPaths = (Get-ChildItem $rootFolder -Force).FullName
    $nonProductiveURL = 'https://www.youtube.com','https://www.facebook.com'

    function Get-UriSchemeAndAuthority
    {
        param(
            [string]$InputString
        )

        $Uri = $InputString -as [uri]
        if($Uri){
            $FullUri = $Uri.AbsoluteUri
            $Path = $Uri.PathAndQuery

            $SlashIndex = $FullUri.Length - $Path.Length

            return $FullUri.Substring(0,$SlashIndex)
        } else {
            throw "Malformed URI"
        }
    }


    #--------------------------------------------------------------------------------------------------------------#
    #------------------------------------------------- DATA ANALYZER PART------------------------------------------# 
    #--------------------------------------------------------------------------------------------------------------#

    foreach($fpath in $folderPaths)
    {
        $fname = $fpath.split('\')[4]

        if(($fname -like '*-*-*') -and ($currentDate -ne $fname))
        {            
            $timeObj = @()
            $finalObj = @()

            $reportPath = "$env:USERPROFILE\analyzeTest\automatedReports\automatedReport_$fname.csv"
            $backupPath = "$env:USERPROFILE\analyzeTest\Backup\$fname"
            $appDataPath = "$env:USERPROFILE\analyzeTest\$fname\$hostName"+"_ApplicationUsage.csv"
            $browserUsagePath = "$env:USERPROFILE\analyzeTest\$fname\$hostName"+"_browserUsage.csv"
            $meetingInfoPath = "$env:USERPROFILE\analyzeTest\$fname\$hostName"+"_scheduledMeetings.csv"

            $appData = import-csv $appDataPath
            $browserUsage = Import-Csv $browserUsagePath
            $meetingInfo = Import-Csv $meetingInfoPath
    
            $appData = $appData | Sort-Object -Property Date                

            for($i=0;$i -lt $appData.count;$i++)
            {   
                if($i -ne $appData.count-1)
                {
                    $start = [datetime]::ParseExact($appData[$i].Date,'dd-MM-yyyy HH:mm:ss', $null)
                    $end = [datetime]::ParseExact($appData[$i+1].Date,'dd-MM-yyyy HH:mm:ss', $null)
                    $timeDiff = NEW-TIMESPAN –Start $start -End $end 

                    #$timeDiff = NEW-TIMESPAN –Start $appData[$i].Date -End $appData[$i+1].Date      
                    $timeExec = "{0:D2}:{1:D2}:{2:D2}" -f $timeDiff.Hours, $timeDiff.Minutes, $timeDiff.Seconds
     
                    $timeObj += [pscustomobject]@{
                        ApplicationName = $appData[$i].CurrentApplication # Application Name
                        Date = $appData[$i].Date
                        ExecutionTime = $timeExec # Execution Time
                    }
                }
            }
        
            #Remove entries from "App & Browser Usage" while meetings scheduled and total meeting time calculation
            $appInfo = $timeObj
            $browserInfo = $browserUsage | Sort-Object -Property Date

            $validMeetings = $meetingInfo | ?{($_.Subject -notlike '*cancel*' -and ($_.Subject -ne ""))}
        
            [timespan]$temp = [timespan]'00:00:00'
            foreach($meeting in $validMeetings)
            {
                #$meetingStart = [datetime]::ParseExact($meeting.StartTime,'dd-MM-yyyy HH:mm:ss', $null)
                #$meetingEnd = [datetime]::ParseExact($meeting.EndTime,'dd-MM-yyyy HH:mm:ss', $null)

                $meetingTimeCalculation = New-TimeSpan -Start $meeting.StartTime -End $meeting.EndTime
                $meetingTime = [String]$meetingTimeCalculation.Hours+':'+[String]$meetingTimeCalculation.Minutes+':'+[String]$meetingTimeCalculation.Seconds
                [timespan]$temp += [timespan]$meetingTime
                $appInfo = $appInfo | ?{!(($_.Date -ge $startTime) -and ($_.Date -le $endTime))} 
                $browserInfo = $browserInfo | ?{!(($_.Date -ge $startTime) -and ($_.Date -le $endTime))}                    
            }
            $meetingTimeFormat = "{0:D2}:{1:D2}:{2:D2}" -f $temp.Hours, $temp.Minutes, $temp.Seconds
                        
            $browserObj = @()
            $browserInfo | ForEach-Object{
                if(($_.URL -ne "") -and ($_.Date -ne ""))
                {
                    $baseUri = Get-UriSchemeAndAuthority $_.URL
                    $browserDate = $_.Date     
                    $browserObj += [pscustomobject]@{
                        baseUri = $baseUri 
                        Date = $browserDate
                    }
                }
            }
        
            $browserFiltered = $browserObj | ?{$nonProductiveURL -contains $_.baseUri} | Select Date -Unique | Group-Object Date
        
            # Initialization [Productive Apps]
            [timespan]$temp = [timespan]'00:00:00'
            $appsProductive = $appInfo | ?{$_.ApplicationName -ne "System Idle"}|%{$obj = $_.ExecutionTime; [timespan]$temp += [timespan]$obj}
            $appsProductiveTime = "{0:D2}:{1:D2}:{2:D2}" -f $temp.Hours, $temp.Minutes, $temp.Seconds

            # Initialization [Non-Productive Apps]
            [timespan]$temp = [timespan]'00:00:00'
            $appsNonProductive = $appInfo | ?{$_.ApplicationName -eq "System Idle"}|%{$obj = $_.ExecutionTime; [timespan]$temp += [timespan]$obj}
            $appsNonProductiveTime = "{0:D2}:{1:D2}:{2:D2}" -f $temp.Hours, $temp.Minutes, $temp.Seconds

            # Initialization [Non-Productive Browser]
            $browserNonProductive = New-TimeSpan -Minutes $browserFiltered.Count
            $browserNonProductiveTime = "{0:D2}:{1:D2}:{2:D2}" -f $browserNonProductive.Hours, $browserNonProductive.Minutes, $browserNonProductive.Seconds
            
            # Productive Apps - Non-Productive Browser
            [timespan]$temp = [timespan]'00:00:00'
            [timespan]$temp = [timespan]$appsProductiveTime - [timespan]$browserNonProductiveTime
            $productiveTimeInitial = "{0:D2}:{1:D2}:{2:D2}" -f $temp.Hours, $temp.Minutes, $temp.Seconds
            if($productiveTimeInitial -like '*-*-*')
            {
                $productiveTimeInitial = "00:00:00"
            }

            # Productive Apps - Non-Productive Browser + Scheduled Meetings
            [timespan]$temp = [timespan]'00:00:00'
            [timespan]$temp = [timespan]$productiveTimeInitial + [timespan]$meetingTimeFormat
            $productiveTimeFormat = "{0:D2}:{1:D2}:{2:D2}" -f $temp.Hours, $temp.Minutes, $temp.Seconds
            if($productiveTimeFormat -like '*-*-*')
            {
                $productiveTimeFormat = "00:00:00"
            } 
                 
            # Non-Productive Apps + Non-Productive Browser
            [timespan]$temp = [timespan]'00:00:00'
            [timespan]$temp = [timespan]$appsNonProductiveTime + [timespan]$browserNonProductiveTime
            $nonProductiveTimeFormat = "{0:D2}:{1:D2}:{2:D2}" -f $temp.Hours, $temp.Minutes, $temp.Seconds

            $finalObj += [pscustomobject]@{
                user_name = $env:USERNAME
                machine_name = $hostName
                productive_time = $productiveTimeFormat
                nonproductive_time = $nonProductiveTimeFormat
                report_date = $fname
            }

            $finalObj | Export-Csv $reportPath -NoTypeInformation    
            Move-Item -Path $fpath -Destination $backupPath -Force
        }
    }
