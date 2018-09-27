Function Test-AgentJob{
    [CmdletBinding()]
    param
    (
        [ValidateNotNull()]
        [System.String]
        $SQLServer = $env:COMPUTERNAME,

        [ValidateNotNull()]
        [System.String]
        $SQLInstanceName = 'MSSQLSERVER',

        [System.Management.Automation.PSCredential]
        $SetupCredential = [System.Management.Automation.PSCredential]::Empty,

        [Parameter()]
        [ValidateSet('WindowsUser', 'SqlLogin')]
        [System.String]
        $LoginType = 'WindowsUser',

        [ValidateNotNull()]
        [System.String]
        $AgentJobName,

        [ValidateNotNull()]
        [ValidateSet('Status', 'Start', 'Stop')]
        [System.String]
        $Action = "Status"
    )

    try{
        if($SetupCredential.UserName -eq $null){
            $sqlconnection = Connect-SQL -SQLServer $SQLServer -SQLInstanceName $SQLInstanceName -LoginType $LoginType
        } Else{
            $sqlconnection = Connect-SQL -SQLServer $SQLServer -SQLInstanceName $SQLInstanceName -SetupCredential $SetupCredential -LoginType $LoginType
        }
        
        $jobobj = $sqlconnection.JobServer.Jobs | Where-Object {$_.Name -eq $AgentJobName}
        if($Action -eq "Status"){
            $ActionStatus = $jobobj.CurrentRunStatus
            Write-Verbose "Current Job Status: $ActionStatus" -Verbose
        } elseif ($Action -eq "Start") {
            $jobobj.Start()
            $ActionStatus = "Started"
            Write-Verbose "Job $AgentJobName started" -Verbose
        } elseif($Action -eq "Stop") {
            $jobobj.Stop()
            $ActionStatus = "Stopped"
            Write-Verbose "Job $AgentJobName stopped" -Verbose
        }
        return New-Object PSObject -Property @{
            Status=$ActionStatus
            Error=$null
		}
    } Catch {
        $Exception= $_.Exception.Message	
        Write-Host "Exception Occured: $Exception"
        return New-Object PSObject -Property @{
            Status=$null
            Error=$Exception
		}
    }
}
