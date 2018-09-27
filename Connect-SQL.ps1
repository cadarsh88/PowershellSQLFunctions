<#
    .SYNOPSIS
        Connect to a SQL Server Database Engine and return the server object.

    .PARAMETER SQLServer
        String containing the host name of the SQL Server to connect to.

    .PARAMETER SQLInstanceName
        String containing the SQL Server Database Engine instance to connect to.

    .PARAMETER SetupCredential
        PSCredential object with the credentials to use to impersonate a user when connecting.
        If this is not provided then the current user will be used to connect to the SQL Server Database Engine instance.

    .PARAMETER LoginType
        If the SetupCredential is set, specify with this parameter, which type
        of credentials are set: Native SQL login or Windows user Login. Default
        value is 'WindowsUser'.
    
    .NOTES
        This Function depends on another function IMPORT-SQLPSMODULE
#>
function Connect-SQL
{
    [CmdletBinding()]
    param
    (
        [ValidateNotNull()]
        [System.String]
        $SQLServer = $env:COMPUTERNAME,

        [ValidateNotNull()]
        [System.String]
        $SQLInstanceName = 'MSSQLSERVER',

        [ValidateNotNull()]
        [System.Management.Automation.PSCredential]
        $SetupCredential,

        [Parameter()]
        [ValidateSet('WindowsUser', 'SqlLogin')]
        [System.String]
        $LoginType = 'WindowsUser'
    )

    Import-SQLPSModule

    if ($SQLInstanceName -eq 'MSSQLSERVER')
    {
        $databaseEngineInstance = $SQLServer
    }
    else
    {
        $databaseEngineInstance = "$SQLServer\$SQLInstanceName"
    }

    if ($SetupCredential)
    {
        $sql = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server

        if ($LoginType -eq 'SqlLogin')
        {
            $sql.ConnectionContext.LoginSecure = $false
            $sql.ConnectionContext.Login = $SetupCredential.Username
            $sql.ConnectionContext.SecurePassword = $SetupCredential.Password
        }

        if ($LoginType -eq 'WindowsUser')
        {
            $sql.ConnectionContext.ConnectAsUser = $true
            $sql.ConnectionContext.ConnectAsUserPassword = $SetupCredential.GetNetworkCredential().Password
            $sql.ConnectionContext.ConnectAsUserName = $SetupCredential.GetNetworkCredential().UserName
        }

        $sql.ConnectionContext.ServerInstance = $databaseEngineInstance
        $sql.ConnectionContext.Connect()
    }
    else
    {
        $sql = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $databaseEngineInstance
    }

    if ( $sql.Status -match '^Online$' )
    {
        Write-Verbose "Connected to database engine instance: $databaseEngineInstance" -Verbose
        return $sql
    }
    else
    {
        $errorMessage = $script:localizedData.FailedToConnectToDatabaseEngineInstance -f $databaseEngineInstance
        New-InvalidOperationException -Message $errorMessage
    }
}
