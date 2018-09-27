<#
    .SYNOPSIS
        Imports the module SQLPS in a standardized way.

    .PARAMETER Force
        Forces the removal of the previous SQL module, to load the same or newer
        version fresh.
        This is meant to make sure the newest version is used, with the latest
        assemblies.

#>
function Import-SQLPSModule
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [Switch]
        $Force
    )

    if ($Force.IsPresent)
    {
        Write-Verbose "Force Removal of Module" -Verbose
        Remove-Module -Name @('SqlServer','SQLPS','SQLASCmdlets') -Force -ErrorAction SilentlyContinue
    }

    <#
        Check if either of the modules are already loaded into the session.
        Prefer to use the first one (in order found).
        NOTE: There should actually only be either SqlServer or SQLPS loaded,
        otherwise there can be problems with wrong assemblies being loaded.
    #>
    $loadedModuleName = (Get-Module -Name @('SqlServer', 'SQLPS') | Select-Object -First 1).Name
    if ($loadedModuleName)
    {
        Write-Verbose "SQL Powershell Module already imported" -Verbose
        return
    }

    $availableModuleName = $null

    # Get the newest SqlServer module if more than one exist
    $availableModule = Get-Module -FullyQualifiedName 'SqlServer' -ListAvailable |
        Sort-Object -Property 'Version' -Descending |
        Select-Object -First 1 -Property Name, Path, Version

    if ($availableModule)
    {
        $availableModuleName = $availableModule.Name
        Write-Verbose "SqlServer module found" -Verbose
    }
    else
    {
        Write-Verbose "SqlServer module not found, trying SQLPS module" -Verbose

        <#
            After installing SQL Server the current PowerShell session doesn't know about the new path
            that was added for the SQLPS module.
            This reloads PowerShell session environment variable PSModulePath to make sure it contains
            all paths.
        #>
        $env:PSModulePath = [System.Environment]::GetEnvironmentVariable('PSModulePath', 'Machine')

        <#
            Get the newest SQLPS module if more than one exist.
        #>
        $availableModule = Get-Module -FullyQualifiedName 'SQLPS' -ListAvailable |
            Select-Object -Property Name, Path, @{
                Name = 'Version'
                Expression = {
                    # Parse the build version number '120', '130' from the Path.
                    (Select-String -InputObject $_.Path -Pattern '\\([0-9]{3})\\' -List).Matches.Groups[1].Value
                }
            } |
            Sort-Object -Property 'Version' -Descending |
            Select-Object -First 1

        if ($availableModule)
        {
            # This sets $availableModuleName to the Path of the module to be loaded.
            $availableModuleName = $availableModule.Path
        }
    }

    if ($availableModuleName)
    {
        try
        {
            Write-Debug "Debug message, pushing location"
            Push-Location

            <#
                SQLPS has unapproved verbs, disable checking to ignore Warnings.
                Suppressing verbose so all cmdlet is not listed.
            #>
            Import-Module -Name $availableModuleName -DisableNameChecking -Verbose:$False -Force:$Force -ErrorAction Stop

            Write-Verbose "Imported SQLPS Powershell Module" -Verbose
        }
        catch
        {
            $errorMessage = "Failed to import powershell SQL Module"
            Write-Host "Invalid Operation, Error: $errorMessage" -ForegroundColor Red
        }
        finally
        {
            Write-Debug "Debug message, popping location"
            Pop-Location
        }
    }
    else
    {
        $errorMessage = "Powershell SQLPS Module not found."
        Write-Host "Error: $errormessage" -ForegroundColor Red
    }
}
