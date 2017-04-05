[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$uncServer,
    [Parameter(Mandatory=$True,Position=2)]
    [string]$uncFullPath,
    [Parameter(Mandatory=$True,Position=3)]
    [string]$uncUser,
    [Parameter(Mandatory=$True,Position=4)]
    [string]$uncPass, # no point converting to secure string for my use?

    [Parameter(Mandatory=$True,Position=5)]
    [string]$zipDestination, 
    [Parameter(Mandatory=$True,Position=6)]
    [string]$zipSource,
    [Parameter(Mandatory=$True,Position=7)]
    [string]$zipPassword, # no point converting to secure string for my use?

    [Parameter(Mandatory=$True,Position=8)]
    [string]$zipBatchFile
)
<#
    Using PowerShell to get round windows scheduler issues with context/
        permissions, when using "user not logged in" context, and UNC paths
        when not part of a domain and without a domain user (and when
        I can't explicitly set the NTFS user permission rather than group 
        permission on the UNC target etc.)

        I found that even when running as a local admin, with batch log on,
        just using VBScript and Batch files I could not elegantly surmount
        these issues.  PowerShell has no such issue using (network) UNC paths.

        This UNC network path is for a HooToo travel router with a Samba share:
        not part of any domain or anything.

        In fact, I can just stick PowerShell in the middle of a VBScript &
        Batch file sandwich just to get 'round the permission issue - as
        I'm doing here. (Remember Execution policy set to bypass!)

    Useful PowerShell websites for my reference:
    parameters:     https://technet.microsoft.com/en-us/library/jj554301.aspx
    unc:            http://antonkallenberg.com/2013/04/20/powershell-unc-path-credentials
#>

# 
net use $uncServer $uncPass /USER:$uncUser
try 
{
    $q = [char]34
    $cmd = "cmd.exe /c --%"  # --% to stop zipPassword expansion if it contains special chars

    Invoke-Expression "$cmd $zipBatchFile $uncFullPath\$zipDestination $zipSource $q$zipPassword$q"


}
catch [System.Exception]
{
    
}
finally
{
    net use $uncServer /delete
}