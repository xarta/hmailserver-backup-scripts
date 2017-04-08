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
    [string]$scriptPath 
)

net use $uncServer $uncPass /USER:$uncUser

try 
{
    $args = "ApprovedDeleteOldBackUps"
    $cmd = "cmd.exe /c cscript $scriptPath" 
    "$cmd $args"
    Invoke-Expression "$cmd $args"
}
catch [System.Exception]
{

    $_.Exception.Message
}
finally
{
    net use $uncServer /delete
}