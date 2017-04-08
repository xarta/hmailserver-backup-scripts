#net use $uncServer $uncPass /USER:$uncUser
net use \\XWIFI02 "1LlwSg0VCus9FVkP0OPZbOH9" /USER:admin



try 
{
$args = "ApprovedDeleteOldBackUps"

$cmd = "cmd.exe /c cscript G:\XARTA-SCRIPTS\XartaBackup.vbs" 

Invoke-Expression "$cmd $args"
#Invoke-Expression "$cmd $zipBatchFile $uncFullPath\$zipDestination $zipSource $q$zipPassword$q"


}
catch [System.Exception]
{
    
}
finally
{
    net use \\XWIFI02 /delete
    #net use $uncServer /delete
}