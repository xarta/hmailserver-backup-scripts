[CmdletBinding()]
Param(
    [Parameter(Mandatory=$True,Position=1)]
    [string]$mysqlPath,
    [Parameter(Mandatory=$True,Position=2)]
    [string]$defaultsFile,
    [Parameter(Mandatory=$True,Position=3)]
    [string]$optionUser,
    [Parameter(Mandatory=$True,Position=4)]
    [string]$backupUser,
    [Parameter(Mandatory=$True,Position=5)]
    [string]$backupPassword,
    [Parameter(Mandatory=$True,Position=6)]
    [string]$database
)
try 
{

    $q = [char]34;
    $semicolon = ';';
    $optionDefault = "--defaults-extra-file";
    $options = "--verbose";
    $optionExecute = "--execute";

    # If the user already exists, it won't be overwritten. A non-fatal error will be raised
    $executeCreateUser = "CREATE USER '$backupUser'@'localhost' IDENTIFIED BY";
    $executeGrantUser = "GRANT EVENT, LOCK TABLES, SELECT, SHOW VIEW,  TRIGGER ON $database.* TO '$backupUser'@'localhost'";
    $executeFlushPrivileges = "FLUSH PRIVILEGES";

    [string]$mysql = $mysqlPath.trim();
    # could still use trim here more elgantly I think, but trying different things while learning
    [string]$mysql = $mysql -replace "^'", "";   # had to use special syntax to pass a string with spaces from vbscript to powershell
    [string]$mysql = $mysql -replace "'$", "";   # had to use special syntax to pass a string with spaces from vbscript to powershell
    $optionUser = $optionUser -replace "'", "";     # had to use special syntax for parameter with leading - ... confusion with powershell command
    [Array]$arguments1 = "$optionDefault=$q$defaultsFile$q", $options, $optionUser, "$optionExecute=$q$executeCreateUser $backupPassword$semicolon$q";
    [Array]$arguments2 = "$optionDefault=$q$defaultsFile$q", $options, $optionUser, "$optionExecute=$q$executeGrantUser$semicolon$q";
    [Array]$arguments3 = "$optionDefault=$q$defaultsFile$q", $options, $optionUser, "$optionExecute=$q$executeFlushPrivileges$semicolon$q";
    
    Write-Host $mysql $arguments1;
    & $mysql $arguments1;

    Write-Host $mysql $arguments2;
    & $mysql $arguments2;
    
    Write-Host $mysql $arguments3;
    & $mysql $arguments3;    
}
catch [System.Exception]
{
    $_.Exception.Message;
}
finally
{

}