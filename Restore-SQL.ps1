#Calvin Seok 011553098

#import the sqlserver module to set up the server
try {
    if (Get-Module -Name sqlps) { Remove-Module sqlps }
    Import-Module -Name SqlServer
# Set up variables 
    $Instance = "SRV19-PRIMARY\SQLEXPRESS"
    $database = "ClientDB"
    $schema = "dbo"
    $table = "Client_A_Contacts"
# if database exists, delete it 
    $serverobject = New-Object -TypeName Microsoft.SqlServer.Management.Smo.Server -ArgumentList $Instance
    $databaseObject = Get-SqlDatabase -ServerInstance $Instance -Name $database -ErrorAction SilentlyContinue
    if($databaseObject) {
        Write-Host -ForegroundColor Green "(SQL): $($database) Detected. Deleting."

        $serverobject.KillAllProcesses($database)

        $databaseObject.UserAccess = "Single"

        $databaseObject.Drop()

        Write-Host -ForegroundColor Green "(SQL): $($database) Successfully Deleted"
    }
    else {
        Write-Host -ForegroundColor Green "(SQL): $($database) Not Found"
    }
#open csv file and import data from it
    $NewClients = Import-CSV -Path $PSScriptRoot\NewClientData.csv
#create database
    Write-Host -ForegroundColor Green "(SQL): Constructing Database -> Inserting Data"
    $NewClients | Write-SqlTableData -ServerInstance $Instance -database $database -table $table -SchemaName $schema -Force
    
    Read-SqlTableData -ServerInstance $Instance -database $database -SchemaName $schema -table $table

    Write-Host -ForegroundColor Green "(SQL): Tasks Finished"

#generate output file for sql results
Invoke-Sqlcmd -Database ClientDB –ServerInstance .\SQLEXPRESS -Query ‘SELECT * FROM dbo.Client_A_Contacts’ > .\SqlResults.txt

}
catch{
    Write-Host -ForegroundColor Red "Error"
}