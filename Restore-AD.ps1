# Calvin Seok 011553098

function CreateAD {
    param (
        [string]$name,
        [string]$path,
        [string]$description
    )

    $OUdisplayname = "OU=$name,$path"

    # Check if the OU exists, if it already exists, delete it before recreating it. 
    try {
        $existingOU = Get-ADOrganizationalUnit -Identity $OUdisplayname
        Write-Verbose "OU '$OUdisplayname' already exists. Deleting..."
        Remove-ADOrganizationalUnit -Identity $existingOU -Recursive -Confirm:$false
    }
    catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
        Write-Verbose "Creating new OU '$OUdisplayname'"
    }

    New-ADOrganizationalUnit -Name $name -Path $path -Description $description
}

#Create AD OU
Write-Host -ForegroundColor Green "(AD): Preparing Tasks"
$AdRoot = (Get-ADDomain).DistinguishedName
$OUName = "Finance"
$OUDisplayName = "Finance"
$ADPath = "OU=$($OUName),$($AdRoot)"

try {
    CreateAD -name $OUName -path $AdRoot -description $OUDisplayName

    # Import finance personnel from CSV
    $NewFinanceUsers = Import-Csv $PSScriptRoot\financePersonnel.csv

    ForEach ($ADUser in $NewFinanceUsers) {
        $First = $ADUser.First_Name
        $Last = $ADUser.Last_Name
        $FullName = $First + " " + $Last
        $Postal = $ADUser.PostalCode
        $Office = $ADUser.OfficePhone
        $Mobile = $ADUser.MobilePhone

        New-ADUser -GivenName $First `
            -Surname $Last `
            -Name $FullName `
            -PostalCode $Postal `
            -OfficePhone $Office `
            -MobilePhone $Mobile `
            -Path $ADPath
    }
#Dump AD results into file ADResults.txt
    Get-ADUser -Filter * -SearchBase $ADPath -Properties DisplayName,PostalCode,OfficePhone,MobilePhone > .\AdResults.txt
    Write-Host -ForegroundColor Green "(AD): Tasks Finished"
}
catch {
    Write-Host -ForegroundColor Red "(AD): Sorry, an error occurred - $_"
}