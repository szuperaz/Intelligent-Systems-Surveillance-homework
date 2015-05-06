<#
.SYNOPSIS
Counts the number of colleagues per companies working with a university professor.

.PARAMETER Name
The name (Common Name) of the university professor

.PARAMETER OrderByPartners
If present, the output is sorted by the number of partners, descending.

.NOTES
IRF homework by Zizi
#>

#parameters
param (
	[Parameter(Mandatory=$true)][string] $Name,
	[switch] $OrderByPartners
)
#####################################################################################################################################

# CHECKING PARAMETERS
# checking $Name
try {
    $user = Get-ADUser $Name
}
catch {
    throw "$Name does not exsist!"
}
# teachers are stored in the Faculties OU
if (!$user.DistinguishedName.Contains("OU=Faculties")) {
    throw "$Name does not a university professor!"
}
#####################################################################################################################################

# COLLECTING THE COLLEAGUES WHO WORK WITH $NAME  
# what does this long-long command do?
# 1: selelcts the procejts in which $Name is participating
# (projects are groups with their DN containing 'OU=Projects')
# 2: selects the members of the projects
# (pjocets can contain gorups, users and computers, but we only select the users)
# 3: selects only those users who work for a partner company
# (partner companies and their workers are found in the Partners OU)
# 4: selects each user only once
$colleagues = Get-ADPrincipalGroupMembership $Name | Where-Object {$_.DistinguishedName -like "*OU=Projects*"} |
              Get-ADGroupMember -Recursive | Where-Object {$_.objectClass -eq "user"} |
              Where-Object{$_.DistinguishedName -like "*OU=Partners*"} |
              Select-Object -Unique
#####################################################################################################################################

# COUNT THE NUM OF COLLEAGUES/COMPANY
# dictionary: key: company name, value: num of collagues working with $Name
$companies_colleagues = @{}
foreach($colleague in $colleagues) {   
    $DNComponents = $colleague.DistinguishedName.Split(',')
    # the second part of the DN contains the company name
    $company = $DNComponents[1]
    $company = $company.SubString(3)
    if($companies_colleagues.ContainsKey($company)) {
        $lastNum = $companies_colleagues[$company]
        $lastNum++
        $companies_colleagues[$company] = $lastNum
    }
    else {
        $companies_colleagues.Add($company, 1)
    }
}
#####################################################################################################################################

#OUTPUT
if ($companies_colleagues.Count -eq 0) {
    Write-Output "$Name currently does not work with colleagues from partner companies"
}
else {
    # ordering by value, descending
    if ($OrderByPartners) {
        foreach ($key in ($companies_colleagues.GetEnumerator() | Sort-Object -Property Value -Descending).Key) {
            Write-Output "$key ($($companies_colleagues[$key]))"
        }
     }
    # ordering by key, ascending
    else {
        foreach ($key in ($companies_colleagues.GetEnumerator() | Sort-Object -Property Key).Key) {
        Write-Output "$key ($($companies_colleagues[$key]))"
        }
    }
}
        