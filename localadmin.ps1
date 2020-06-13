param([parameter(Mandatory=$true)][string]$ComputerName)

function Find-ADObjects($domain, $class, $filter, $attributes = "distinguishedName")
{
    $dc = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext ([System.DirectoryServices.ActiveDirectory.DirectoryContextType]"domain", $domain);
    $dn = [System.DirectoryServices.ActiveDirectory.Domain]::GetDomain($dc);
    
    $ds = New-Object System.DirectoryServices.DirectorySearcher;
    $ds.SearchRoot = $dn.GetDirectoryEntry();
    $ds.SearchScope = "subtree";
    $ds.PageSize = 1024;
    $ds.Filter = "(&(objectCategory=$class)$filter)";
    $ds.PropertiesToLoad.AddRange($attributes.Split(","))
    $result = $ds.FindAll();
    $ds.Dispose();
    return $result;
}
#$ComputerName = "LSFLS14367461"
$adminRoles = @("CN=Kar_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Kar,OU=HealthCare,DC=gaia,DC=sll,DC=se","CN=Sos_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Sos,OU=HealthCare,DC=gaia,DC=sll,DC=se","CN=Lit_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lit,OU=Administration,DC=gaia,DC=sll,DC=se","CN=Ita_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ita,OU=Reference,DC=gaia,DC=sll,DC=se","CN=Dan_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Dan,OU=HealthCare,DC=gaia,DC=sll,DC=se","CN=Hsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Hsf,OU=Administration,DC=gaia,DC=sll,DC=se","CN=Lsf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Lsf,OU=Administration,DC=gaia,DC=sll,DC=se","CN=Fut_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Fut,OU=PublicTransportation,DC=gaia,DC=sll,DC=se","CN=Int_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Int,OU=Administration,DC=gaia,DC=sll,DC=se","CN=Trf_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Trf,OU=Administration,DC=gaia,DC=sll,DC=se","CN=Sll_Wrk_LocalAdmin_SLLeKlient,OU=Workstation,OU=Groups,OU=Sll,DC=gaia,DC=sll,DC=se","CN=Ser_Wrk_LocalAdmin_SLLeKlient,OU=eApplication,OU=Groups,OU=Ser,OU=Administration,DC=gaia,DC=sll,DC=se")
$adminRolesRegex = [string]::Join('|',$adminRoles)
$adVarde = (Find-ADObjects "gaia" "computer" "(cn=$ComputerName)" "cn,MemberOf").Properties
if ($adVarde.memberof -match $adminRolesRegex) {
    $true
}
else {
    $false
}