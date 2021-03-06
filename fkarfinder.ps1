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

(Find-ADObjects "gaia" "user" "(userworkstations=*$ComputerName*)(cn=F*)" "cn,userworkstations").Properties