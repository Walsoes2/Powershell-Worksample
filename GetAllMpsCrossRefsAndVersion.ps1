(Param
[String]$ScomManagementServer = "<MANAGEMENTSERVER>",
[String]$FilePath = "C:\Admin\Results.csv"
)
# Load SCOM module
Import-Module OperationsManager

# Connect to the SCOM management group
New-SCOMManagementGroupConnection -ComputerName $ScomManagementServer

$AllMPs = Get-SCOMManagementPack

[System.Collections.ArrayList]$rapport = @()

$refs = $AllMPs | ForEach-Object { Select-Object -InputObject $_ -ExpandProperty references } 


for ($i = 0; $i -lt $AllMPs.Count; $i++)
{ 
    for ($j = 0; $j -lt $refs[$i].Name.Length; $j++)
                                            { 
    $rapport += [pscustomobject]@{
           Managmentpack       = $AllMPs[$i].Name
           References          = $refs[$j].Name
           KeyToken            = $refs[$j].KeyToken
           Version             = $refs[$j].Version
           Id                  = $refs[$j].Id  
           VersionId           = $refs[$j].VersionId
           
           }  
     } 
}

$rapport | Export-Csv $FilePath -Force -NoTypeInformation -NoClobber -Encoding UTF8
