<# 
MIT License

Copyright (c) 2025 Henrik Walsöe Vikström | Github: https://github.com/Walsoes/powershell

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice, this permission notice, and the following attribution clause shall be included in all copies or substantial portions of the Software:

"This software includes contributions from Henrik Walsöe Vikström (c) 2025. Original author must be credited where credit is due."

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

 #>

#requires -Version 5.1

[CmdletBinding()]
param (
    [Parameter(Mandatory=$false, Position=0)]
    [string]$Server1 = "pubt-ap-08",

    [Parameter(Mandatory=$false, Position=1)]
    [string]$Server2 = "pubt-ap-07",

    [Parameter(Mandatory=$false)]
    [bool]$ShouldRestoreCluster = $false,

    [Parameter(Mandatory=$false)]
    [string]$Filter = "Pubt",

    [Parameter(Mandatory=$false)]
    [bool]$Interactive = $true,

    [Parameter(Mandatory=$false)]
    [switch]$Help
)

# Usage and Help


if (-not $Server1 -or -not $Server2 -or $Help) {
    Write-Host ""
    Write-Host "ShouldRestoreCluster and Interactive is set to $true if nothing else is inputed. '-Filter' can be used as a lightweight way to make sure a specific cluster resource group is NOT moved to a specific node to avoid unnecessary downtime."
    Write-Host "" 
    Write-Host "Usage: '.\FailoverPatchScript.ps1 [-Server1 <Servername1/node1>] [-Server2 <Server2/Node2>] [-$ShouldRestoreCluster <$true/$false> ] [-Filter <""> [-Interactive <$true/$false>] '"
    Return;
}


# Initering av variable
$servers = @($Server1,$Server2)
[System.Collections.ArrayList]$formatedCluster = @()

$getUpdateParam = @{            
        NameSpace = 'root/ccm/ClientSDK'
        ClassName = 'CCM_SoftwareUpdate'
        Filter = 'EvaluationState < 8'
} 



$updatesAvailableOnServer = @()
$updatesNotAvailable = @()
$ClusterFilter = $Filter + "*"

$updatesAvailableOnServer.clear()
$updatesNotAvailable.clear()
$error.clear()


#Check if there are any updates to install

If((Test-Connection -ComputerName $servers -Count 1 -Quiet) -contains $false)
{
write-OUtput "One or both of the servers are not reachable"
Write-output "Check if they are powered on and pingable on the WS-man protocol. Enable PS-remoting on the target machine with 'winrm quickconfig' before trying again"
exit 0
} 


try {

$ErrorActionPreference = 'stop'

foreach($server in $servers) {
$updates = Get-CimInstance @getUpdateParam -ComputerName $server

     foreach ($update in $updates) {
    if ($updates.Count -gt -1) {
        Write-Host "Windows Updates available on $($server):" -ForegroundColor Green -NoNewline
        $updatesAvailableOnServer += $server
       
            Write-Host $update.name 

            
        } else 
        {
            $updatesNotAvailable += $server
            Write-host "`r"
            Write-Host "No updates available on $($server)" -foregroundcolor Yellow 
            Write-host "Continuing updating the other clustred server if there are any..."
           <# Exit #>
        }
    }
   }
  
}

catch { Write-Error "An error occurred: $error" }


if($updatesNotAvailable.Count -eq 2) { Write-host "No updates available for either of the servers. Exiting..." ; exit 0 > $null }



# Import the FailoverClusters module if not already imported
Import-Module FailoverClusters
Import-Module OperationsManager


# "Class"

#Adds a read-only class which will act as a backup when restoring the cluster
Add-Type -TypeDefinition @"
public class ClusterRolesBackup {
    public readonly string ClusterName;
    public readonly string[] Nodes;
    public readonly string[] State;
    public readonly string[] Rolenames;
    public readonly string[] Roles;
    public readonly string[] Status;
    public readonly string[] OwnerNode;

    public ClusterRolesBackup(string clustername, string[] nodes, string[] state,string[] rolenames, string[] roles, string[] status, string[] ownerNode){
        ClusterName = clustername;
        Nodes = nodes;
        State = state;
        Rolenames = rolenames;
        Roles = roles;
        Status = status;
        OwnerNode = ownerNode;
     
    }
}
"@ -ErrorAction Ignore #EA only for fixing the script

############################ FUNCTIONS-SECTION ##############################

# Set cluster in MM. If not specificed it is 180 min.
function Check-CCMIsUpdatesInstalled {
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSSession,

        [Parameter(Mandatory = $true)]
        [hashtable]$SplatParams
    )


    $scriptBlock = {
        param ($SplatParams)
        $InstalledUpdates = 0


        $updates = Get-CimInstance @SplatParams
        foreach ($update in $updates) {
            if ($update.State -eq 0) {  # 0 means the update is installed successfully
                $InstalledUpdates++
            }
        } 

        if($InstalledUpdates -eq $updates.count) { 

        return $true }
    }

    return Invoke-Command -Session $PSSession -ScriptBlock $scriptBlock -ArgumentList $SplatParams
}

function Set-MMforCluster {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [string]$ScomServer = "SCOM-AP-31",
        [Parameter(Mandatory = $true)]
        [string]$Clustername,
        [Parameter(Mandatory = $true)]
        [string[]]$Nodes,
        [Parameter(Mandatory = $true)]
        [string[]]$Roles,
        [Parameter(Mandatory = $false)]
        [int]$duration = 180
    )

    Invoke-Command -ComputerName $ScomServer -ScriptBlock {
        param ($Clustername, $Nodes, $Roles, $duration)

        Import-Module OperationsManager

        # Function to check if the instance is already in maintenance mode
        function Set-MaintenanceMode {
            param (
                [Parameter(Mandatory = $true)]
                [string]$displayName,
                [int]$duration,
                [string]$comment
            )

            $instance = Get-SCOMClassInstance -DisplayName $displayName   # FQDN of node

            if ($instance) {
                if ($instance.InMaintenanceMode -eq $false) {
                    Start-SCOMMaintenanceMode -Instance $instance -EndTime (Get-Date).AddMinutes($duration) -Comment $comment
                    $nodesJoined = $nodes -join "," 
                    $rolesJoined = $roles -join ","

                    Write-Host "$displayName is now in maintenance mode including $nodesJoined and $rolesJoined"
                } else {
                    Write-Host "$displayName is already in maintenance mode" -ForegroundColor DarkYellow
                }
            } else {
                Write-Host "Instance with display name '$displayName' not found."
            }
        }

        # Put the cluster in maintenance mode
        Set-MaintenanceMode -displayName $Clustername -duration $duration -comment "Planned Maintenance for Cluster"

        Start-Sleep -Seconds 15

    } -ArgumentList $Clustername, $Nodes, $Roles, $duration

    Write-Host "`r Finished." -ForegroundColor Green
}

function Test-PendingReboot {
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSsession
        )
   

return Invoke-Command -Session $PSsession -ScriptBlock {
   
   $hasRegvalue1 = (Get-Item "HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData").Valuecount -gt 0
   $hasRegValue2 = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\SMS\Mobile Client\Reboot Management\RebootData") -ne $NULL 
   
   if ($hasRegvalue1 -or $hasRegValue2) { 
   return $true } 
   else { 
   return $false }      
            
    }
}

function Invoke-Reboot { 

[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSsession 
        )   

    Invoke-Command -Session $PSsession -ScriptBlock {
    cmd.exe /c "msg * /TIME:30 This server will reboot in 30 seconds. Planned maintence."
    Start-Sleep -Seconds 30 }

    Restart-Computer -ComputerName $Pssession.ComputerName -Wait -For PowerShell -Timeout 900 -Delay 15 -Force # Wait up to 15 minutes before generating an error, checking every 15 secs if Powershell is Available on remote server
 }   
  
  <# Test if the patched computer is reacheable, i.e, has rebooted and is remotable
    do
    {
      Start-Sleep -Seconds 15
      $PingSuccessFull = [System.Net.Sockets.TcpClient]::new().ConnectAsync($Pssession.ComputerName, 5985).Wait(250)  #Use port 5986 if it has to be secure. 
      
    }
    while ($PingSuccessFull -ne $true)
       
   } #>

function Get-ClusterInformation {
[CmdletBinding()]
param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSsession,
        [Parameter(Mandatory = $false)]
        [string]$ClusterFilter = $ClusterFilter
        )

$clusterResources = Invoke-Command -Session $PSsession -ScriptBlock { param($filter) Get-ClusterResource | 
Where-Object { $_.Name -like $filter } } -ArgumentList $ClusterFilter | select Cluster, Name, State, Ownergroup,Ownernode

$clusterNodes = Invoke-Command -Session $PSsession -ScriptBlock { Get-ClusterNode }  # From Module FailoverClusters

$ClusterRolesstatus = New-Object ClusterRolesBackup -ArgumentList $clusterResources.Cluster[0],$clusterNodes.Name, $clusterNodes.State, $clusterResources.Name, $clusterResources.Ownergroup, $clusterResources.State, $clusterResources.OwnerNode

  return $ClusterRolesstatus

}

function Get-IsNotOwnerNode {
[CmdletBinding()]
  param (
        [Parameter(Mandatory = $true)]
        [System.Object]$clusterObject)


if(($clusterObject.ownerNode | Get-unique).count -eq 1)
    { 
        $IsNotOwnerNode = if ($clusterObject.Ownernode[0] -eq $Server1) { $server2 } else { $server1 }
        return $IsNotOwnerNode
    }
}

function Get-IsOwnerNode {
[CmdletBinding()]
  param (
        [Parameter(Mandatory = $true)]
        [System.Object]$clusterObject)


if(($clusterObject.ownerNode | Get-unique).count -eq 1)
    { 
        
        return $clusterObject.Ownernode[0]
    }
}

function Set-ClusterRolesOnline {
[CmdletBinding()]
  param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSsession,

        [Parameter(Mandatory = $true)]
        [System.Object]$clusterObject,

        [Parameter(Mandatory = $true)]
        [string]$ClusterFilter = $ClusterFilter #Systemacroym + "*" OR something used by "-like '<>'"
        
            )

    $clusterRoles = $clusterObject.Roles
    $ClusterRoleStatus = $clusterObject.status
 
 for ($i = 0; $i -lt $clusterRoles.count ; $i++)
 { 
            if($ClusterRoleStatus[$i].ToString() -ne "Online")
        {
          
    Write-Host "$($clusterroles[$i]):" -NoNewline
    Write-Host -ForegroundColor Red " $($ClusterRoleStatus[$i])"
      
    $startRole = Read-Host "Set $($clusterroles[$i]) status to 'Online' y/n? (if the role is not suppose to be online don't input 'y')" 
      if($startRole.ToLower() -eq "y" -or $interactive -eq $false)
      {
            try
            {
             $null = Invoke-Command -Session $PSsession -scriptblock { param($roleName) Start-ClusterResource -Name $roleName } -ArgumentList $($clusterRoles[$i]) -AsJob -JobName StartClusterRole -ErrorAction Stop
        
            # Oklart låter denna kommentar vara kvar 2025-04-15 innvoke-Command -Session $PSsession -scriptblock {  param($nodeName) Stop-ClusterGroup -name $nodeName } -ArgumentList  $clusterrole.Ownergroup.ToString() -AsJob -JobName stopClusterRole -ErrorAction Stop
            
             $finished = Wait-Job -Name StartClusterRole    

             if ($finished.State -ne "Completed") 
                 {
                        Write-host $finished.Error

                 } 
             else
                { 
                   return Invoke-Command -Session $PSsession -ScriptBlock { Get-ClusterResource | Where-Object { $_.Name -like $ClusterFilter } } | select Cluster, Name, State, Ownergroup,Ownernode 
    
                }
            
            }
            catch 
            {
                Write-Error "An error occurred: $error"
            }
            finally
            {
                Get-Job | Remove-Job -Force
            }

   

         }  
      }
   } 
}

function Set-ClusterNodesOnline {
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSsession,

        [Parameter(Mandatory = $true)]
        [System.Object]$clusterObject
    )

    $clusterNodes = $clusterObject.Nodes
    $clusterState = $clusterObject.state
 
   Write-host "Asserting that the cluster nodes are online in the cluster."


 for ($i = 0; $i -lt $clusterNodes.count; $i++)
 { 
            if($clusterState[$i].ToString() -ne "Up") {
        
    Write-Host "$($clusternode[$i]):" -NoNewline
    Write-Host -ForegroundColor Red " $($clusterState[$i])"
      

    $startNode = Read-Host "Set $($clusterNodes[$i]) status to 'Up' y/n?" 

      if($startNode.ToLower() -eq "y" -or $interactive -eq $false)
      {
            try
            {
            $null = Invoke-Command -Session $PSsession -scriptblock { param($nodeName) Start-ClusterNode -Name $nodename } -ArgumentList $clusterNodes[$i] -AsJob -JobName StartClusterNodes -ErrorAction Stop
            
            $finished = Wait-Job -Name StartClusterNodes

                if ($finished.State -ne "Completed") 
                  {
                   Write-host $finished.Error
                   }
                else 
                  { 
                  Invoke-Command -Session $PSsession -ScriptBlock { Get-ClusterNode | ft }
                  }    

            }
            catch 
            {
                Write-Error "An error occurred: $error"
            }

            finally
            { 
             Get-Job | Remove-Job -Force
            } 
    
    }  else { <#DO nothing#> }
     
   }
       else { Write-host "$($clusterNodes[$i]) is now Online" }

    
} }

function Move-ClusterResourceGroupToSingleNode {
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSsession,

        [Parameter(Mandatory = $true)]
        [System.Object]$clusterObject,

        [Parameter(Mandatory = $false)]
        [string]$ClusterNode = "",

        [Parameter(Mandatory = $false)]
        [string]$ClusterFilter = $ClusterFilter, #Systemacroym + "*" OR something used by "-like '<>'"

        [Parameter(Mandatory = $true)]
        [bool]$MoveWithoutLogic

    )

$ClusterResourceMoveToNode = $clusterNode  # The cluster node to move all resource groups to if $MoveWithoutLogic is $true
$ClusterResourceMoveFromNode = ""



If($MoveWithoutLogic -eq $true)
{
$Startresourcemove = Read-Host "Move all cluster resource groups to $ClusterResourceMoveToNode y/n?" 

      if($Startresourcemove.ToLower() -eq "y" -or $Interactive -eq $false)
      {
            try
            {
            $ClusterResourceMoveFromNode = Get-IsOwnerNode -clusterObject $clusterObject -ErrorAction Stop

            $null = Invoke-Command -Session $PSsession -scriptblock { param($ClusterResourceMoveFromNode,$ClusterResourceMoveToNode)
             
             Get-ClusterNode $ClusterResourceMoveFromNode | Get-ClusterGroup | Move-ClusterGroup -Node $ClusterResourceMoveToNode 
             
             } -ArgumentList $ClusterResourceMoveFromNode,$ClusterResourceMoveToNode -AsJob -JobName MoveClusterResources -ErrorAction Stop
            
            $finished = Wait-Job -Name MoveClusterResources

                if ($finished.State -ne "Completed") 
                  {
                   Write-host $finished.Error
                   }
                else 
                  { 
                  return Invoke-Command -Session $PSsession -ScriptBlock { Get-ClusterResource | Where-Object { $_.Name -like $ClusterFilter } } | select Cluster, Name, State, Ownergroup,Ownernode 
                  }    

            }
            catch 
            {
                Write-Error "An error occurred: $error"
            }

            finally
            { 
             Get-Job | Remove-Job -Force
            }

    } else { return }

}

elseif (($clusterObject.ownerNode | Get-unique).count -gt 1)
    {
     
   Write-host "All cluster roles or cluster resources groups are distrobuted amoung different nodes. Moving them to one single node..."

   $tallyNode1 = 0 #server1
   $tallyNode2 = 0 #server2
   $clusterResourcesGroups = $clusterObject.roles

 for ($i = 0; $i -lt $clusterResourcesGroups.count; $i++)
 { 
    if($clusterObject.Ownernode[$i] -eq $Server1)
    { $tallyNode1++}
    else
    { $tallyNode2++}

}

if ($tallyNode1 -ge $tallyNode2)
    {  
    
    Write-Host " $server1 " 
    $ClusterResourceMoveToNode = $server1    # The cluster node to move all resource groups to because it owns more or equal resource groups
    $ClusterResourceMoveFromNode = $server2
    }
    elseif($filter

    else
    {
    write-host " $Server2 "

    $ClusterResourceMoveToNode = $server2
    $ClusterResourceMoveFromNode = $server1    }
    
     

    $Startresourcemove = Read-Host "Move all cluster resource groups to $ClusterResourceMoveToNode y/n?" 

      if($Startresourcemove.ToLower() -eq "y" -or $Interactive -eq $false)
      {
            try
            {
            $null = Invoke-Command -Session $PSsession -scriptblock { param($ClusterResourceMoveFromNode,$ClusterResourceMoveToNode)
             
             Get-ClusterNode $ClusterResourceMoveFromNode | Get-ClusterGroup | Move-ClusterGroup -Node $ClusterResourceMoveToNode 
             
             } -ArgumentList $ClusterResourceMoveFromNode,$ClusterResourceMoveToNode -AsJob -JobName MoveClusterResources -ErrorAction Stop
            
            $finished = Wait-Job -Name MoveClusterResources

                if ($finished.State -ne "Completed") 
                  {
                   Write-host $finished.Error
                   }
                else 
                  { 
                  return Invoke-Command -Session $PSsession -ScriptBlock { Get-ClusterResource | Where-Object { $_.Name -like $ClusterFilter } } | select Cluster, Name, State, Ownergroup,Ownernode 
                  }    

            }
            catch 
            {
                Write-Error "An error occurred: $error"
            }

            finally
            { 
             Get-Job | Remove-Job -Force
            } 
    
    }  else { return }
     
   } }
 
function Restore-Cluster {  
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Object]$clusterObject

    )

    $Clustername = $clusterObject.ClusterName
    $ClusterroleName = $clusterObject.Roles
    $clusterOwnerNode = $clusterObject.Ownernode


      try
       {        
           for ($i = 0; $i -lt $clusterobject.Rolenames.count; $i++)
            
           { 

           Move-ClusterGroup -Cluster $Clustername -Name $ClusterroleName[$i] -Node $clusterOwnernode[$i] -IgnoreLocked -ErrorAction Stop
            }

       }
            catch 
            {
                Write-Error "An error occurred: $error"
            }
            
 }

function Print-Clusterinformation { 
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Object]$cluster
    )

[System.Collections.ArrayList]$formatedCluster = @()


for ($i = 0; $i -lt $cluster.Roles.Count; $i++) {
   $formatedCluster += [PSCustomObject]@{
        ClusterName = $cluster.ClusterName
        Node        = $cluster.Nodes[$i]
        State       = $cluster.State[$i]
        [char]124   = [char]124
        RoleName    = $cluster.Rolenames[$i]
        Role        = $cluster.Roles[$i]
        Status      = $cluster.Status[$i]
        OwnerNode   = $cluster.OwnerNode[$i]

    } 

  }
  return $formatedCluster
}

function Start-PatchingCluster {
[CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [System.Management.Automation.Runspaces.PSSession]$PSsession
        ,
        [Parameter(Mandatory = $true)]
        [hashtable]$getUpdateParam

        )
              
Invoke-Command -Session $PSsession -ScriptBlock { 
            
            Param($getUpdateParam) 

            $installUpdateParams = @{
                Namespace = 'root/ccm/ClientSDK'
                ClassName = 'CCM_SoftwareUpdatesManager'
                MethodName = 'InstallUpdates'
                Arguments = @{CCMUpdates = [ciminstance[]](Get-CimInstance @getUpdateParam) }
            }
            
            Invoke-CimMethod @installUpdateParams   

        } -ArgumentList $getUpdateParam
    


    #Check Íf updates has been downloaded and installed and also checking for pending reboot
    
    do  
    {
        $IsUpdatesInstalled = Check-CCMIsUpdatesInstalled -PSSession $PSsession -SplatParams $getUpdateParam 
        $HasPendingReboot = Test-PendingReboot -PSsession $PSsession
       
        Write-host "Checking every third minute if all updates has been installed and if a pending reboot is imminent." 

        Start-Sleep -Seconds 180

    } while((-not $IsUpdatesInstalled) -and (-not $HasPendingReboot))
   
    Write-Host "All updates have been installed on $($Pssession.computername) - initiating a forced reboot." -ForegroundColor Green

    Invoke-Reboot -PSsession $PSsession

}


##############################  Stage 1 "MAIN"

#Information gathering, set MM, print clusterresources, Set-clusternodes online (if necessary)
try
{
$ErrorActionPreference = 'Stop'

# Retrive data about the cluster Computername could be either server1 or server2. 
$PSsession = New-PSSession -ComputerName $server2 


# Retriving cluster information
$ClusterRolesBackup = Get-ClusterInformation $PSsession


#Set Cluster in MM
$setMaintence = Read-Host "Set the cluster in Maintence mode in SCOM? y/n?" 
      if($setMaintence.ToLower() -eq "y" -or $Interactive -eq $false)
      { Set-MMforCluster -Clustername $ClusterRolesBackup.ClusterName -Nodes $ClusterRolesBackup.Nodes -Roles $ClusterRolesBackup.Rolenames }
      else 
      { Write-host "Cluster is not in maintencemode but the script can still be runned." -ForegroundColor yellow }


Write-host "`r"
Write-Host  "Cluster Information" -NoNewline -BackgroundColor Black -foregroundColor green


Print-Clusterinformation $ClusterRolesBackup | Format-Table

Write-Host "`r" 
Write-Host "Notice the current state of the cluster, expecially that everything is Up/Online and what ownernode each cluster role has."
Write-Host "The script will first put all cluster-nodes and resources online (if desirable) and move all roles/resources to one node on a server"
Write-Host "If all roles are online and are on the same ownernode the other server/node will commence patching."
Write-Host "When that server has updated and rebooted the clusterroles will be moved to the newly patched server which then becomes the new ownernode and the server that has no roles on it will comence patching."
Write-Host "After that the cluster roles will be configured according to the clusterbackup object, i.e, the roles will be distributed evenly between servers if that is desirable." -ForegroundColor Yellow




#Check if all clusterNodes and resources are Up and move them to one single node. 

$Continue = Read-Host "Press y to continue"  ## INTERACTIVE

if($Continue.ToLower() -eq "y" -or $Interactive -eq $false) {

 
Set-ClusterNodesOnline -PSsession $PSsession -clusterObject $ClusterRolesBackup

Set-ClusterRolesOnline -PSsession $PSsession -clusterObject $ClusterRolesBackup -ClusterFilter $Filter

Move-ClusterResourceGroupToSingleNode -PSsession $PSsession -clusterObject $ClusterRolesBackup -MoveWithoutLogic $false

$IsNotOwnerNode = Get-IsNotOwnerNode -clusterObject (Get-ClusterInformation $PSsession)

 Write-Host "The Server with currently no Cluster Resource groups is $IsNotOwnerNode and is ready to be patched or restarted if already patched."

}
}

catch
{
    Write-Warning "An error occured: $_"
}

finally
{
    Get-PSSession | Remove-PSSession 
    $ErrorActionPreference = "Continue"
}
###################

    
#Check if all owner nodes are on the same server and start appyling updates on the other node
#----------------------------------------------------------------------



$ShouldStartUpdate = Read-Host "Start patching or rebooting a patched server? y/n?"

     if($ShouldStartUpdate.ToLower() -eq "y" -or $Interactive -eq $false)
     {
       
     try   {
            $ErrorActionPreference = "stop"   

            $PSsession = New-PSSession -ComputerName $IsNotOwnerNode
            
            $ComputerNameIfRebootIsNeeded = $servers | where { $_ -notin $updatesAvailableOnServer }

             if($ComputerNameIfRebootIsNeeded -ne $NULL ) {  #göe bättre FROM start of script if a server has been patched but not rebooted, skip patch section and move on to STAGE 2
            
             
                $PSsession = New-PSSession -ComputerName $ComputerNameIfRebootIsNeeded
            
                $rebootReq = Test-PendingReboot -PSsession $PSsession


                    if($rebootReq) { Invoke-Reboot $PSsession; Write-host "$($PSsession.ComputerName) has been rebooted" -ForegroundColor Green }

                    elseif($rebootReq -eq $false)
                    { 
                       Write-host "Only one server has awating patches the other node/server has been patched and rebooted" 
            
                    }

                                                            }

            else {
            
                        
            Write-Host "Applying updates on $($IsNotOwnerNode):"     

            Start-PatchingCluster -PSsession $PSsession -getUpdateParam $getUpdateParam
                        
            Write-Host "The first node/server, $($IsNotOwnerNode), has been patched and rebooted and is ready to host all clustergroup resources" -ForegroundColor Green
      }
            
                        
       
           
  }         

    catch { Write-Error "An error has occured: $_ "}
     
    finally { 

          Get-PSSession | Remove-PSSession  #remove any pssesssions if there are any else make a new one


         $PSsession = New-PSSession $IsNotOwnerNode

         $ClusterInfoAfterStage1 = Get-ClusterInformation $PSsession 
                
         $ClusterInfoAfterStage1

         #$PSsession    
         $IsOwnerNode = Get-IsOwnerNode $ClusterInfoAfterStage1 # Used later in stage 2 just to skip opening a pssession
      
        $ErrorActionPreference = "Continue"  
            } 

     }
    
   #### STAGE 1 completed: Updates applied on 1 server and server 1 has rebooted.
     

#############################   Stage 2 "MAIN2
$ShouldStartUpdate = Read-Host "Move all cluster resources group to the patched node, $($IsNotOwnerNode), and continue patchning? y/n?"

     if($ShouldStartUpdate.ToLower() -eq "y" -or $Interactive -eq $false)
     {
       
     try   {
            $ErrorActionPreference = "stop"
          
            $PSsession = New-PSSession -ComputerName $IsOwnerNode
            
            Move-ClusterResourceGroupToSingleNode -PSsession $PSsession -clusterObject $ClusterInfoAfterStage1 -ClusterNode $IsNotOwnerNode -MoveWithoutLogic $true
            
            Write-Host "Applying updates on $($IsOwnerNode):"     

            Start-PatchingCluster -PSsession $PSsession -getUpdateParam $getUpdateParam

            }

    catch { Write-Error "An error has occured: $_ "}
     
    finally { 
     
      Get-PSSession | Remove-PSSession  #remove any pssesssions if there are any else make a new one


         $PSsession = New-PSSession $IsOwnerNode

      $ClusterInfoAfterStage2 = Get-ClusterInformation $PSsession 


      
      Write-Host "Both servers/nodes in the cluster has been patched and clusterresource groups are being restored with the backupobject if the boolean 'ShouldRestoreCluster' is set to $true" 

      if($ShouldRestoreCluster -eq $true) {

      Restore-Cluster $ClusterRolesBackup 


      Print-Clusterinformation $ClusterInfoAfterStage2 | ft

      Write-Host "Cluster has been patched."
      
      }
      
     else { 
     
     Print-Clusterinformation $ClusterInfoAfterStage2 | ft

     Write-Host "Cluster has been patched."
     
    }
      

  } 

}
    
    

