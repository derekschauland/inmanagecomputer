﻿
Function Get-RemoteAppliedGPOs
{
    <# 
    .SYNOPSIS 
       Gather applied GPO information from local or remote systems. 
    .DESCRIPTION 
       Gather applied GPO information from local or remote systems. Can utilize multiple runspaces and  
       alternate credentials. 
    .PARAMETER ComputerName 
       Specifies the target computer for data query. 
    .PARAMETER ThrottleLimit 
       Specifies the maximum number of systems to inventory simultaneously  
    .PARAMETER Timeout 
       Specifies the maximum time in second command can run in background before terminating this thread. 
    .PARAMETER ShowProgress 
       Show progress bar information 
 
    .EXAMPLE 
       $a = Get-RemoteAppliedGPOs 
       $a.AppliedGPOs |  
            Select Name,AppliedOrder | 
            Sort-Object AppliedOrder 
        
       Name                            appliedOrder 
       ----                            ------------ 
       Local Group Policy                         1 
        
       Description 
       ----------- 
       Get all the locally applied GPO information then display them in their applied order. 
 
    .NOTES 
       Author: Zachary Loeber 
       Site: http://www.the-little-things.net/ 
       Requires: Powershell 2.0 
 
       Version History 
       1.0.0 - 09/01/2013 
        - Initial release 
    #>	
	[CmdletBinding()]
	Param
	(
		[Parameter(HelpMessage = "Computer or computers to gather information from",
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[Alias('DNSHostName', 'PSComputerName')]
		[string[]]$ComputerName = $env:computername,
		[Parameter(HelpMessage = "Maximum number of concurrent threads")]
		[ValidateRange(1, 65535)]
		[int32]$ThrottleLimit = 32,
		[Parameter(HelpMessage = "Timeout before a thread stops trying to gather the information")]
		[ValidateRange(1, 65535)]
		[int32]$Timeout = 120,
		[Parameter(HelpMessage = "Display progress of function")]
		[switch]$ShowProgress,
		[Parameter(HelpMessage = "Set this if you want the function to prompt for alternate credentials")]
		[switch]$PromptForCredential,
		[Parameter(HelpMessage = "Set this if you want to provide your own alternate credentials")]
		[System.Management.Automation.Credential()]
		$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	
	BEGIN
	{
		# Gather possible local host names and IPs to prevent credential utilization in some cases 
		Write-Verbose -Message 'Remote Applied GPOs: Creating local hostname list'
		$IPAddresses = [net.dns]::GetHostAddresses($env:COMPUTERNAME) | Select-Object -ExpandProperty IpAddressToString
		$HostNames = $IPAddresses | ForEach-Object {
			try
			{
				[net.dns]::GetHostByAddress($_)
			}
			catch
			{
				# We do not care about errors here... 
			}
		} | Select-Object -ExpandProperty HostName -Unique
		$LocalHost = @('', '.', 'localhost', $env:COMPUTERNAME, '::1', '127.0.0.1') + $IPAddresses + $HostNames
		
		Write-Verbose -Message 'Remote Applied GPOs: Creating initial variables'
		$runspacetimers = [HashTable]::Synchronized(@{ })
		$runspaces = New-Object -TypeName System.Collections.ArrayList
		$bgRunspaceCounter = 0
		
		if ($PromptForCredential)
		{
			$Credential = Get-Credential
		}
		
		Write-Verbose -Message 'Remote Applied GPOs: Creating Initial Session State'
		$iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()
		foreach ($ExternalVariable in ('runspacetimers', 'Credential', 'LocalHost'))
		{
			Write-Verbose -Message "Remote Applied GPOs: Adding variable $ExternalVariable to initial session state"
			$iss.Variables.Add((New-Object -TypeName System.Management.Automation.Runspaces.SessionStateVariableEntry -ArgumentList $ExternalVariable, (Get-Variable -Name $ExternalVariable -ValueOnly), ''))
		}
		
		Write-Verbose -Message 'Remote Applied GPOs: Creating runspace pool'
		$rp = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $ThrottleLimit, $iss, $Host)
		$rp.ApartmentState = 'STA'
		$rp.Open()
		
		# This is the actual code called for each computer 
		Write-Verbose -Message 'Remote Applied GPOs: Defining background runspaces scriptblock'
		$ScriptBlock = {
			[CmdletBinding()]
			Param
			(
				[Parameter(Position = 0)]
				[string]$ComputerName,
				[Parameter(Position = 1)]
				[int]$bgRunspaceID
			)
			$runspacetimers.$bgRunspaceID = Get-Date
			
			try
			{
				Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: Start' -f $ComputerName)
				$WMIHast = @{
					ComputerName = $ComputerName
					ErrorAction = 'Stop'
				}
				if (($LocalHost -notcontains $ComputerName) -and ($Credential -ne $null))
				{
					$WMIHast.Credential = $Credential
				}
				
				# General variables 
				$GPOPolicies = @()
				$PSDateTime = Get-Date
				
				#region GPO Data 
				
				$GPOQuery = Get-WmiObject @WMIHast `
				-Namespace "ROOT\RSOP\Computer" `
				-Class 'RSOP_GPLink' `
				| Where-Object {$_.appliedorder -ne 0} |
				Select @{ n = 'linkOrder'; e = { $_.linkOrder } },
					   @{ n = 'appliedOrder'; e = { $_.appliedOrder } },
					   @{ n = 'GPO'; e = { $_.GPO.ToString().Replace("RSOP_GPO.", "") } },
					   @{ n = 'Enabled'; e = { $_.Enabled } },
					   @{ n = 'noOverride'; e = { $_.noOverride } },
					   @{ n = 'SOM'; e = { [regex]::match($_.SOM, '(?<=")(.+)(?=")').value } },
					   @{ n = 'somOrder'; e = { $_.somOrder } }
				foreach ($GP in $GPOQuery)
				{
					$AppliedPolicy = Get-WmiObject @WMIHast `
					-Namespace 'ROOT\RSOP\Computer' `
					-Class 'RSOP_GPO' -Filter $GP.GPO
					$ObjectProp = @{
						'Name' = $AppliedPolicy.Name
						'GuidName' = $AppliedPolicy.GuidName
						'ID' = $AppliedPolicy.ID
						'linkOrder' = $GP.linkOrder
						'appliedOrder' = $GP.appliedOrder
						'Enabled' = $GP.Enabled
						'noOverride' = $GP.noOverride
						'SourceOU' = $GP.SOM
						'somOrder' = $GP.somOrder
					}
					
					$GPOPolicies += New-Object PSObject -Property $ObjectProp
				}
				
				Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: Share session information' -f $ComputerName)
				
				# Modify this variable to change your default set of display properties 
				$defaultProperties = @('ComputerName', 'AppliedGPOs')
				$ResultProperty = @{
					'PSComputerName' = $ComputerName
					'PSDateTime' = $PSDateTime
					'ComputerName' = $ComputerName
					'AppliedGPOs' = $GPOPolicies
				}
				$ResultObject = New-Object -TypeName PSObject -Property $ResultProperty
				
				# Setup the default properties for output 
				$ResultObject.PSObject.TypeNames.Insert(0, 'My.AppliedGPOs.Info')
				$defaultDisplayPropertySet = New-Object System.Management.Automation.PSPropertySet('DefaultDisplayPropertySet', [string[]]$defaultProperties)
				$PSStandardMembers = [System.Management.Automation.PSMemberInfo[]]@($defaultDisplayPropertySet)
				$ResultObject | Add-Member MemberSet PSStandardMembers $PSStandardMembers
				#endregion GPO Data 
				
				Write-Output -InputObject $ResultObject
			}
			catch
			{
				Write-Warning -Message ('Remote Applied GPOs: {0}: {1}' -f $ComputerName, $_.Exception.Message)
			}
			Write-Verbose -Message ('Remote Applied GPOs: Runspace {0}: End' -f $ComputerName)
		}
		
		function Get-Result
		{
			[CmdletBinding()]
			Param
			(
				[switch]$Wait
			)
			do
			{
				$More = $false
				foreach ($runspace in $runspaces)
				{
					$StartTime = $runspacetimers.($runspace.ID)
					if ($runspace.Handle.isCompleted)
					{
						Write-Verbose -Message ('Remote Applied GPOs: Thread done for {0}' -f $runspace.IObject)
						$runspace.PowerShell.EndInvoke($runspace.Handle)
						$runspace.PowerShell.Dispose()
						$runspace.PowerShell = $null
						$runspace.Handle = $null
					}
					elseif ($runspace.Handle -ne $null)
					{
						$More = $true
					}
					if ($Timeout -and $StartTime)
					{
						if ((New-TimeSpan -Start $StartTime).TotalSeconds -ge $Timeout -and $runspace.PowerShell)
						{
							Write-Warning -Message ('Timeout {0}' -f $runspace.IObject)
							$runspace.PowerShell.Dispose()
							$runspace.PowerShell = $null
							$runspace.Handle = $null
						}
					}
				}
				if ($More -and $PSBoundParameters['Wait'])
				{
					Start-Sleep -Milliseconds 100
				}
				foreach ($threat in $runspaces.Clone())
				{
					if (-not $threat.handle)
					{
						Write-Verbose -Message ('Remote Applied GPOs: Removing {0} from runspaces' -f $threat.IObject)
						$runspaces.Remove($threat)
					}
				}
				if ($ShowProgress)
				{
					$ProgressSplatting = @{
						Activity = 'Remote Applied GPOs: Getting info'
						Status = 'Remote Applied GPOs: {0} of {1} total threads done' -f ($bgRunspaceCounter - $runspaces.Count), $bgRunspaceCounter
						PercentComplete = ($bgRunspaceCounter - $runspaces.Count) / $bgRunspaceCounter * 100
					}
					Write-Progress @ProgressSplatting
				}
			}
			while ($More -and $PSBoundParameters['Wait'])
		}
	}
	PROCESS
	{
		foreach ($Computer in $ComputerName)
		{
			$bgRunspaceCounter++
			$psCMD = [System.Management.Automation.PowerShell]::Create().AddScript($ScriptBlock)
			$null = $psCMD.AddParameter('bgRunspaceID', $bgRunspaceCounter)
			$null = $psCMD.AddParameter('ComputerName', $Computer)
			$null = $psCMD.AddParameter('Verbose', $VerbosePreference)
			$psCMD.RunspacePool = $rp
			
			Write-Verbose -Message ('Remote Applied GPOs: Starting {0}' -f $Computer)
			[void]$runspaces.Add(@{
					Handle = $psCMD.BeginInvoke()
					PowerShell = $psCMD
					IObject = $Computer
					ID = $bgRunspaceCounter
				})
			Get-Result
		}
	}
	END
	{
		Get-Result -Wait
		if ($ShowProgress)
		{
			Write-Progress -Activity 'Remote Applied GPOs: Getting share session information' -Status 'Done' -Completed
		}
		Write-Verbose -Message "Remote Applied GPOs: Closing runspace pool"
		$rp.Close()
		$rp.Dispose()
	}
}


function rename-incomputer
{
	<# .externalhelp Inmanagecomputer.psm1-Help.xml #>
	
	[CmdletBinding()]
	
	param (
		[Parameter(Mandatory=$true,ValueFromPipeline=$true, ValueFromPipelineByPropertyName=$true)][string[]]$computername,
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)]
		[string[]]$newname,
		[Parameter(Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)][string]$prefix,
		[switch]$reboot
	)
	

	if ((Test-Path $computername) -eq $true)
	{
		Write-verbose "Using CSV FIle...`n"
		$computername = Import-Csv $computername -Delimiter "," | select -ExpandProperty computername
	}
	else
	{
		Write-verbose "Using supplied parameters..."
		$computername = $computername
	}
	
	
	foreach ($computer in $computername)
	{
		if (!(Test-Connection $($computer) -ErrorAction SilentlyContinue))
		{
			Write-Host "$($computer) not available at this time"
		}
		else
		{
			$gpos = Get-RemoteAppliedGPOs -ComputerName $computer
			$gpoWinrm = $gpos.appliedgpos
			
			$RM = (($gpoWinrm | select name) -match "WinRM")
			$WMI = (($gpoWinrm | select name) -match "WMI Firewall")
			
			if (($RM) -and ($wmi))
			{
				Write-Host "WinRM GPOs were found - $($gpoWinrm.name) - to allow management through WinRM and WMI."
				$quit = "0"
			}
			else
			{
				Write-Host "Cannot rename and reboot computers remotely with PowerShell - WinRM is not configured completely:`n Need to Firewall. See Weblink. Script Exiting."
				$ie = New-Object -com internetexplorer.application
				$ie.navigate2("http://www.grouppolicy.biz/2014/05/enable-winrm-via-group-policy/")
				$ie.visible = $true
				$quit = "1"
			}
			
			
			#	if ($os -match 10)
			#	{
			#		Write-LogInfo -LogPath $fulllogpath -Message "[$(time-now)] WinRM GPOs are applied to $env:COMPUTERNAME - Proceeding."
			#	}
			#	
			
			if ($quit -eq "1")
			{
				break;
			}
			
			if ($credential -eq $null)
			{
				$credential = Get-Credential -Message "Credentials not provided yet - please provide Domain or Local Credentials as required."
				#Write-LogInfo -LogPath $fulllogpath -Message "[$(time-now)] Credentials provided to allow rename and reboot."
			}
			
			if (!$reboot)
			{
				Write-Host "Note: Computers renamed will need to be manually rebooted!"
				
				foreach ($computer in $computername)
				{
					$newcomputername = $prefix + "-" + $newname
					
					if ((Get-CimInstance win32_computersystem -ComputerName $computer -Filter "Caption Like '%'").partofdomain -eq $true)
					{
						Write-Host "Computer $computer is part of a Domain - will use provided credentials as Domain Credentials"
						#Write-LogInfo -LogPath $fulllogpath -Message "[$(time-now)] Computer is joined to a Domain - Domain Credentials should have been provided"
						
						Rename-Computer -ComputerName $computer -NewName $newcomputername -DomainCredential $credential
						Write-Host "Computer $computer has been renamed to $newcomputername - you will need to restart to see the changes."
						#Write-LogInfo -LogPath $fulllogpath -Message "[$(time-now)] Computer has been renamed - no restart specified. A manual restart will be required."
					}
					else
					{
						Write-Host "Computer $computer is not part of a domain - will use provided credentials as Local"
						#Write-LogInfo -LogPath $fulllogpath -message "Computer $computer is not part of a Domain - will use provided credentials as Local Credentials"
						Rename-Computer -ComputerName $computer -NewName $newcomputername -LocalCredential $credential
						
						Write-Host "Computer $computer has been renamed to $newcomputername - you will need to restart to see the changes."
						
						#Write-LogInfo -LogPath $fulllogpath -Message "[$(time-now)] Computer has been renamed - no restart specified. A manual restart will be required."
					}
				}
			}
			else
			{
				Write-Host "Note: Computers renamed will be rebooted right away!"
				
				foreach ($computer in $computername)
				{
					$newcomputername = $prefix + "-" + $newname
					
					if ((Get-CimInstance win32_computersystem -ComputerName $computer -Filter "Caption Like '%'").partofdomain -eq $true)
					{
						Write-Host "Computer $computer is part of a Domain - will use provided credentials as Domain Credentials"
						#Write-LogInfo -LogPath $fulllogpath -message "Computer $computer is part of a Domain - Domain Credentials should have been provided"
						
						Rename-Computer -ComputerName $computer -NewName $newcomputername -DomainCredential $credential
						#Write-LogInfo -LogPath $fulllogpath -message "Computer $computer has been renamed and will reboot in 10 seconds"
						
						Write-Host "Computer $computer has been renamed to $newcomputername - and will reboot in 10 seconds."
						Start-Sleep 10
						Restart-Computer -ComputerName $computer -Force
						
					}
					else
					{
						Write-Host "Computer $computer is not part of a domain - will use provided credentials as Local"
						#Write-LogInfo -LogPath $fulllogpath -message "Computer $computer is not part of a Domain - Local Credentials should have been provided."
						
						Rename-Computer -ComputerName $computer -NewName $newcomputername -LocalCredential $credential
						#Write-LogInfo -LogPath $fulllogpath -message "Computer $computer will be restarted after a 10 second delay."
						
						Write-Host "Computer $computer has been renamed to $newcomputername - and will reboot in 10 seconds."
						Start-Sleep 10
						Restart-Computer -ComputerName $computer -force
					}
				}
			}
			
		}
		
	}
	
	
}

Export-ModuleMember -Function rename-incomputer
Export-ModuleMember -Function Get-RemoteAppliedGPOs

#use a -filter with ciminstance on Powershell 2.0 environments to avoid DMTF issues
#itknowledgeexchange.techtarget.com/powershell/cim-session-oddity



