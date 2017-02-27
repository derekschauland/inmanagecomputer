<#$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sut = (Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace '\.Tests\.', '.'
. "$here\$sut"
#>
Import-Module ".\INmanagecomputer.psm1"

InModuleScope inmanagecomputer {
	
	Describe "Rename-INComputer" -Tags "Function" {
		
		It "Function rename-incomputer exists" { test-path Function:\Rename-InComputer | should be $true }
		
		$testcomputername = $env:computernam
		$tesnewname = $env:COMPUTERNAME + "-new"
		$testprefix = "PC-"
		
		
		Context "Input Paramteer tests" {
			
			
			it "Should not accept null computername parameter" {
				{ rename-incomputer -computername '' } | Should throw
			}
			
			it "Should not accept a null newname parameter"{
				{ rename-incomputer -newname '' } | Should throw
				
			}
			
			it "should not accept a null prefix parameter"{
				{ rename-incomputer -prefix ''} | Should throw
			}
			
			it "Should accept a null reboot parameter"{
				{ rename-incomputer -computername $testcomputername -newname $testnewname -testprefix $testprefix -reboot } | Should not be null
			}
			
			
			
		}
		
		Context "Group Policy Checking"{
			
			#$gpos = Get-RemoteAppliedGPOs -ComputerName $testcomputername
			
			$gpos = mock  get-remoteappliedgpos { $testcomputername }
			$credential = mock Get-Credential
			
			
			#((Get-CimInstance win32_computersystem -ComputerName $computer -Filter "Caption Like '%'").partofdomain -eq $true)
			
			it "Gpos should not be null"{
				{ $gpos } | Should not be null
			}
			
			$gpoWinrm = $gpos.appliedgpos
			
			it "GPOWinRM should not be null" {
				{ $gpowinrm } | Should not be null
			}
			
			$RM = (($gpoWinrm | select name) -match "WinRM")
			$WMI = (($gpoWinrm | select name) -match "WMI Firewall")
			
			
			it "RM should not be null" {
				{ $RM } | Should not be null
			}
			
			it "WMI should not be null" {
				{ $WMI } | Should not be null
			}
		}
		Context "Function rename-incomputer execution testing"{
#			it "Should have GPOs applied to allow for WINRM" {
#				{$RM} |  Should Match "WinRM"
#			}
			
			it "Should have GPOs applied to allow for WMI" {
				{$WMI} | Should belike "*WMI*"
			}
			
			it "Credential should not be null"{
				
				{ $credential} | Should not be null
			}
			
		}
		
		it "Check to see if the computer is domain joined" {
			{ (Get-CimInstance win32_computersystem -computername $testcomputername -filter "Caption Like '%'").partofdomain } | Should not be null
		}
	}
	
	
}

