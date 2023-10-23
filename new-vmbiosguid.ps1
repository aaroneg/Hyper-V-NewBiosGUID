function New-VMBIOSGUID
{
	<#
	.SYNOPSIS
		Changes the BIOSGUID for Hyper-V guests running on Hyper-V versions 8/2012 or later.
	.DESCRIPTION
		Changes the BIOSGUID for Hyper-V guests running on Hyper-V versions 8/2012 or later.
		A GUID can be supplied. If not, one is automatically generated.
		If the virtual machine is running, this script will attempt to shut it down prior to the operation. Once the replacement is complete, the virtual machine will be turned back on.
	.PARAMETER VM
		The name or virtual machine object (from Get-VM) of the virtual machine whose BIOSGUID is to be changed.
	.PARAMETER NewID
		The new GUID to assign to the virtual machine. If empty, a new GUID will be automatically generated.
	.PARAMETER ComputerName
		The Hyper-V host that owns the virtual machine to be modified.
	.PARAMETER Timeout
		Number of seconds to wait when shutting down the guest before assuming the shutdown failed and ending the script.
		Default is 300 (5 minutes).
		If the virtual machine is off, this parameter has no effect.
	.PARAMETER Force
		Suppresses prompts. If this parameter is not used, you will be prompted to shut down the virtual machine if it is running and you will be prompted to replace the BIOSGUID.
		Force can shut down a running virtual machine. It cannot affect a virtual machine that is saved or paused.
	.PARAMETER WhatIf
		Performs normal WhatIf operations by displaying the change that would be made. However, the new BIOSGUID is automatically generated on each run. The one that WhatIf displays will not be used.
	.NOTES
		Version 1.0
		February 29, 2016
		Author: Eric Siron
		(c) 2016 Altaro Software

		This script comes with no warranty, express or implied. Neither Altaro Software nor Eric Siron are liable for any damages, intentional or otherwise, that arise from its use in any capacity.
	.INPUTS
		Microsoft.HyperV.PowerShell.VirtualMachine or System.String
		System.GUID
	.EXAMPLE
		New-VMBIOSGUID -VM svtest
		
		Replaces the BIOS GUID on the virtual machine named svtest with an automatically-generated ID.

	.EXAMPLE
		New-VMBIOSGUID svtest

		Exactly the same as example 1; uses positional parameter.

	.EXAMPLE
		Get-VM svtest | New-VMBIOSGUID

		Exactly the same as example 1 and 2; uses the pipeline.

	.EXAMPLE
		New-VMBIOSGUID svtest -Force

		Exactly the same as examples 1, 2, and 3; prompts suppressed.

	.EXAMPLE
		New-VMBIOSGUID svtest -NewID $Guid

		Replaces the BIOS GUID of svtest with the supplied ID. These IDs can be generated with [System.Guid]::NewGuid().

	.EXAMPLE
		New-VMBIOSGUID svtest -WhatIf

		Shows how the BIOS GUID will be changed. TIP: Use this to view the current BIOS GUID without changing it.
	#>
	#requires -Version 4
	#requires -Modules Hyper-V
	#requires -RunAsAdministrator

	[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='High')]
	param
	(
		[Parameter(Mandatory=$true, ValueFromPipeline=$true, Position=1)][PSObject]$VM,
		[Parameter()][System.GUID]$NewID,
		[Parameter()][String]$ComputerName = $env:COMPUTERNAME,
		[Parameter()][UInt32]$Timeout = 300,
		[Parameter()][Switch]$Force
	)

	begin
	{
		<# adapted from http://blogs.msdn.com/b/taylorb/archive/2008/06/18/hyper-v-wmi-rich-error-messages-for-non-zero-returnvalue-no-more-32773-32768-32700.aspx #>
		function Process-WMIJob
		{
			param
			(
				[Parameter(ValueFromPipeline=$true)][System.Management.ManagementBaseObject]$WmiResponse,
				[Parameter()][String]$WmiClassPath = $null,
				[Parameter()][String]$MethodName = $null,
				[Parameter()][String]$VMName,
				[Parameter()][String]$ComputerName
			)
	
			process
			{
				$ErrorCode = 0
 
				if($WmiResponse.ReturnValue -eq 4096)
				{
					$Job = [WMI]$WmiResponse.Job
 
					while ($Job.JobState -eq 4)
					{
				
						Write-Progress -Activity ('Modifying virtual machine {0}' -f $VMName, $ComputerName) -Status ('{0}% Complete' -f $Job.PercentComplete) -PercentComplete $Job.PercentComplete
						Start-Sleep -Milliseconds 100
						$Job.PSBase.Get()
					}
 
					if($Job.JobState -ne 7)
					{
						if ($Job.ErrorDescription -ne "")
						{
							throw $Job.ErrorDescription
						}
						else
						{
							$ErrorCode = $Job.ErrorCode
						}
						Write-Progress $Job.Caption "Completed" -Completed $true
					}
				}
				elseif ($WmiResponse.ReturnValue -ne 0)
				{
					$ErrorCode = $WmiResponse.ReturnValue
				}
 
				if($ErrorCode -ne 0)
				{
					if($WmiClassPath -and $MethodName)
					{
						$PSWmiClass = [WmiClass]$WmiClassPath
						$PSWmiClass.PSBase.Options.UseAmendedQualifiers = $true
						$MethodQualifiers = $PSWmiClass.PSBase.Methods[$MethodName].Qualifiers
						$IndexOfError = [System.Array]::IndexOf($MethodQualifiers["ValueMap"].Value, [String]$ErrorCode)
						if($IndexOfError -ne "-1")
						{
							throw('Error Code: {0}, Method: {1}, Error: {2}' -f $ErrorCode, $MethodName, $MethodQualifiers["Values"].Value[$IndexOfError])
						}
						else
						{
							throw('Error Code: {0}, Method: {1}, Error: Message Not Found' -f $ErrorCode, $MethodName)
						}
					}
				}
			}
		}
	}
	process
	{
		Write-Verbose -Message 'Validating input...'
		$VMName = ''
		$InputType = $VM.GetType()
		if($InputType.FullName -eq 'System.String')
		{
			$VMName = $VM
		}
		elseif($InputType.FullName -eq 'Microsoft.HyperV.PowerShell.VirtualMachine')
		{
			$VMName = $VM.Name
			$ComputerName = $VM.ComputerName
		}
		else
		{
			throw('You must supply a virtual machine name or object.')
		}

		if($NewID -ne $null)
		{
			try
			{
				$NewID = [System.Guid]::Parse($NewID)
			}
			catch
			{
				throw('Provided GUID cannot be parsed. Supply a valid GUID or leave empty to allow an ID to be automatically generated.')
			}
		}

		Write-Verbose -Message ('Establishing WMI connection to Virtual Machine Management Service on {0}...' -f $ComputerName)
		$VMMS = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_VirtualSystemManagementService -ComputerName $ComputerName
		Write-Verbose -Message 'Acquiring an empty paramater object for the ModifySystemSettings function...'
		$ModifySystemSettingsParams = $VMMS.GetMethodParameters('ModifySystemSettings')
		Write-Verbose -Message ('Establishing WMI connection to virtual machine {0}' -f $VMName)
		$VMObject = Get-WmiObject -Namespace root\virtualization\v2 -Class Msvm_ComputerSystem -Filter "ElementName = '$VMName'"
		if($VMObject -eq $null)
		{
			throw('Virtual machine {0} not found on computer {1}' -f $VMName, $ComputerName)
		}
		Write-Verbose -Message ('Verifying that {0} is off...' -f $VMName)
		$OriginalState = $VMObject.EnabledState
		if($OriginalState -ne 3)
		{
			if($OriginalState -eq 2 -band ($Force.ToBool() -bor $PSCmdlet.ShouldProcess($VMName, 'Shut down')))
			{
				$ShutdownComponent = $VMObject.GetRelated('Msvm_ShutdownComponent')
				Write-Verbose -Message 'Initiating shutdown...'
				Process-WMIJob -WmiResponse $ShutdownComponent.InitiateShutdown($true, 'Change BIOSGUID') -WmiClassPath $ShutdownComponent.ClassPath -MethodName 'InitiateShutdown' -VMName $VMName -ComputerName $ComputerName -ErrorAction Stop
				# the InitiateShutdown function completes as soon as the guest's integration services respond; it does not wait for the power state change to complete
				Write-Verbose -Message ('Waiting for virtual machine {0} to shut down...' -f $VMName)
				$TimeoutCounterStarted = [datetime]::Now
				$TimeoutExpiration = [datetime]::Now + [timespan]::FromSeconds($Timeout)
				while($VMObject.EnabledState -ne 3)
				{
					$ElapsedPercent = [UInt32]((([datetime]::Now - $TimeoutCounterStarted).TotalSeconds / $Timeout) * 100)
					if($ElapsedPercent -ge 100)
					{
						throw('Timeout waiting for virtual machine {0} to shut down' -f $VMName)
					}
					else
					{
						Write-Progress -Activity ('Waiting for virtual machine {0} on {1} to stop' -f $VMName, $ComputerName) -Status ('{0}% timeout expiration' -f ($ElapsedPercent)) -PercentComplete $ElapsedPercent
						Start-Sleep -Milliseconds 250
						$VMObject.Get()
					}
				}
			}
			elseif($OriginalState -ne 2)
			{
				throw('Virtual machine must be turned off to replace the BIOS GUID. It is not in a state this script can work with.' -f $VMName)
			}
		}
		Write-Verbose -Message ('Retrieving all current settings for virtual machine {0}' -f $VMName)
		$CurrentSettingsDataCollection = $VMObject.GetRelated('Msvm_VirtualSystemSettingData')
		Write-Verbose -Message 'Extracting the settings data object from the settings data collection object...'
		$CurrentSettingsData = $null
		foreach($SettingsObject in $CurrentSettingsDataCollection)
		{
			 $CurrentSettingsData = [System.Management.ManagementObject]($SettingsObject)
		}
		
		if($NewID -eq $null)
		{
			Write-Verbose 'Generating new GUID...'
			$NewID = [System.Guid]::NewGuid()
		}


		$OriginalGUID = $CurrentSettingsData.BIOSGUID
		Write-Verbose -Message ('Orginal BIOS GUID: {0}' -f $OriginalGUID)
		Write-Verbose -Message 'Changing BIOSGUID in data object...'
		$CurrentSettingsData['BIOSGUID'] = "{$($NewID.Guid.ToUpper())}"
		Write-Verbose -Message ('New BIOS GUID: {0}' -f $CurrentSettingsData.BIOSGUID)
		Write-Verbose -Message 'Assigning modified data object as parameter for ModifySystemSettings function...'
		$ModifySystemSettingsParams['SystemSettings'] = $CurrentSettingsData.GetText([System.Management.TextFormat]::CimDtd20)
		if($Force.ToBool() -bor $PSCmdlet.ShouldProcess($VMName, ('Change BIOSGUID from {0} to {1}' -f $OriginalGUID, "{$($NewID.Guid.ToUpper())}")))
		{
			Write-Verbose -Message ('Instructing Virtual Machine Management Service to modify settings for virtual machine {0}' -f $VMName)
			Process-WMIJob -WmiClassPath $VMMS.ClassPath ($VMMS.InvokeMethod('ModifySystemSettings', $ModifySystemSettingsParams, $null))
			Process-WMIJob -WmiResponse ($VMMS.InvokeMethod('ModifySystemSettings', $ModifySystemSettingsParams, $null)) -WmiClassPath $VMMS.ClassPath -MethodName 'ModifySystemSettings' -VMName $VMName -ComputerName $ComputerName
		}
		$VMObject.Get()
		if($OriginalState -ne $VMObject.EnabledState)
		{
			Write-Verbose -Message ('Returning {0} to its original running state.' -f $VMName)
			Process-WMIJob -WmiResponse $VMObject.RequestStateChange($OriginalState) -WmiClassPath $VMObject.ClassPath -MethodName 'RequestStateChange' -VMName $VMName -ComputerName $ComputerName -ErrorAction Stop
		}
	}
}
