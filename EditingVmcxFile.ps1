#Requires -RunAsAdministrator

<#

Editing a .VMCX file
https://blogs.msdn.microsoft.com/virtual_pc_guy/2017/04/18/editing-a-vmcx-file/

#>

function Test-CimMethodSucceeded
{
    [OutputType([bool])]
    [CmdletBinding()]
    param (
        [System.Management.ManagementBaseObject] $Result
    )

    # ReturnValue = 0    : Completed with no error
    # ReturnValue = 4096 : Method parameters checked - job started
    return ($Result.ReturnValue -eq 0) -or ($Result.ReturnValue -eq 4096)
}


function Wait-JobComplete
{
    [OutputType([void])]
    [CmdletBinding()]
    param (
        [System.Management.ManagementBaseObject] $Result
    )

    # Already completed.
    if ($Result.ReturnValue -eq 0) { return }
    
    if ($result.ReturnValue -ne 4096)
    {
        throw ('The result is indicating errors. {0}' -f $result.ReturnValue)
    }

    # Get the job object.
    $job = [wmi] $result.Job

    # JobState = 3 : Starting
    # JobState = 4 : Running
    while (($job.JobState -eq 3) -or ($job.JobState -eq 4))
    {
        Write-Progress -Activity 'Waiting job complete...' -PercentComplete $job.PercentComplete
        Start-Sleep -Seconds 1

        # Reflesh the job object.
        $job = [wmi]$result.Job
    }
    Write-Progress -Activity 'Job completed' -Completed

    # JobState = 7 : Completed
    if ($job.JobState -ne 7)
    {
        if ($job.ErrorCode -eq 32773)
        {
            throw ('Failed the job due to invalid parameter ({0}). {1}' -f $job.ErrorCode, $job.ErrorDescription)
        }
        else
        {
            throw ('Failed the job with {0}. {1}' -f $job.ErrorCode, $job.ErrorDescription)
        }
    }
}


$vmConfigFilePathOriginal = 'D:\VM\HyperV-Socket\Virtual Machines\D62EFEF5-0E97-4F51-BB0E-2D858275CDBF.vmcx'
$vmConfigFilePathModified = 'D:\VM\HyperV-Socket.mod'


##
## Load a VM definition.
##

# Retrieve the virtual system management service.
$vsms = Get-WmiObject -Namespace 'root\virtualization\v2' -Class 'Msvm_VirtualSystemManagementService'

# Load the .vmcx file.
$result = $vsms.ImportSystemDefinition($vmConfigFilePathOriginal, $null, $true)
if (-not (Test-CimMethodSucceeded -Result $result))
{
    throw ('Failed ImportSystemDefinition method with {0}.' -f $result.ReturnValue)
}
Wait-JobComplete -Result $result

# Retrieve a VM definition.
$vmDefinition = [wmi]$result.ImportedSystem

# Get the settings data of VM.
$vmDefinitionSettingData = $vmDefinition.GetRelated(
                               'Msvm_VirtualSystemSettingData',
                               'Msvm_SettingsDefineState',
                               $null,
                               $null,
                               'SettingData',
                               'ManagedElement',
                               $false,
                               $null
                           ) | Select-Object -First 1


##
## Change the settings of VM.
##

#
# Modify the VM name.
#

$vmDefinitionSettingData.ElementName = 'NewVMName1'

# Apply the changes.
$vmDefinitionSettingDataText = $vmDefinitionSettingData.GetText([System.Management.TextFormat]::CimDtd20)
$result = $vsms.ModifySystemSettings($vmDefinitionSettingDataText)
if (-not (Test-CimMethodSucceeded -Result $result))
{
    throw ('Failed ModifySystemSettings method with {0}.' -f $result.ReturnValue)
}
Wait-JobComplete -Result $result

#
# Modify the memory settings.
#

# Get the memory settings data of VM.
$vmDefinitionMemorySettingData = $vmDefinitionSettingData.GetRelated(
                                     'Msvm_MemorySettingData',
                                     'Msvm_VirtualSystemSettingDataComponent',
                                     $null,
                                     $null,
                                     'PartComponent',
                                     'GroupComponent',
                                     $false,
                                     $null
                                 ) | Select-Object -First 1

$vmDefinitionMemorySettingData.DynamicMemoryEnabled = $false  # Dynamic memory enable/disable.
$vmDefinitionMemorySettingData.Reservation = 256              # Minimum RAM.
$vmDefinitionMemorySettingData.Limit = 1048576                # Maximum RAM.
$vmDefinitionMemorySettingData.VirtualQuantity = 2048         # RAM. Initial memory available at startup.
$vmDefinitionMemorySettingData.TargetMemoryBuffer = 30        # Memory buffer.
$vmDefinitionMemorySettingData.Weight = 5000                  # Memory weight.

# Apply the changes.
$vmDefinitionMemorySettingDataText = $vmDefinitionMemorySettingData.GetText([System.Management.TextFormat]::CimDtd20)
$result = $vsms.ModifyResourceSettings($vmDefinitionMemorySettingDataText)
if (-not (Test-CimMethodSucceeded -Result $result))
{
    throw ('Failed ModifyResourceSettings method with {0}.' -f $result.ReturnValue)
}
Wait-JobComplete -Result $result

#
# Export the modified VM definition.
#

# Get the export settings data of VM definition.
$vmDefinitionExportSettingData = $vmDefinition.GetRelated(
                                     'Msvm_VirtualSystemExportSettingData',
                                     'Msvm_SystemExportSettingData',
                                     $null,
                                     $null,
                                     'SettingData',
                                     'ManagedElement',
                                     $false,
                                     $null
                                 ) | Select-Object -First 1

# ExportNoSnapshots. No snapshots will be exported with the virtual machine.
$vmDefinitionExportSettingData.CopySnapshotConfiguration = 1

# The virtual machine run-time information will not be copied.
$vmDefinitionExportSettingData.CopyVmRuntimeInformation = $false

# The virtual machine storage will not be copied.
$vmDefinitionExportSettingData.CopyVmStorage = $false

# A sub-directory will be created.
$vmDefinitionExportSettingData.CreateVmExportSubdirectory = $true

# Export the modified VM definition to a new file.
$vmDefinitionExportSettingDataText = $vmDefinitionExportSettingData.GetText([System.Management.TextFormat]::CimDtd20)
$result = $vsms.ExportSystemDefinition($vmDefinition, $vmConfigFilePathModified, $vmDefinitionExportSettingDataText)
if (-not (Test-CimMethodSucceeded -Result $result))
{
    throw ('Failed ExportSystemDefinition method with {0}.' -f $result.ReturnValue)
}
Wait-JobComplete -Result $result


Write-Host ('Original: {0}' -f $vmConfigFilePathOriginal)
Write-Host ('Modified: {0}' -f $vmConfigFilePathModified)

<#

Hyper-V WMI: Rich Error Messages for Non-Zero ReturnValue (no more 32773, 32768, 32700…)
https://blogs.msdn.microsoft.com/taylorb/2008/06/19/hyper-v-wmi-rich-error-messages-for-non-zero-returnvalue-no-more-32773-32768-32700/


$vsms = Get-WmiObject -Namespace 'root\virtualization\v2' -Class 'Msvm_VirtualSystemManagementService'
$wmiclass = [wmiclass]$vsms.ClassPath
$wmiclass | fl *

$wmiclass.PSBase.Options.UseAmendedQualifiers = $true
$methodQualifiers = $wmiclass.PSBase.Methods['ModifySystemSettings'].Qualifiers
$indexOfError = [System.Array]::IndexOf($methodQualifiers['valueMap'].Value, [string]4096)
$methodQualifiers['Values'].Value[$indexOfError]

#>
