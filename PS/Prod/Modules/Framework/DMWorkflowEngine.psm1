<#
.SYNOPSIS
    Desktop Management Workflow Engine Module
    
.DESCRIPTION
    Generic workflow orchestrator that executes steps defined in workflow configuration files.
    Provides dynamic module loading and function execution based on workflow definitions.
    
.NOTES
    Version: 2.0.0
    Author: Desktop Management Team
#>

# Import required modules
Using Module .\DMLogger.psm1

<#
.SYNOPSIS
    Executes a workflow based on configuration file.
    
.DESCRIPTION
    Loads workflow configuration and executes all enabled steps in order.
    
.PARAMETER WorkflowFile
    Path to workflow configuration file (.psd1)
    
.PARAMETER Context
    Hashtable containing workflow context (UserInfo, ComputerInfo, etc.)
    
.OUTPUTS
    Boolean - true if workflow completed (even with some step failures)
    
.EXAMPLE
    $Context = @{
        UserInfo = $User
        ComputerInfo = $Computer
        JobType = 'Logon'
    }
    Invoke-DMWorkflow -WorkflowFile ".\Config\Workflow-Logon.psd1" -Context $Context
#>
Function Invoke-DMWorkflow {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [String]$WorkflowFile,
        
        [Parameter(Mandatory=$True)]
        [Hashtable]$Context
    )
    
    Try {
        # Load workflow configuration
        If (-not (Test-Path -Path $WorkflowFile)) {
            Write-DMLog "Workflow Engine: Workflow file not found: $WorkflowFile" -Level Error
            Return $False
        }
        
        [Hashtable]$Workflow = Import-PowerShellDataFile -Path $WorkflowFile
        
        Write-DMLog "Workflow Engine: Loaded workflow: $($Workflow.JobType) - $($Workflow.Description)" -Level Verbose
        Write-DMLog "Workflow Engine: Total steps: $($Workflow.Steps.Count)" -Level Verbose
        Write-DMLog ""
        
        # Sort steps by Order
        [Array]$SortedSteps = $Workflow.Steps | Sort-Object -Property Order
        
        [Int]$TotalSteps = $SortedSteps.Count
        [Int]$EnabledSteps = ($SortedSteps | Where-Object { $_.Enabled -eq $True }).Count
        [Int]$ExecutedSteps = 0
        [Int]$SuccessfulSteps = 0
        [Int]$FailedSteps = 0
        [Int]$SkippedSteps = 0
        
        Write-DMLog "Workflow Engine: Enabled steps: $EnabledSteps of $TotalSteps" -Level Info
        Write-DMLog ""
        
        # Execute each step
        [String]$CurrentPhase = ""
        
        ForEach ($Step in $SortedSteps) {
            # Log phase header if changed
            If ($Step.Phase -ne $CurrentPhase) {
                $CurrentPhase = $Step.Phase
                Write-DMLog "========================================" -Level Info
                Write-DMLog "PHASE: $CurrentPhase" -Level Info
                Write-DMLog "========================================" -Level Info
                Write-DMLog ""
            }
            
            # Check if step is enabled
            If (-not $Step.Enabled) {
                Write-DMLog "--- $($Step.Name) [DISABLED] ---" -Level Verbose
                $SkippedSteps++
                Continue
            }
            
            Write-DMLog "--- $($Step.Name) ---" -Level Info
            
            $ExecutedSteps++
            
            # Execute the step
            [Boolean]$StepResult = Invoke-DMWorkflowStep -Step $Step -Context $Context
            
            If ($StepResult) {
                $SuccessfulSteps++
                Write-DMLog "$($Step.Name): Completed successfully" -Level Verbose
            } Else {
                $FailedSteps++
                Write-DMLog "$($Step.Name): Failed" -Level Warning
                
                # Check if we should continue on error
                If (-not $Step.ContinueOnError) {
                    Write-DMLog "Workflow Engine: Step failed and ContinueOnError is False, stopping workflow" -Level Error
                    Break
                }
            }
            
            Write-DMLog ""
        }
        
        # Workflow summary
        Write-DMLog "Workflow Engine: Execution Summary" -Level Info
        Write-DMLog "  Total Steps: $TotalSteps" -Level Info
        Write-DMLog "  Enabled: $EnabledSteps" -Level Info
        Write-DMLog "  Executed: $ExecutedSteps" -Level Info
        Write-DMLog "  Successful: $SuccessfulSteps" -Level Info
        Write-DMLog "  Failed: $FailedSteps" -Level Info
        Write-DMLog "  Skipped: $SkippedSteps" -Level Info
        Write-DMLog ""
        
        Return $True
    }
    Catch {
        Write-DMLog "Workflow Engine: Fatal error - $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Executes a single workflow step.
    
.DESCRIPTION
    Loads the module and calls the function with specified parameters.
    
.PARAMETER Step
    Step configuration hashtable
    
.PARAMETER Context
    Workflow context hashtable
    
.OUTPUTS
    Boolean - true if step succeeded
#>
Function Invoke-DMWorkflowStep {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [Hashtable]$Step,
        
        [Parameter(Mandatory=$True)]
        [Hashtable]$Context
    )
    
    Try {
        # Get module path (go up from Framework to Modules root)
        [String]$ModulePath = Join-Path (Split-Path $PSScriptRoot -Parent) $Step.Module
        
        # Import required modules first (if specified)
        If ($Step.ContainsKey('RequiredModules') -and $Null -ne $Step.RequiredModules) {
            ForEach ($RequiredModule in $Step.RequiredModules) {
                [String]$ReqModulePath = Join-Path (Split-Path $PSScriptRoot -Parent) $RequiredModule
                Try {
                    Import-Module $ReqModulePath -Force
                    Write-DMLog "Workflow Step: Imported required module: $RequiredModule" -Level Verbose
                } Catch {
                    Write-DMLog "Workflow Step: Warning - Could not import required module: $RequiredModule" -Level Verbose
                }
            }
        }
        
        # Import the main module
        If (-not (Test-Path -Path $ModulePath)) {
            Write-DMLog "Workflow Step: Module not found: $ModulePath" -Level Error
            Return $False
        }
        
        Import-Module $ModulePath -Force
        
        # Build parameter hashtable
        [Hashtable]$FunctionParams = @{}
        
        ForEach ($ParamKey in $Step.Parameters.Keys) {
            [String]$ParamValue = $Step.Parameters[$ParamKey]
            
            # Check if static value (prefix: "Static:")
            If ($ParamValue -like "Static:*") {
                [String]$StaticValue = $ParamValue.Replace("Static:", "")
                $FunctionParams[$ParamKey] = $StaticValue
            }
            # Otherwise, get from context
            ElseIf ($Context.ContainsKey($ParamValue)) {
                $FunctionParams[$ParamKey] = $Context[$ParamValue]
            }
            Else {
                Write-DMLog "Workflow Step: Warning - Context does not contain '$ParamValue'" -Level Verbose
            }
        }
        
        # Call the function
        [String]$FunctionName = $Step.Function
        
        If (-not (Get-Command -Name $FunctionName -ErrorAction SilentlyContinue)) {
            Write-DMLog "Workflow Step: Function not found: $FunctionName" -Level Error
            Return $False
        }
        
        # Execute with splatting
        [Boolean]$Result = & $FunctionName @FunctionParams
        
        Return $Result
    }
    Catch {
        Write-DMLog "Workflow Step Error: $($_.Exception.Message)" -Level Error
        Return $False
    }
}

<#
.SYNOPSIS
    Gets workflow configuration for a job type.
    
.DESCRIPTION
    Loads and returns the workflow configuration file for the specified job type.
    
.PARAMETER JobType
    Job type (Logon, Logoff, TSLogon, TSLogoff)
    
.OUTPUTS
    Hashtable - workflow configuration
    
.EXAMPLE
    $Workflow = Get-DMWorkflowConfig -JobType "Logon"
#>
Function Get-DMWorkflowConfig {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$True)]
        [ValidateSet('Logon', 'Logoff', 'TSLogon', 'TSLogoff')]
        [String]$JobType
    )
    
    Try {
        [String]$WorkflowFile = Join-Path $PSScriptRoot "..\..\Config\Workflow-$JobType.psd1"
        
        If (-not (Test-Path -Path $WorkflowFile)) {
            Write-DMLog "Get Workflow Config: File not found: $WorkflowFile" -Level Error
            Return $Null
        }
        
        Return Import-PowerShellDataFile -Path $WorkflowFile
    }
    Catch {
        Write-DMLog "Get Workflow Config: Error - $($_.Exception.Message)" -Level Error
        Return $Null
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Invoke-DMWorkflow',
    'Invoke-DMWorkflowStep',
    'Get-DMWorkflowConfig'
)

