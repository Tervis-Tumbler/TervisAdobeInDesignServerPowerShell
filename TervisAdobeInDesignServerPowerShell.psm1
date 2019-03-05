function Get-TervisInDesignServerComputerName {
    if (-not $Script:InDesignServerComputerName) {
        $Script:InDesignServerComputerName = Get-TervisApplicationNode -ApplicationName InDesign -IncludeCredential:$False -IncludeIPAddress:$false |
        Select-Object -ExpandProperty ComputerName |
        Select-Object -First 1 -Skip 1
    }
    $Script:InDesignServerComputerName
}

function Set-TervisInDesignServerComputerName {
    Set-InDesignServerComputerName -ComputerName $(Get-TervisInDesignServerComputerName)
}

Set-TervisInDesignServerComputerName

function Get-InDesignServerInstance {
    if (-not $Script:InDesignServerInstances) {
        $Script:InDesignServerInstances = 8080..8099 |
        ForEach-Object {
            New-InDesignServerInstance -ComputerName (Get-TervisInDesignServerComputerName) -Port $_
        }
    }
    $Script:InDesignServerInstances
}

function Invoke-TervisInDesignServerRunScript {
    param (
        $ScriptText,
        $ScriptLanguage,
        $ScriptFile,
        $ScriptArgs,
        $InDesignServerInstances = (Get-InDesignServerInstance),
        [Switch]$AsRSJob
    )

    $InDesignServerInstance = Select-InDesignServerInstance -InDesignServerInstances $InDesignServerInstances -SelectionMethod Random
    $RunScriptParametersProperty = $PSBoundParameters | ConvertFrom-PSBoundParameters -ExcludeProperty InDesignServerInstances, AsRSJob -AsHashTable

    if (-not $AsRSJob) {
        $Results = Invoke-InDesignServerRunScript @RunScriptParametersProperty -InDesignServerInstance $InDesignServerInstance
        [PSCustomObject]@{
            Results = $Results
            InDesignServerInstance = $InDesignServerInstance
        }
    
        $InDesignServerInstance.Locked = $False
    } elseif ($AsRSJob) {
        Start-RSJob -ScriptBlock {
            $Parameters = $Using:RunScriptParametersProperty
            $Results = Invoke-InDesignServerRunScript @Parameters -InDesignServerInstance $Using:InDesignServerInstance

            [PSCustomObject]@{
                Results = $Results
                InDesignServerInstance = $Using:InDesignServerInstance
            }

            $($Using:InDesignServerInstance).Locked = $False
        }
    }
}

function Select-InDesignServerInstance {
    param (
        $InDesignServerInstances = (Get-InDesignServerInstance),
        [ValidateSet("Lock","Random","Port")][Parameter(Mandatory)]$SelectionMethod,
        [String]$Port
    )
    if ($SelectionMethod -eq "Lock") {
        Lock-TervisInDesignServerInstance -InDesignServerInstances $InDesignServerInstances
    } elseif ($SelectionMethod -eq "Random") {
        $RandomIndex = (Get-Random) % $InDesignServerInstances.Count
        $InDesignServerInstances[$RandomIndex]
    } elseif ($SelectionMethod -eq "Port") {
        #https://stackoverflow.com/questions/345187/math-mapping-numbers
        $IndexSelector = $Port.Substring($Port.length - 2)
        $PortSuffixLowerBound = 00
        $PortSuffixUpperBound = 99
        $ArrayIndexLowerBound = 0
        $ArrayIndexUpperBound = $InDesignServerInstances.Count - 1
        $Index = (
            $IndexSelector - $PortSuffixLowerBound
        ) / (
        $PortSuffixUpperBound - $PortSuffixLowerBound
        ) * (
            ($ArrayIndexUpperBound - $ArrayIndexLowerBound) + $ArrayIndexLowerBound
        )
        $InDesignServerInstances[$Index]
    }
}

function Lock-TervisInDesignServerInstance {
    param (
        $InDesignServerInstances
    )
    while (-not $InDesignServerInstance) {
        Lock-Object -InputObject $InDesignServerInstances -ScriptBlock {
            $InDesignServerInstance = $InDesignServerInstances |
            Where-Object {-not $_.Locked} |
            Select-Object -First 1
            
            if ($InDesignServerInstance) {
                $InDesignServerInstance.Locked = $true
                $InDesignServerInstance    
            } else {
                Start-Sleep -Seconds 1
            }
        }
    }
}

function Invoke-TervisInDesignServerInstanceProvision {
    $InDesignServerInstances = (Get-InDesignServerInstance)
    $InDesignServerInstances | Invoke-InDesignServerInstanceProvision
}

function Get-TervisInDesignServerInstanceListeningPorts {
    $InDesignServerInstances = (Get-InDesignServerInstance)
    $CIMSession = New-CimSession -ComputerName $Script:InDesignServerComputerName
    Get-NetTCPConnection -State Listen -LocalPort $InDesignServerInstances.Port -CimSession $CIMSession
}

function Invoke-TervisAdobeInDesignServerProvision {
    Invoke-ApplicationProvision -ApplicationName InDesign -EnvironmentName Infrastructure
    $ComputerName = Get-TervisInDesignServerComputerName
    Disable-InternetExplorerESC -ComputerName $ComputerName
    
    Read-Host "\\tervis.prv\applications\Installers\Adobe\Adobe InDesign CC Server 2019\Set-up.exe"

    Get-TervisAdobeProvisioningToolkitSerializeInDesignServerProvisioningXML -OutPath $RemotePath
    Invoke-AdobeProvisioningToolkitVolumeSerialize -ProvisioningXMLFilePath $RemotePath\prov.xml -ComputerName $ComputerName
    
    Set-InDesignServerComputerName -ComputerName $ComputerName
    Install-InDesignServerMMCSnapIn
    Install-InDesignServerService
    Read-Host "Set InDesignServerService x64 service to run as Local System account"
    Invoke-TervisInDesignServerInstanceProvision

    Install-TervisWebToPrintPolaris -ComputerName $ComputerName
    
    Copy-TervisWebToPinrtInDesignServerTemplate -ComputerName $ComputerName
    Copy-TervisWebToPrintInDesingServerJobOptions -ComputerName $ComputerName
}

function Get-TervisInDesignServerWireSharkeCaptureFilter {
    $Instances = Get-InDesignServerInstance
    $Filters = $Instances.Port |
    % {
        "tcp port $_"
    } 
    $Filters -join " or "
}