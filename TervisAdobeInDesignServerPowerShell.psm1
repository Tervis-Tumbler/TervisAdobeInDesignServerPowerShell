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
        $Script:InDesignServerInstances = [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8080
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8081
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8082
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8083
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8084
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8085
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8086
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8087
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8088
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8089
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8090
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8091
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8092
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8093
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8094
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8095
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8096
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8097
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8098
        },
        [PSCustomObject]@{
            ComputerName = (Get-TervisInDesignServerComputerName)
            Port = 8099
        } |
        Add-Member -PassThru -MemberType ScriptProperty -Name WebServiceProxy -Value {
            $This | Add-Member -Force -MemberType NoteProperty -Name WebServiceProxy -Value $(
                $Proxy = New-WebServiceProxy -Class "InDesignServer$($This.Port)" -Namespace "InDesignServer$($This.Port)" -Uri (
                    Get-InDesignServerWSDLURI -ComputerName $This.ComputerName -Port $This.Port
                )
                $Proxy.Url = "http://$($This.ComputerName):$($This.Port)/"
                $Proxy
            )
            $This.WebServiceProxy
        } |
        Add-Member -MemberType NoteProperty -Name Locked -Value $False -Force -PassThru
    }
    $Script:InDesignServerInstances
}

function Invoke-TervisInDesignServerRunScript {
    param (
        $ScriptText,
        $ScriptLanguage,
        $ScriptFile,
        $ScriptArgs,
        $Port,
        $InDesignServerInstances = (Get-InDesignServerInstance),
        [Switch]$AsRSJob
    )

    $InDesignServerInstance = Select-InDesignServerInstance -InDesignServerInstances $InDesignServerInstances -SelectionMethod Port -Port $Port

    $Proxy = $InDesignServerInstance.WebServiceProxy
    $RunScriptParametersProperty = $PSBoundParameters | ConvertFrom-PSBoundParameters -ExcludeProperty InDesignServerInstances, Port, AsRSJob -AsHashTable
    $Parameter = New-Object -TypeName "InDesignServer$($InDesignServerInstance.Port).RunScriptParameters" -Property $RunScriptParametersProperty
    $ErrorString = ""
    $Results = New-Object -TypeName "InDesignServer$($InDesignServerInstance.Port).Data"

    if (-not $AsRSJob) {
        #$SessionID = $Proxy.BeginSession()
        $Response = $Proxy.RunScript($Parameter, [Ref]$ErrorString, [Ref]$Results)
        #$Proxy.EndSession($SessionID)
        
        if ($ErrorString) { Write-Error -Message $ErrorString }
        if ($Response.result) { Write-Verbose -Message $Response.result }
        $Results
    
        $InDesignServerInstance.Locked = $False
    } elseif ($AsRSJob) {
        Start-RSJob -ScriptBlock {
            $Response = $($Using:Proxy).RunScript($Using:Parameter, [Ref]$Using:ErrorString, [ref]$Using:Results)
            if ($Using:ErrorString) { Write-Error -Message $Using:ErrorString }
            if ($Response.result) { Write-Verbose -Message $Response.result }
            $Using:Results
        
            $($Using:InDesignServerInstance).Locked = $False
        }
    }
}

function Select-InDesignServerInstance {
    param (
        [Parameter(Mandatory)]$InDesignServerInstances,
        [ValidateSet("Lock","Random","Port")][Parameter(Mandatory)]$SelectionMethod,
        [String]$Port
    )
    if ($SelectionMethod -eq "Lock") {
        Lock-TervisInDesignServerInstance -InDesignServerInstances $InDesignServerInstances
    } elseif ($SelectionMethod -eq "Random") {
        $RandomIndex = (Get-Random) % $InDesignServerInstances.Count
        $InDesignServerInstances[$RandomIndex]
    } elseif ($SelectionMethod -eq "Port") {
        $Index = $Port.Substring($Port.length - 2)
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
    $InDesignServerInstances | New-InDesignServerInstance
}

function Get-TervisInDesignServerInstanceListeningPorts {
    $InDesignServerInstances = (Get-InDesignServerInstance)
    $CIMSession = New-CimSession -ComputerName $Script:InDesignServerComputerName
    Get-NetTCPConnection -State Listen -LocalPort $InDesignServerInstances.Port -CimSession $CIMSession
}

function Invoke-TervisAdobeInDesignServerProvision {
    Invoke-ApplicationProvision -ApplicationName InDesign -EnvironmentName Infrastructure
    $ComputerName = "INF-InDesign02"
    Disable-InternetExplorerESC -ComputerName $ComputerName
    
    Read-Host "\\tervis.prv\applications\Installers\Adobe\Adobe InDesign CC Server 2019\Set-up.exe"

    Get-TervisAdobeProvisioningToolkitSerializeInDesignServerProvisioningXML -OutPath $RemotePath
    Invoke-AdobeProvisioningToolkitSerializeLoad -ProvisioningXMLFilePath $RemotePath\prov.xml -ComputerName $ComputerName
    
    Set-InDesignServerComputerName -ComputerName $ComputerName
    Install-InDesignServerMMCSnapIn
    Install-InDesignServerService
    Read-Host "Set InDesignServerService x64 service to run as Local System account"
    Invoke-TervisInDesignServerInstanceProvision

    Install-TervisWebToPrintPolaris -ComputerName $ComputerName
    
    Copy-TervisWebToPinrtInDesignServerTemplate -ComputerName $ComputerName
    #Copy PDF preset
}
