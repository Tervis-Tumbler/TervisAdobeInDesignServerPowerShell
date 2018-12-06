$Script:InDesignServerComputerName = Get-TervisApplicationNode -ApplicationName InDesign -IncludeCredential:$False -IncludeIPAddress:$false |
Select-Object -ExpandProperty ComputerName

function Set-TervisInDesignServerComputerName {
    Set-InDesignServerComputerName -ComputerName $Script:InDesignServerComputerName
}

$Script:InDesignServerInstances = [PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8080
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8081
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8082
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8083
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8084
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8085
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8086
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8087
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8088
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8089
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8090
},
[PSCustomObject]@{
    ComputerName = $InDesignServerComputerName
    Port = 8091
} |
Add-Member -PassThru -MemberType ScriptProperty -Name WebServiceProxy -Value {
    $This | Add-Member -Force -MemberType NoteProperty -Name WebServiceProxy -Value $(
        $Proxy = New-WebServiceProxy -Class InDesignServer -Namespace InDesignServer -Uri (
            Get-InDesignServerWSDLURI -ComputerName $This.ComputerName -Port $This.Port
        )
        $Proxy.Url = "http://$($This.ComputerName):$($This.Port)/"
        $Proxy
    )
    $This.WebServiceProxy
} |
Add-Member -MemberType NoteProperty -Name Locked -Value $False -Force -PassThru

function Get-InDesignServerInstance {
    $Script:InDesignServerInstances
}

function Invoke-TervisInDesignServerRunScript {
    param (
        $ScriptText,
        $ScriptLanguage,
        $ScriptFile,
        $ScriptArgs,
        $InDesignServerInstances = (Get-InDesignServerInstance)
    )

    $InDesignServerInstance = Lock-TervisInDesignServerInstance -InDesignServerInstances $InDesignServerInstances

    $Proxy = $InDesignServerInstance.WebServiceProxy
    $Parameter = New-Object -TypeName InDesignServer.RunScriptParameters -Property $PSBoundParameters
    $ErrorString = ""
    $Results = New-Object -TypeName InDesignServer.Data

    Start-RSJob -ScriptBlock {
        $Response = $($Using:Proxy).RunScript($Using:Parameter, [Ref]$Using:ErrorString, [ref]$Using:Results)
        Write-Error -Message $Using:ErrorString
        Write-Verbose -Message $Response.result
        $Using:Results
    
        $($Using:InDesignServerInstance).Locked = $False    
    }
}

function Lock-TervisInDesignServerInstance {
    param (
        $InDesignServerInstances
    )
    Lock-Object -InputObject $InDesignServerInstances -ScriptBlock {
        $InDesignServerInstance = $InDesignServerInstances |
        Where-Object {-not $_.Locked} |
        Select-Object -First 1
    
        $InDesignServerInstance.Locked = $true
        $InDesignServerInstance
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