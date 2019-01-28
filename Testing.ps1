function test-export {
    $Content = Get-content -ReadCount 0 -Raw -Encoding Byte -LiteralPath "C:\Users\c.magnuson\OneDrive - tervis\Downloads\Packlist download testing\WebToPrintSamples\WebToPrintBach2\WhiteInkMaskAsGreyColorAsPNG\16DWT 5031857a-4498-48fe-9546-46fd88b0fe1e.pdf"
    $Content
}

function test-export2 {
    Get-content -Raw -Encoding Byte -LiteralPath "C:\Users\c.magnuson\OneDrive - tervis\Downloads\Packlist download testing\WebToPrintSamples\WebToPrintBach2\WhiteInkMaskAsGreyColorAsPNG\16DWT 5031857a-4498-48fe-9546-46fd88b0fe1e.pdf"
}

$Instances = Get-InDesignServerInstance
$InDesignServerInstance = $Instances[0]

$InDesignServerInstance = [PSCustomObject]@{
    ComputerName = (Get-TervisInDesignServerComputerName)
    Port = 8080
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
        }

New-Object -TypeName "InDesignServer$($InDesignServerInstance.Port).RunScriptParameters"


$SessionID = $InDesignServerInstance.WebServiceProxy.BeginSession()
$InDesignServerInstance.WebServiceProxy | gm
$InDesignServerInstance.WebServiceProxy.EndSession($SessionID)

Invoke-TervisInDesignServerRunScript -ScriptText "`$.writeln('test')" -ScriptLanguage "javascript"