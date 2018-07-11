#msg * "Política de grupo LITORAL-TesteDeRemocaoImpressora: excluindo impressoras Lexmark."
Try {
    $version = Get-WmiObject -Class Win32_OperatingSystem | ForEach-Object -MemberName Caption
} Catch {
    $version = "Windows 7 or Above"
}
$ipEnds = @(240, 238, 242, 243, 244, 245, 246, 248, 247, 241, 237)
If ($version -like '*10*') {
    foreach ($ipEnd in $ipEnds){
        $hostAdress = "143.54.196.$ipEnd"
        Get-Printer | where PortName -match $hostAdress | Remove-Printer
        Get-PrinterPort | where Name -match $hostAdress | Remove-PrinterPort
    }
} Else {
    $ports = @{}
    Get-WmiObject Win32_TCPIPPrinterPort | ForEach-Object {
        $ports.Add($_.Name, $_.HostAddress)
    }
    Get-WmiObject Win32_Printer | ForEach-Object {
        $currentWMIO = $_
        $printer = New-Object PSObject -Property @{
            "Name" = $currentWMIO.Name
            "DriverName" = $currentWMIO.DriverName
            "HostAddress" = $ports[$currentWMIO.PortName]
        }
        foreach ($ipEnd in $ipEnds){
            $hostAddress = "143.54.196.$ipEnd"
            If ($printer.HostAddress -match $hostAddress) {
                $port = Get-WmiObject Win32_TCPIPPrinterPort | Where { $_.Name -eq $currentWMIO.portname }
                $currentWMIO.Delete()
                $port.Delete()
            }
        }
    }
}