$computers = Get-ADComputer -Filter {Name -Like "LIT*" -and Enabled -eq $true}
$lista = @()
foreach($computer in $computers){
    $nameStr = $computer.DistinguishedName
    $ouPosition = $nameStr.IndexOf('OU=')
    if ($ouPosition -gt 0) {
        $ouStrBase = $nameStr.Substring($ouPosition + 3)
        $ou = $ouStrBase.Substring(0, $ouStrBase.IndexOf(','))
        $name = $nameStr.Substring(3, $nameStr.IndexOf(',OU') - 3)
        $patrimonio = $name -replace '\D+(\d+)','$1'
        if ($patrimonio.Length -gt 6) {
            $patrimonio = $patrimonio.Substring(1)
        }
        $obj = [PSCustomObject]@{
            Patrimonio = $patrimonio
            Local = $ou
            Nome = $name
        }
        $lista += $obj
    }
}

Write-Output $lista | Export-CSV -Path "lista-por-setor.csv" -Delimiter ";" -NoTypeInformation