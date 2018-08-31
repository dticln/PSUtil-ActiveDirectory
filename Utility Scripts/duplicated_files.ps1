##############################
#.SYNOPSIS
# Finaliza a execuï¿½ï¿½o do script
##############################
Function finalizar {
    Write-Output 'Pressione qualquer tecla para finalizar.'
    cmd /c pause | out-null
    break
}

Function exportarEmExcel {
    Param($list)
    Try {
        $gray = 0xdcdcdc
        $lowGray = 0xf5f5f5
        $curColor = $gray
        $excel = New-Object -com 'excel.application'
        $pastas = $excel.Workbooks.Add()
        $tabela = $pastas.Worksheets.Item(1)
        $tabela.Name = 'Duplicatas'
        $tabela.Cells.Item(1, 1) = 'Identificador único'
        $tabela.Cells.Item(1, 2) = 'Nome do arquivo'
        $tabela.Cells.Item(1, 3) = 'Caminho'
        $i = 2
        $lastHash = $list[0];
        foreach($item in $list) {
            $tabela.Cells.Item($i, 1) = $item.HashID
            $tabela.Cells.Item($i, 2) = $item.Name
            $tabela.Cells.Item($i, 3) = $item.Path
            if ($lastHash -ne $item.HashID){
                $curColor = @{$true=$gray;$false=$lowGray}[$curColor -eq $lowGray]
            }
            $lastHash = $item.HashID
            $tabela.Cells.Item($i, 1).Interior.Color = $curColor
            $tabela.Cells.Item($i, 2).Interior.Color = $curColor
            $tabela.Cells.Item($i, 3).Interior.Color = $curColor
            $i++
        }
        $pastas.SaveAs("$PSScriptRoot\relatorio.xlsx")
        $excel.Quit()
        Write-Output "Relatorio de duplicatas salvo em '$PSScriptRoot\relatorio.xlsx'."
        finalizar
    } Catch {
        $ex = $_.Exception.Message
        Write-Output "Falha ao salvar o arquivo de relatorio: $ex"
        finalizar
    }
}

$list = Get-ChildItem -Path "\\ad.ufrgs.br\LITORAL\Direção Geral\Direção Acadêmica\Biblioteca" -Recurse 
$found = @{}
foreach($item in $list){
    $hash = Get-FileHash -Path $item.PSPath -ErrorAction SilentlyContinue
    if ($hash.Hash) {
        if(!$found.ContainsKey($hash.Hash)) {
            $found[$hash.Hash] = @()
            $found[$hash.Hash] += $item
        } else {
            $found[$hash.Hash] += $item
            $double = $hash.Hash
            $duplicated = $found[$hash.Hash]
            Write-Output "O arquivo '$double' está presente em $($duplicated.Length) diretórios."
        }
    } 
}

$export = @()
foreach($key in $found.Keys){
    if($found[$key].Length -gt 1) {
        foreach($file in $found[$key]) { 
            $export += [PSCustomObject]@{
                HashID = $key
                Name = $file.Name
                Path = $file.FullName
            }
        }
    }
}

exportarEmExcel($export)
 
Get-ChildItem -Path "\\ad.ufrgs.br\LITORAL\Direção Geral\Assessoria das Direções" -Recurse | Sort-Object 
