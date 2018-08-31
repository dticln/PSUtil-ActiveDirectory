

Function DisplayInBytes($num) 
{
    $suffix = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
    $index = 0
    while ($num -gt 1kb) 
    {
        $num = $num / 1kb
        $index++
    } 

    return "{0:N1} {1}" -f $num, $suffix[$index]
}

Function exportarEmExcel {
    Param($arquivos)
    Try {
        $excel = New-Object -com 'excel.application'
        $pastas = $excel.Workbooks.Add()
        $tabela = $pastas.Worksheets.Item(1)
        $tabela.Name = 'Relatorio'
        $tabela.Cells.Item(1, 1) = 'Tamanho'
        $tabela.Cells.Item(1, 2) = 'Nome'
        $tabela.Cells.Item(1, 3) = 'Local'
        $i = 2
        foreach($arquivo in $arquivos) {
            $tabela.Cells.Item($i, 1) = DisplayInBytes $arquivo.Length 
            $tabela.Cells.Item($i, 2) = $arquivo.Name
            $tabela.Cells.Item($i, 3) = $arquivo.Directory
            $i++
        }
        $pastas.SaveAs("$PSScriptRoot\relatorio_tamanho.xlsx")
        $excel.Quit()
        Write-Output "Relatorio de tamanho salvo em '$PSScriptRoot\relatorio_tamanho.xlsx'."
        finalizar
    } Catch {
        Write-Output 'Falha ao salvar o arquivo de relatorio.'
        finalizar
    }
}

##############################
#.SYNOPSIS
# Finaliza a execução do script
##############################
Function finalizar {
    Write-Output 'Pressione qualquer tecla para finalizar.'
    cmd /c pause | out-null
    break
}

$recursive = Get-ChildItem -Path "\\ad.ufrgs.br\LITORAL\Direção Geral\Direção Acadêmica\Biblioteca" -Recurse 
$onlyFiles = $recursive | Where-Object -Property Attributes -eq 'Archive'
$sorted = $onlyFiles | Sort-Object Length -Descending | Select-Object Length, Name, Directory
exportarEmExcel $sorted