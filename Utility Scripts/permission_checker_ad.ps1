Function listaDiretorios {
    $lista = Get-ChildItem '\\ad.ufrgs.br\LITORAL' | where-object {($_.PsIsContainer)} | Get-ACL
    $array = @()
    listaDiretorioRecursivo($lista, '\\ad.ufrgs.br\LITORAL', $array)
}

Function listaDiretorioRecursivo {
    Param($lista, $caminho, $array)
    $caminho = $lista.Path
    Write-Output $caminho
    ForEach($item in $caminho) {
        if ($caminho) {
            Write-Output $path.Substring( $path.LastIndexOf('\')[0])
        }
        #$childrens = Get-ChildItem "$caminho$pasta" | where-object {($_.PsIsContainer)} | Get-ACL
        #$array += $childrens
        #listaDiretorioRecursivo($childrens, "$caminho$pasta", $array)
    }
}

listaDiretorios