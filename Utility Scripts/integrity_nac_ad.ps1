##############################
#.SYNOPSIS
# Script de verificação de erros entre registros AD e relatórios NAC
#
#.DESCRIPTION
# O script realiza a verificação de incompatibilidades entre os registros
# obtidos através do domínio ad.ufrgs.br/litoral e os relatórios gerados pelo sistema NAC.
# Verificando problemas com erro no registro do patrimonio, duplicatas, registros incompatíveis,
# unidade organizacional incorreta e computadores patrimoniados sem registro.
##############################

##############################
#.SYNOPSIS
# Inicio do script de verificação de erros entre registros AD e relatórios NAC
#
#.DESCRIPTION
# Função com a lógica de verificação para cada patrimonio na lista do NAC
#
#.PARAMETER listaIps
# Nome da lista do NAC em formato .xls
#
#.EXAMPLE
# executaScript 'nome_do_arquivo.xls'
#
#.NOTES
# A ordem das construções lógicas impactam nos registros de erros, certos erros podem ser 
# mostrados duas vezes ou omitidos de acordo com a lógica da função
##############################
Function executaScript {
    Param($listaIps)
    $excel, $pasta = abreExcel $listaIps
    $planinha = $pasta.ActiveSheet
    Write-Output 'Verificando integridade entre os registros NAC e o AD.'
    $litoral = consultaAD
    $erros = @()
    # Para cada patrimonio na lista, faça:
    for($i = 2; $planinha.Cells.Item($i,5).Value(); $i++) {
        $patrimonio = $planinha.Cells.Item($i,5).Value()
        if($patrimonio -ne "-"){
            # Lógica interna do script, pode ser alterada
            # para obter relatórios diferentes
            $nomeNac = $planinha.Cells.Item($i,2).Value().ToUpper()
            if ($nomeNac  -notlike "*$patrimonio") {
                $erros += erroPatrimonio $nomeNac
            }
            $computadores = $litoral | Where-Object {$_.Name -Like "*$patrimonio" -and $_.Enabled -eq $true}
            $count = $computadores.Count
            if ($count -ge 1) {
                $erros += erroDuplicatas $nomeNac $computadores
            } elseif ($computadores){
                if ($nomeNac -ne $computadores.Name){
                    $erros += erroIncompativel $nomeNac $computadores
                }
                $ou = $computadores.DistinguishedName
                $setor = $computadores.Name.Substring(4,3)
                if (verificaOU $computadores.Name $ou $setor){
                    $erros += erroOU $nomeNac $ou
                }
            } else {
                $erros += erroSemRegistro($nomeNac)
            }
        }
    }
    $excel.Quit()
    exportarEmExcel $erros
}

##############################
#.SYNOPSIS
# Gera um registro de patrimonio
#
#.DESCRIPTION
# O erro de patrimonio ocorre quando o patrimonio registrado no NAC
# não confere com o patrimonio que consta no nome
#
#.PARAMETER nomeNac
# Nome de registro no NAC
##############################
Function erroPatrimonio {
    Param($nomeNac)
    $erro = criaErro $nomeNac 'O patrimonio no NAC nao confere com o nome.'
    return $erro
}

##############################
#.SYNOPSIS
# Gera um registro de duplicatas
#
#.DESCRIPTION
# O erro de patrimonio ocorre quando há mais de um registro no AD
# com o mesmo patrimonio no nome
#
#.PARAMETER nomeNac
# Nome de registro no NAC
#
#.PARAMETER computadores
# Lista de computadores com o patrimonio duplicado
##############################
Function erroDuplicatas {
    Param($nomeNac, $computadores)
    $erro = criaErro $nomeNac 'Dois ou mais computadores com o mesmo patrimonio.'
    foreach ($computador in $computadores) {
        $erro.Alias = "$($erro.Alias), AD:$($computador.Name)"
    }
    return $erro
}

##############################
#.SYNOPSIS
# Gera um registro de incompatibilidade
#
#.DESCRIPTION
# O erro de incompatibilidade ocorre quando há um nome no AD
# e outro nome registrado no NAC
#
#.PARAMETER nomeNac
# Nome de registro no NAC
#
#.PARAMETER computador
# Computador registrado no AD
##############################
Function erroIncompativel {
    Param($nomeNac, $computador)
    $erro = criaErro $nomeNac 'Nomes incompativeis entre AD e NAC.'
    $erro.Alias = "$($erro.Alias), AD:$($computador.Name)"
    return $erro
}

##############################
#.SYNOPSIS
# Gera um registro de OU 
#
#.DESCRIPTION
# O erro de OU ocorre quando o computador no AD não está registrado
# na OU correta. A verificação ocorre no nome do computador.
#
#.PARAMETER nomeNac
# Nome de registro no NAC
#
#.PARAMETER ou
# Unidade organizacional do objeto
##############################
Function erroOU {
    Param($nomeNac, $ou)
    $erro = criaErro $nomeNac 'O computador provavelmente esta na OU errada.'
    $erro.Alias = "$($erro.Alias), OU:$ou"
    return $erro
}

##############################
#.SYNOPSIS
# Gera um registro de ausência de registro
#
#.DESCRIPTION
# O erro de ausência de registro ocorre quando o computador no NAC não possui
# um registro no AD
##############################
Function erroSemRegistro {
    $erro = criaErro $nomeNac 'Nome nao possui registro no AD.'
    return $erro
}

##############################
#.SYNOPSIS
# Verifica se o computador está na OU
#
#.DESCRIPTION
# Verifica com base no nome do computador do AD se o computador possui uma OU vinculada
#
#.PARAMETER nome
# Nome de registro no AD
# Padrão do nome: LIT-[SETOR COM 3 DÍGITOS][PATRIMONIO COM 6 DÍGITOS]
#
#.PARAMETER ou
# Lista de OUs vinculadas ao objeto
# Padrão da OU: LIT-[SETOR]-[SETOR] (enquanto houverem setores)
#
#.PARAMETER setor
# Sigla do setor com três dígitos
##############################
Function verificaOU {
    Param($nome, $ou, $setor)
    return $nome -match "LIT-[a-zA-Z]{3}[0-9]{6}" -and $ou -notmatch "OU=(?:(?!OU=|$setor).)*$setor"
}

##############################
#.SYNOPSIS
# Executa pesquisa nos objetos do AD
#
#.PARAMETER nome
# Filtro de nome para consulta no AD, utilizado "*LIT*"
##############################
Function consultaAD {
    return Get-ADComputer -Filter {Name -Like "*LIT*"}
}

##############################
#.SYNOPSIS
# Cria um registro de erro genérico
#
#.PARAMETER nomeNac
# Nome de registro no NAC
#
#.PARAMETER incoerencia
# Registro escrito de erro
##############################
Function criaErro {
    Param($nomeNac, $incoerencia)
    $obj = [PSCustomObject]@{
        Nome = $nomeNac
        Incoerencia = $incoerencia
        Alias = "NAC:$nomeNac"
    }
    return $obj
}

##############################
#.SYNOPSIS
# Abre arquivo excel exportado do NAC
#
#.PARAMETER listaIps
# Nome do arquivo de listas de IPs do NAC
##############################
Function abreExcel {
    Param($listaIps)
    try {
        $excel = New-Object -com 'excel.application'
        $pasta = $excel.Workbooks.Open("$PSScriptRoot\$listaIps")
        return $excel, $pasta
    } catch {
        Write-Error "Erro ao abrir o arquivo Excel. Não foi possível localizar o arquivo '$PSScriptRoot\$listaIps'."
        break
    }
}

##############################
#.SYNOPSIS
# Exporta registros para um arquivo Excel
#
#.DESCRIPTION
# Exporta os registros para o arquivo "relatorio.xlsx"
# Formato:
# ------------------------------------------------------------
# | Nome no NAC   | Incoerencia        | Descricao detalhada |
# ------------------------------------------------------------
# | LIT-DTI123456 | Descricao do erro  | AD:LIT-DLG123456    |
# ------------------------------------------------------------
#
#.PARAMETER erros
# Lista de erros gerados ao longo da execução
##############################
Function exportarEmExcel {
    Param($erros)
    Try {
        $excel = New-Object -com 'excel.application'
        $pastas = $excel.Workbooks.Add()
        $tabela = $pastas.Worksheets.Item(1)
        $tabela.Name = 'Relatorio'
        $tabela.Cells.Item(1, 1) = 'Nome no NAC'
        $tabela.Cells.Item(1, 2) = 'Incoerencia'
        $tabela.Cells.Item(1, 3) = 'Descricao'
        $i = 2
        foreach($erro in $erros) {
            $tabela.Cells.Item($i, 1) = $erro.Nome
            $tabela.Cells.Item($i, 2) = $erro.Incoerencia
            $tabela.Cells.Item($i, 3) = $erro.Alias
            $i++
        }
        $pastas.SaveAs("$PSScriptRoot\relatorio.xlsx")
        $excel.Quit()
        Write-Output "Relatorio de integridade salvo em '$PSScriptRoot\relatorio.xlsx'."
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

# Ponto de entrada do script
Clear-Host
$arquivo = Read-Host -Prompt 'Digite o nome do relatório do NAC que deseja analisar'
executaScript $arquivo