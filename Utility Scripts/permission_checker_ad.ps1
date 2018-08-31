
Function verificaPermissoes {
    Param($pasta)
    Process {
        Write-Output 'Verificando permissões em pastas compartilhadas.'
        $registros = @{}
        Try {
            $diretorios = Get-ChildItem $pasta -Directory -ErrorAction SilentlyContinue -Recurse
            Foreach ($diretorio In $diretorios.FullName) {
                $acl = Get-ChildItem $diretorio -ErrorAction SilentlyContinue | where-object {($_.PsIsContainer)} | Get-ACL -ErrorAction SilentlyContinue
                If ($acl) {
                    Foreach($idRef In $acl.Access.IdentityReference){
                        If ($idRef.Value -match 'AD\\[0-9]{8}' -and $acl.Owner -notcontains $idRef.Value) {
                            If ($registros.Contains($idRef.Value)){
                                If ($registros[$idRef.Value] -notcontains $diretorio) {
                                    $registros[$idRef.Value] += $diretorio
                                }
                            } Else {
                                $lista = @()
                                $lista += $diretorio
                                $registros.Add($idRef.Value, $lista)
                            }
                        }
                    }
                }
            }
            organizaDados $registros
        } Catch {
            Write-Output 'Falha na verificação de permissões.'
            finalizar
        }
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

##############################
#.SYNOPSIS
# Exporta registros para um arquivo Excel
#
#.DESCRIPTION
# Exporta os registros para o arquivo "relatorio.xlsx"
# Formato:
# ----------------------------------------------------
# | Cartão UFRGS | Nome do usuário | Pasta           |
# ----------------------------------------------------
# | 12345678     | João da Silva   | //local/past    |
# ----------------------------------------------------
#
#.PARAMETER erros
# Lista de erros de permissao gerados ao longo da execução
##############################
Function exportarEmExcel {
    Param($erros)
    Process{
        Try {
            $excel = New-Object -com 'excel.application'
            $pastas = $excel.Workbooks.Add()
            $tabela = $pastas.Worksheets.Item(1)
            $tabela.Name = 'Relatorio'
            $tabela.Cells.Item(1, 1) = 'Cartão UFRGS'
            $tabela.Cells.Item(1, 2) = 'Nome de usuário'
            $tabela.Cells.Item(1, 3) = 'Pasta'
            $i = 2
            foreach($erro in $erros) {
                $tabela.Cells.Item($i, 1) = $erro.Cartao
                $tabela.Cells.Item($i, 2) = $erro.Nome
                $tabela.Cells.Item($i, 3) = $erro.Local
                $i++
            }
            $pastas.SaveAs("$PSScriptRoot\relatorio.xlsx")
            $excel.Quit()
            Write-Output "Relatorio de permissões salvo em '$PSScriptRoot\relatorio.xlsx'."
            finalizar
        } Catch {
            Write-Output 'Falha ao salvar o arquivo de relatorio.'
            finalizar
        }
    }
}

Function organizaDados {
    Param($registros)
    Process {
        Write-Output 'Organizando dados...'
        $saida = @()
        Foreach ($key In $registros.Keys) {
            $item = $registros[$key]
            $ufrgsId = $key.Substring($key.Length - 8).ToString()
            $username = Get-ADUser $ufrgsId
            Foreach ($sub in $item) {
                $obj = [PSCustomObject]@{
                    Cartao = $ufrgsId
                    Nome = $username.Name
                    Local = $sub
                }
                $saida += $obj
            }
        }
        exportarEmExcel $saida
    }
}

Clear-Host
$pasta = Read-Host -Prompt 'Digite a pasta de rede que deseja verificar as permissões [\\controller.domain.com\PASTA]'
verificaPermissoes $pasta