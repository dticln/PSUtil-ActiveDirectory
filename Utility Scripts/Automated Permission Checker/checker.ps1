# Script Checker
<# 
.SYNOPSIS
Script para automatização de tarefas no Active Directory

.DESCRIPTION
O script é um conjunto de ferramentas utilizadas para automatizar tarefas de verificação
e e geração de relatórios no Active Directory.
A versão atual conta com as seguintes funcionalidades:
- verificação de permissões (Setores e Grupos);
- geração de relatório de tamanho de arquivos;
- geração de relatório de duplicatas de arquivos;
- possibilidade de envio automatizado de e-mails para as chefias do setor;

Bibliotecas e dependências:
O Script presente nesse arquivo necessita da extensão Microsoft Active Directory,
bibliotecas de abertura de arquivos do Microsoft Office Excel e do arquivo util.ps1.
Para o envio de e-mails, utilizam-se os layouts presentes na pasta "layout".
O Script também necessita de permissão de administrador do domínio para realizar as
consultas nas bases de dado do AD e nos arquivos.  
A pasta "arquivo" é utilização para salvamento dos relatórios em formato HTML e XLSX.

Sobre o arquivo de referência:
A verificação e geração de relatórios necessita de um arquivo de referência
com as informações utilizadas pelo script para a execução.
O arquivo deve ser aceito pelo Excel (que executará o processamento das células) e estar no formato:

  |   A   |      B      |      C       |    D   |   E    |       F       | 
--------------------------------------------------------------------------
1 | Sigla |     Nome    | Responsável  | Cartão | E-mail |     Pasta     | 
--------------------------------------------------------------------------
2 | CLN   | Campus L... | Fulano de... |  12345 | ful... | \\ad.ufrgs... |

.NOTES
Autor: Divisão de Tecnologia da Informação do Campus Litoral Norte.
Script Version: 1.0.0.0

.LINK
https://github.com/dticln/PSUtil-ActiveDirectory
#>
Import-Module '.\util.ps1'

<#
.SYNOPSIS
Recupera uma lista de departamentos de um arquivo excel.

.DESCRIPTION
Recupera uma lista de departamentos (Sigla, Nome, Responsável, Cartão UFRGS, E-mail, Pasta padrão)
de um arquivo excel organizado em tabelas com esses campos.

.PARAMETER filename
Nome do arquivo onde estão as informações dos departamentos 
Ver: arquivo de referência, sinopse do script.

.NOTES
O script foi configurado para coletar a informação 
a partir da segunda linha da tabela.
O retorno esperado é uma lista de Departamentos:
[PSCustomObject]@{ Sigla, Nome, Responsavel, Cartao, Email, Pasta }
#>
Function Get-DepartmentFromFile {
    Param($filename)
    Begin {
        Write-Host "Tentando abrir arquivo $filename."
    }
    Process {
        $departments = @()
        Try {
            $excel, $folder = Open-Excel $filename 
            $sheet = $folder.ActiveSheet
        }
        Catch {
            Write-Error "Não foi possível abrir a planilha $filename."
            $excel.Quit()
            Exit-Program
        }
        For ($i = 2; $sheet.Cells.Item($i, 1).Value(); $i++) {
            $department = [PSCustomObject]@{
                Sigla       = $sheet.Cells.Item($i, 1).Value()
                Nome        = $sheet.Cells.Item($i, 2).Value()
                Responsavel = $sheet.Cells.Item($i, 3).Value()
                Cartao      = $sheet.Cells.Item($i, 4).Value()
                Email       = $sheet.Cells.Item($i, 5).Value()
                Pasta       = $sheet.Cells.Item($i, 6).Value()
            }
            $departments += $department
        }
        $excel.Quit()
        return $departments
    }
    End {
        Write-Host "Lista de setores aberta. Iniciando processo de verificação."
    }
}

<#
.SYNOPSIS
Recupera uma lista de membros e subgrupos de um grupo

.DESCRIPTION
Recupera uma lista com membros e subgrupos de um grupo do Active Directory 
a partir do seu nome 

.PARAMETER department
Nome do departamento que deseja realizar a recuperação de membros e subgrupos

.PARAMETER useLimiter
É possível utilizar o limitador "*Everyone" durante a pesquisa, realizando um filtro
em possíveis subgrupos com everyone no nome.
Essa opção é válida na estrutura de pastas utilizada pelo CLN

.EXAMPLE
Get-ListFromDepartment "LIT DGR Everyone" $true

.NOTES
Há uma chamada recursiva dessa função. Ela deve ter o "useLimiter" padrão: false.
O resultado esperado da função é um Grupo:
[PSCustomObject]@{ Nome, Descricao, Subgrupos, Usuarios }
#>
Function Get-ListFromDepartment {
    Param(
        $department,
        $useLimiter = $false
    )
    Begin {}
    Process {
        Try {
            $adGroup = Get-ADGroup `
                -Properties Name, GroupCategory, Description `
                -Filter {Name -like $department} 
            $group = [PSCustomObject]@{
                Nome      = $adGroup.Name
                Descricao = $adGroup.Description
                Subgrupos = @()
                Usuarios  = @()
            }
            $members = Get-ADGroupMember $adGroup
            Foreach ($member In $members) {
                If ($member.objectClass -eq 'group') {
                    If (!$useLimiter -or $member.Name -notlike '*Everyone') {
                        $group.Subgrupos += Get-ListFromDepartment $member.Name
                    }
                }
                Else {
                    $group.Usuarios += $member.Name
                }
            }
            return $group
        }
        Catch {
            return $null
        }
    }
    End {}
}

<#
.SYNOPSIS
Cria tabela HTML a partir de Grupo

.DESCRIPTION
Cria tabela HTML a partir de um Grupo. Em um primeiro momento, verifica-se a existência de usuários 
no grupo: caso haja, serão listados em formato de células. A função faz chamados recursivos para cada subgrupo.
A diversas tabelas geradas são concatenadas e retornadas em formato de texto (HTML). 

.PARAMETER groups
Grupo inicial e subgrupos das chamadas recursivas

.EXAMPLE
$htmlText = Convert-GroupToTable $groups

.NOTES
Há uma verificação manual para grupos definidos pelo CPD. É realizado um match no padrão:
'([G])[0-9]{5}\w+'. Exemplo: G12345
A função necessita dos layouts: table-header, table-cell e table.
#>
Function Convert-GroupToTable {
    Param($groups)
    Begin {}
    Process {
        $table = ''
        If ($groups.Usuarios.Length -gt 0) {
            If ($groups.Nome -Match '([G])[0-9]{5}\w+') {
                $description = "O grupo <b>$($groups.Nome)</b> ($($groups.Descricao)) é criado e administrado pelo CPD:"
            }
            Else {
                $description = "O grupo <b>$($groups.Nome)</b> ($($groups.Descricao)) é administrado pela DTICLN:"
            }
            $cells = Get-Layout 'table-header' @{
                header = $groups.Nome
            }
            Foreach ($user in $groups.Usuarios) {
                $cells += Get-Layout 'table-cell' @{
                    content = $user
                }
            }
            $table = Get-Layout 'table' @{
                description = $description
                cells       = $cells
            }
        }
        Foreach ($grupo In $groups.Subgrupos) {
            $table += Convert-GroupToTable $grupo
        }
        return $table
    }
    End {}
}

<#
.SYNOPSIS
Cria tabela HTML a partir de de uma lista de Permissões

.DESCRIPTION
Cria uma tabela HTML a uma lista de permissões manuais atribuídas a um determinado setor.
Para cada usuário, verifica-se o número de permissões manuais na pasta do setor. Caso haja 
mais de 10 registros manuais, a tabela constará com o aviso de múltiplos registros.
Caso contrário, serão listados os registros de cada usuário.

.PARAMETER permissions
Lista de permissões manuais geradas pela função Set-ObjectToDepartment .

.EXAMPLE
$htmlText = Convert-PermissionToTable $permissions

.NOTES
Essa função necessita do layouts: table-header, table-cell-permission e table
#>
Function Convert-PermissionToTable {
    Param($permissions) 
    Begin {}
    Process {
        $cells = Get-Layout 'table-header' @{
            header = 'Permissões configuradas manualmente'
        }
        $users = $permissions.Nome | Select-Object -Unique
        Foreach ($user In $users) {
            $registries = $permissions | Where-Object {$_.Nome -match $user}
            $count = @($registries).Count
            If ($count -gt 10) {
                $permissions = $permissions | Where-Object { -not ($_.Nome -match $user)}
                $cells += Get-Layout 'table-cell-permission' @{
                    name   = $user -Replace "- [0-9]{8}", ""
                    id     = $registries.Cartao | Get-Unique
                    folder = "Foram encontradas $count pastas com permissões manuais na pasta do setor. Entrar em contato com a DTI para mais informações."
                }
            }
        }
        Foreach ($permission In $permissions) {
            $cells += Get-Layout 'table-cell-permission' @{
                name   = $permission.Nome -Replace "- [0-9]{8}", ""
                id     = $permission.Cartao
                folder = $permission.Local 
            }
        }
        $table = Get-Layout 'table' @{
            description = "Além dos grupos descritos acima, existem as seguintes permissões configuradas de forma manual:"
            cells       = $cells
        }
        return $table
    }
    End {}
}

<#
.SYNOPSIS
Recupera informações de permissões manuais inseridas nas pastas.

.DESCRIPTION
Recupera uma lista de permissões inseridas de forma manual.
A função tem como lógica, filtrar todas as permissões de uma pasta e subpastas.
O filtro segue a lógica: caso a permissão dê match com 'AD\\[0-9]{8}' e a permissão 
não seja do criador da pasta, adiciona-se uma lista de permissões desse usuário em 
uma lista geral. Para cada registro de pasta com permissão manual, um novo item é 
adicionado na lista do usuário.
Os resultados são reagrupados utilizando a função Group-Permission: as listas 
de cada usuário são desfeitas e transformadas em uma grande lista. 

.PARAMETER folder
Pasta raiz da verificação de permissão. Por padrão, utiliza-se \\ad.ufrgs.br\LITORAL

.EXAMPLE
$permissions = Find-ManualPermission '\\ad.ufrgs.br\LITORAL'

.NOTES
A função pode disparar uma exceção que finaliza a execução do programa.
A causa mais provável é a falta de credenciais para o acesso.
#>
Function Find-ManualPermission {
    Param($folder)
    Begin {
        Write-Host 'Verificando permissões em pastas compartilhadas.'
    }
    Process {
        $registries = @{}
        Try {
            $directories = Get-ChildItem $folder -Directory -ErrorAction SilentlyContinue -Recurse
            Foreach ($directory In $directories.FullName) {
                $acls = Get-ChildItem $directory -ErrorAction SilentlyContinue | where-object {($_.PsIsContainer)} | Get-ACL -ErrorAction SilentlyContinue
                Foreach ($acl In $acls) {
                    Foreach ($idRef In $acl.Access.IdentityReference) {
                        If ($idRef.Value -match 'AD\\[0-9]{8}' -and $acl.Owner -notcontains $idRef.Value) {
                            If ($registries.Contains($idRef.Value)) {
                                If ($registries[$idRef.Value] -notcontains $directory) {
                                    $registries[$idRef.Value] += "$directory\$($acl.PSChildName)"
                                }
                            }
                            Else {
                                $list = @()
                                $list += "$directory\$($acl.PSChildName)"
                                $registries.Add($idRef.Value, $list)
                            }
                        }
                    }
                }
            }
            Return Group-Permission $registries
        }
        Catch {
            Write-Host 'Falha na verificação de permissões.'
            Exit-Program
        }
    }
    End {}
}

<#
.SYNOPSIS
Recupera lista com arquivos duplicados em umas estrutura de pasta

.DESCRIPTION
Recupera uma lista de arquivos duplicados. Para cada arquivos,
gera-se uma HASH que identifique seu conteúdo. O resultado inicial é uma lista
de listas com as ocorrências de cada duplicata. 

Esse processo pode ser demorado, levando em consideração o número de itens que 
serão analisados e a combinação da analise desses itens.

Os dados são agrupados posteriormente pela função Group-DuplicatedFiles,
onde as listas de ocorrências são transformadas em uma única lista.

.PARAMETER path
Carinho de rede para a pasta que será analisada recursivamente

.EXAMPLE
$duplicated = Find-DuplicatedFiles '\\ad.ufrgs.br\LITORAL'

.NOTES
Espera-se melhorar essa função, procurando uma forma mais performática
de executar a verificação de unicidade de arquivo. Nome e tamanho não
devem ser usados como referência nessa análise, uma vez que vários arquivos
são encontramos com nomes diferentes.
#>
Function Find-DuplicatedFiles {
    Param ($path)
    Begin {
        Write-Host "Verificando arquivos duplicados em $path."
    }
    Process {
        $list = Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue
        $found = @{}
        Foreach ($item In $list) {
            $hash = Get-FileHash -Path $item.PSPath -ErrorAction SilentlyContinue
            If ($hash.Hash) {
                If (!$found.ContainsKey($hash.Hash)) {
                    $found[$hash.Hash] = @()
                    $found[$hash.Hash] += $item
                }
                Else {
                    $found[$hash.Hash] += $item
                    $double = $hash.Hash
                    $duplicated = $found[$hash.Hash]
                    Write-Host "O arquivo '$double' está presente em $($duplicated.Length) diretórios."
                }
            } 
        }
        Return Group-DuplicatedFiles $found
    }
    End {}
}

<#
.SYNOPSIS
Recupera lista de arquivos agrupados por tamanho em determinado pasta

.DESCRIPTION
Recupera lista de todos os arquivos em uma pasta e em suas subpastas agrupado por
tamanho. 

.PARAMETER path
Local que será alvo da pesquisa

.EXAMPLE
$files = Find-LargeFiles '\\ad.ufrgs.br\LITORAL'

.NOTES
A função retorna uma lista do objeto [PSCustomObject]@{ Tamanho, Nome, Local },
facilitando a iteração e a utilização em outras funções, como Set-ObjectToDepartment.
#>
Function Find-LargeFiles {
    Param ($path)
    Begin {
        Write-Host "Verificando tamanho dos arquivos em $path."
    }
    Process {
        $directories = Get-ChildItem -Path $path -Recurse -ErrorAction SilentlyContinue
        $onlyFiles = $directories | Where-Object -Property Attributes -eq 'Archive'
        $list = $onlyFiles | Sort-Object Length -Descending | Select-Object Length, Name, Directory
        $export = @()
        Foreach ($item in $list) {
            $export += [PSCustomObject]@{
                Tamanho = Format-Bytes $item.Length
                Nome    = $item.Name
                Local   = $item.Directory
            }
        }
        return $export
    }
    End {}
}

<#
.SYNOPSIS
Gera arquivo Excel com a lista de arquivos duplicados

.DESCRIPTION
Gera arquivo Excel com a lista de arquivos duplicados. Cada arquivo igual
é separado por um tom de cinza diferente, intercalado. Ou seja, se o arquivo "x"
se repete quatro vezes e o arquivo "y", duas: teremos quatro linhas com 0xdcdcdc e
duas linhas 0xf5f5f5. O próximo arquivo diferentes desses, receberá novamente 0xdcdcdc.

.PARAMETER list
Lista com objetos que representam as duplicatas. Esses objetos são gerados pela função 
Group-DuplicatedFiles.

.PARAMETER name
Nome que será dado ao arquivo. Por padrão, envia-se o nome do departamento

.EXAMPLE
Export-DuplicatedFiles $list 'LIT DGR ADG'

.NOTES
Para exportar os dados a função utiliza a biblioteca do Excel
#>
Function Export-DuplicatedFiles {
    Param (
        $list,
        $name
    )
    Begin {}
    Process {
        Try {
            $gray = 0xdcdcdc
            $lowGray = 0xf5f5f5
            $curColor = $gray
            $excel = New-Object -com 'excel.application'
            $folder = $excel.Workbooks.Add()
            $table = $folder.Worksheets.Item(1)
            $table.Name = 'Duplicatas'
            $table.Cells.Item(1, 1) = 'Identificador único'
            $table.Cells.Item(1, 2) = 'Nome do arquivo'
            $table.Cells.Item(1, 3) = 'Caminho'
            $i = 2
            $lastHash = $list[0]
            Foreach ($item In $list) {
                $table.Cells.Item($i, 1) = $item.HashID
                $table.Cells.Item($i, 2) = $item.Nome
                $table.Cells.Item($i, 3) = $item.Local
                If ($lastHash -ne $item.HashID) {
                    $curColor = @{$true = $gray; $false = $lowGray}[$curColor -eq $lowGray]
                }
                $lastHash = $item.HashID
                $table.Cells.Item($i, 1).Interior.Color = $curColor
                $table.Cells.Item($i, 2).Interior.Color = $curColor
                $table.Cells.Item($i, 3).Interior.Color = $curColor
                $i++
            }
            $folder.SaveAs("$PSScriptRoot\arquivo\$name.xlsx")
            $excel.Quit()
            Return "$PSScriptRoot\arquivo\$name.xlsx"
        }
        Catch {
            $excel.Quit()
            Write-Host "Não foi possível salvar o relatório '$name.xlsx'."
        }  
    }
    End {
        Write-Host "Relatório de duplicatas salvo em '$name.xlsx'."
    }
}

<#
.SYNOPSIS
Gera arquivo Excel com a lista de arquivos ordenada por tamanho

.DESCRIPTION
Gera arquivo Excel com a lista de arquivos ordenada por tamanho. É possível estabelecer
um limite de itens que serão indexados nesse arquivo.

.PARAMETER list
Lista de arquivos já ordenada por tamanho.

.PARAMETER name
Nome que será dado ao arquivo. Geralmente o nome do setor

.PARAMETER limit
Limite de gravação no arquivo Excel. O número de itens nessa lista pode ser extenso,
o que obriga a gerar um arquivo mais palpável para os usuários finais.

.EXAMPLE
Export-LargeFiles $list 'LIT DGR ADG' 150

.NOTES
Para exportar os dados a função utiliza a biblioteca do Excel
#>
Function Export-LargeFiles {
    Param (
        $list,
        $name,
        $limit = 100
    )
    Begin {}
    Process {
        Try {
            $excel = New-Object -com 'excel.application'
            $folder = $excel.Workbooks.Add()
            $table = $folder.Worksheets.Item(1)
            $table.Name = 'Relatório'
            $table.Cells.Item(1, 1) = 'Tamanho'
            $table.Cells.Item(1, 2) = 'Nome'
            $table.Cells.Item(1, 3) = 'Local'
            $i = 2
            foreach ($item in ($list | Select-Object -First $limit)) {
                $table.Cells.Item($i, 1) = $item.Tamanho 
                $table.Cells.Item($i, 2) = $item.Nome
                $table.Cells.Item($i, 3) = $item.Local
                $i++
            }
            $folder.SaveAs("$PSScriptRoot\arquivo\$name.xlsx")
            $excel.Quit()
            Return "$PSScriptRoot\arquivo\$name.xlsx"
        }
        Catch {
            $excel.Quit()
            Write-Host "Não foi possível salvar o relatório '$name.xlsx'."
        }  
    }
    End {
        Write-Host "Relatório de duplicatas salvo em '$name.xlsx'."
    }
}

<#
.SYNOPSIS
Reagrupa as listas de permissões em uma lista única.

.DESCRIPTION
Reagrupa as listas de permissões de cada usuário em uma lista única.
Gera uma lista objetos Permissão.

.PARAMETER registries
Registros gerados pela primeira parte da função Find-ManualPermission.

.EXAMPLE
return Group-Permission $registries

.NOTES
O resultado da função é uma lista de objetos: [PSCustomObject]@{ Cartao, Nome, Local }
#>
Function Group-Permission {
    Param($registries)
    Begin {
        Write-Host "Organizando dados..."
    }
    Process {
        $output = @()
        Foreach ($key In $registries.Keys) {
            $item = $registries[$key]
            $id = $key.Substring($key.Length - 8).ToString()
            $username = Get-ADUser $id
            Foreach ($sub in $item) {
                $obj = [PSCustomObject]@{
                    Cartao = $id
                    Nome   = $username.Name
                    Local  = $sub
                }
                $output += $obj
            }
        }
        return $output
    }
    End {}
}

<#
.SYNOPSIS
Reagrupa as listas de arquivos duplicados em uma lista única.

.DESCRIPTION
Reagrupa os dados presentes em sub-listas em uma lista única de arquivos duplicados

.PARAMETER duplicated
Registros gerados pela primeira parte da função Find-DuplicatedFiles

.EXAMPLE
return Group-DuplicatedFiles $duplicated

.NOTES
O resultado da função é uma lista de objetos [PSCustomObject]@{ HashID, Nome, Local }
#>
Function Group-DuplicatedFiles {
    Param ($duplicated)
    Begin {}
    Process {
        $export = @()
        Foreach ($key In $duplicated.Keys) {
            If ($duplicated[$key].Length -gt 1) {
                Foreach ($file In $duplicated[$key]) { 
                    $export += [PSCustomObject]@{
                        HashID = $key
                        Nome   = $file.Name
                        Local  = $file.FullName
                    }
                }
            }
        }
        Return $export
    }
    End {}
}

<#
.SYNOPSIS
Atribui objetos a uma Unidade Organizacional

.DESCRIPTION
Organiza objetos a uma Unidade Organizacional a partir da comparação do atributo
"Local". Essa identificação ocorre por meio da comparação do "Local" como substring de uma lista de referencias.
Para cada registro não atribuído, verifica-se se ele é substring de uma das referências, quando identificação a correlação,
cria-se um "match", quando maior o tamanho da string, maior a correlação entre o item
e a referência. Dessa forma, impede-se que uma permissão ou arquivo de uma pasta de uma secretaria
seja atribuída à Direção e não a chefia da secretaria.

.PARAMETER registries
Registros (lista) que serão organizados: podem ser arquivos, permissões em pastas específicas, etc...

.PARAMETER base
Local de referência daquele setor. Pasta raiz na estrutura hierárquica do AD.

.PARAMETER references
Lista com todas as referências presentes no arquivo de setores.

.EXAMPLE
$departmentPermission = Set-ObjectToDepartment $permissions $department.Pasta $departments.Pasta

.NOTES
Como não pareceu uma boa ideia "retirar" itens da lista, pareceu mais interessante 
marcar os itens já "atribuídos" a um setor com o atributo "Attributed" com "Add-Member Attributed $true".
Como desvantagem, os itens vão continuar sendo iterados na próxima execução do script. Como ponto positivos,
pode-ser utilizar essa lista pra mais alguma outra tarefa.
#>
Function Set-ObjectToDepartment {
    Param (
        $registries, 
        $base, 
        $references
    )
    Begin {}
    Process {
        $attributed = @()
        Foreach ($registry In $registries) {
            If ('Attributed' -notin $registry.PSObject.Properties.Name) {
                $bestMatch = $null
                Foreach ($reference In $references) {
                    If ($registry.Local -like "$reference*") {
                        If ($null -eq $bestMatch) {
                            $bestMatch = $reference
                        } 
                        If ($reference.Length -gt $bestMatch.Length) {
                            $bestMatch = $reference
                        }
                    }
                }
                If ($bestMatch -eq $base) {
                    $registry | Add-Member Attributed $true
                    $attributed += $registry
                }
            }
        }
        Return $attributed
    }
    End {}
}

<#
.SYNOPSIS
Função com a lógica principal do programa.

.DESCRIPTION

O seguinte script foi pensado inicialmente como um verificador de permissões nas pastas departamentais,
entretanto a necessidade de mais alguns funções fez com que ele fosse modificado para receber mais funcionalidades. 

São funcionalidades presentes nesse script:

1. Verificação de permissões
Executa a verificação de permissões das pastas departamentais da unidade.
Para que o script funcione de forma apropriada, é necessário seguir o padrão de organização
proposto pelo manual do CPD e do Active Directory.
O script pode ser adaptado para operar em outros ambientes.

Para a execução, é necessário gerar uma lista com os setores (para mais informações, ver: Get-DepartmentFromFile).

Verifica-se as informações de permissões manuais na pasta rais da unidade.
É realizada a atribuição de permissões por setor e a listagem de todos os usuários
que possuem acesso à pasta do setor, distribuído em seus respectivos grupos.

O objetivo do script é gerar e-mails para todos os setores.
Esses e-mails serão armazenados na pasta "arquivo" e podem ser enviados automaticamente,
utilizando a opção $sendEmail.

2. Verificação de arquivos duplicados
Procura em uma pasta específica todos os arquivos duplicados em suas subpastas.
A verificação de duplicação é feita utilizando uma Hash gerada a partir do conteúdo do arquivo,
diminuindo muito a chance de falsas duplicatas e identificando duplicatas que tenham metadados
distintos.

Assim como na verificação de permissão, os dados podem ser enviados por e-mail utilizando $sendEmail

3. Verificação de tamanho de arquivos
Gera uma lista ordenada por tamanho de arquivo para cada setor disposto no arquivo de referência.
Essa lista pode ser limitada para um número x de itens por setor.

Os dados podem ser acessados na pasta arquivo ou enviados por e-mail utilizando $sendEmail.

.PARAMETER sendEmail
Permite o envio automatizado de e-mails para os setores.

.EXAMPLE
Start-Program $true
#>
Function Start-Program {
    Param(
        $option
    )
    Begin {
        Write-Host "Para executar a verificação, é necessário utilizar um arquivo de referência com as informações dos setores."
        $in = Read-HostWithDefault 'Qual o nome do arquivo de referência?' 'setores.xlsx' 
        $departments = Get-DepartmentFromFile $in
        $sendEmail = Show-YesNoQuestion 'Deseja enviar os resultados para o e-mail dos responsáveis pelos setores (S/N)? [N]' `
            'Enviaremos os dados por e-mail.' `
            "Não enviaremos os dados por e-mail. Você pode acessar os dados na pasta 'arquivo'." `
            "Resposta não reconhecida. O envio de e-mail foi desabilitado. "
        If ($sendEmail) {
            $credencial = Get-Credential -Message 'Insira as credenciais utilizadas para o envio dos e-mails (info-ticln@ufrgs.br):'
        }
    }
    Process {
        Switch ($option) {
            ## Verificação de permissões 
            '1' {
                $manualPermission = Show-YesNoQuestion 'Deseja realizar a verificação de permissões manuais (S/N)? [N]' `
                    'Verificaremos as permissões manuais. ' `
                    'Não verificaremos as permissões manuais.' `
                    'Resposta não reconhecida. A verificação de permissões manuais foi desabilitada.'
                If ($manualPermission) {
                    $in = Read-HostWithDefault 'Qual a pasta que será utilizada como referência?' '\\ad.ufrgs.br\LITORAL' 
                    $manual = Find-ManualPermission $in
                }
                Foreach ($department In $departments) {
                    $initials = $department.Sigla.trim()
                    Write-Host "Verificando estrutura do setor $initials."
                    $groups = Get-ListFromDepartment "$initials Everyone" $true
                    $groupTable = Convert-GroupToTable $groups
                    $permissionTable = ''
                    If ($manualPermission) {
                        $departmentPermission = Set-ObjectToDepartment $manual $department.Pasta $departments.Pasta
                        If ($departmentPermission) {
                            $permissionTable = Convert-PermissionToTable $departmentPermission
                        }
                    }
                    $email = Get-Layout 'e-mail-permission' @{ 
                        groups      = $groupTable
                        permissions = $permissionTable
                        department = $department.Nome
                        folder = $departments.Pasta
                    }
                    $file = "$PSScriptRoot\arquivo\$initials.html"
                    $email | Out-File $file -Encoding UTF8 
                    ##.TODO
                    # DEFINIR CORPO DO E-MAIL
                    If ($sendEmail) {
                        Write-Host "Enviando e-mail de confirmação para: $($department.Email)."
                        Send-Email $department.Email 'Verificação de permissões' $file $credencial $false
                    }
                }
            }
            ## Verificação de arquivos duplicados
            '2' {
                Write-Output "*ATENÇÃO: A verificação de duplicatas pode demorar um grande período de tempo."
                $in = Read-HostWithDefault 'Qual a pasta que será utilizada como referência?' '\\ad.ufrgs.br\LITORAL' 
                $duplicated = Find-DuplicatedFiles  $in
                Foreach ($department In $departments) {
                    $initials = $department.Sigla.trim()
                    Write-Output "Analisando duplicatas para $initials."
                    $departmentDuplicated = Set-ObjectToDepartment $duplicated $department.Pasta $departments.Pasta
                    If ($departmentDuplicated.Count -gt 0) {
                        $body = Get-Layout 'e-mail-duplicated' @{ 
                            department = $department.Nome
                            folder = $departments.Pasta
                        }
                        $file = Export-DuplicatedFiles $departmentDuplicated $initials
                        If ($sendEmail) {
                            Write-Host "Enviando e-mail de duplicatas para: $($department.Email)."
                            Send-Email $department.Email 'Relatório de duplicatas' $body $credencial $file
                        }
                    }
                }
            }
            ## Verificação de maiores arquivos por setor
            '3' {
                $in = Read-HostWithDefault 'Qual a pasta que será utilizada como referência?' '\\ad.ufrgs.br\LITORAL' 
                $large = Find-LargeFiles $in
                Foreach ($department In $departments) {
                    $initials = $department.Sigla.trim()
                    Write-Output "Analisando os maiores arquivos para $initials."
                    $departmentLarge = Set-ObjectToDepartment $large $department.Pasta $departments.Pasta
                    If ($departmentLarge.Count -gt 0) {
                        $file = Export-LargeFiles $departmentLarge $initials
                        ##.TODO
                        # DEFINIR CORPO DO E-MAIL
                        If ($sendEmail) {
                            Write-Host "Enviando e-mail de maiores arquivos para: $($department.Email)."
                            Send-Email $department.Email 'Relatório de maiores arquivos' 'body' $credencial $file
                        }
                    }
                }
            }
        }
    }
    End {
        Exit-Program
    }
}

<#
.SYNOPSIS
Exibe cabeçalho com nome do script

.DESCRIPTION
Exibe cabeçalho com o nome do script, versão e autor
#>
Function Show-Label {
    Clear-Host
    Write-Host "=========================================================="
    Write-Host "| Script de verificação automatizada do Active Directory |"
    Write-Host "=========================================================="
    Write-Host "|          Divisão de Tecnologia da Informação           |"
    Write-Host "|                   Versão: 1.0.0.0                      |"
    Write-Host "=========================================================="
}

<#
.SYNOPSIS
Menu de decisão principal

.DESCRIPTION
Menu de decisão principal do programa, apresenta as opções e encaminha para a execução
do script em Start-Program
#>
Function Show-Menu {
    $continue = $false
    Do {
        Show-Label
        Write-Host "Informe qual operação deseja realizar: "
        Write-Host "1 - Verificação de permissões"
        Write-Host "2 - Verificação de arquivos duplicados"
        Write-Host "3 - Relatórios de tamanhos de arquivos"
        Write-Host "0 - Fechar programa"
        If ($continue) {
            Write-Host '* Insira uma resposta válida.'
        }
        $in = Read-Host -Prompt 'Digite a opção desejada?'
        If (($in -eq '1') -or ($in -eq '2') -or ($in -eq '3')) {
            Start-Program $in
            $continue = $false
        }
        ElseIf ($in -eq '0') {
            Exit-Program 
        }
        Else {
            $continue = $true 
        }
    } While ($continue)
}

## Ponto de partida do script
Show-Menu