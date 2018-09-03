Import-Module '.\util.ps1'

<#
.SYNOPSIS
Recupera uma lista de departamentos de um arquivo excel.

.DESCRIPTION
Recupera uma lista de departamentos (Sigla, Nome, Responsável, Cartão UFRGS, E-mail, Pasta padrão)
de um arquivo excel organizado em tabelas com esses campos.

.PARAMETER filename
Nome do arquivo onde estão as informações dos departamentos

.NOTES
O script foi configurado para coletar a informção 
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

.PARAMETER useLimitator
É possível utilizar o limitador "*Everyone" durante a pesquisa, realizando um filtro
em possíveis subgrupos com everyone no nome.
Essa opção é válida na estrutura de pastas utilizada pelo CLN

.EXAMPLE
Get-ListFromDepartment "LIT DGR Everyone" $true

.NOTES
Há uma chamada recursiva dessa função. Ela deve ter o "useLimitator" padrão: false.
O resultado esperado da função é um Grupo:
[PSCustomObject]@{ Nome, Descricao, Subgrupos, Usuarios }
#>
Function Get-ListFromDepartment {
    Param(
        $department,
        $useLimitator = $false
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
                    If (!$useLimitator -or $member.Name -notlike '*Everyone') {
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
Long description

.PARAMETER groups
Parameter description

.EXAMPLE
An example

.NOTES
General notes
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
                $description = "$($groups.Descricao):"
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
Short description

.DESCRIPTION
Long description

.PARAMETER permissions
Parameter description

.EXAMPLE
An example

.NOTES
General notes
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
Short description

.DESCRIPTION
Long description

.PARAMETER folder
Parameter description

.EXAMPLE
An example

.NOTES
General notes
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
                foreach ($acl In $acls) {
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
            return Group-Permission $registries
        }
        Catch {
            Write-Output 'Falha na verificação de permissões.'
            Exit-Program
        }
    }
    End {}
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER registries
Parameter description

.EXAMPLE
An example

.NOTES
General notes
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
Short description

.DESCRIPTION
Long description

.PARAMETER permissions
Parameter description

.PARAMETER base
Parameter description

.PARAMETER references
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Set-ManualPermission {
    Param($permissions, $base, $references)
    Begin {}
    Process {
        $attributed = @()
        Foreach ($registry In $permissions) {
            If ('Attributed ' -notin $registry.PSObject.Properties.Name) {
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
        return $attributed
    }
    End {}
}

<#
.SYNOPSIS
Short description

.DESCRIPTION
Long description

.PARAMETER sendEmail
Parameter description

.EXAMPLE
An example

.NOTES
General notes
#>
Function Start-Program {
    Param($sendEmail = $false)
    Begin {}
    Process {
        If ($sendEmail) {
            Write-Host "Você escolheu iniciar a verificação de permissão por e-mail. "
            $credencial = Get-Credential -Message 'Insira as credenciais utilizadas (dti-cln@ufrgs.br) para o envio dos e-mails:'
        }
        $departments = Get-DepartmentFromFile 'setores.xlsx'
        $manual = Find-ManualPermission '\\ad.ufrgs.br\LITORAL'
        Foreach ($department In $departments) {
            $initials = $department.Sigla.trim()
            Write-Host "Verificando estrutura do setor $initials."
            $groups = Get-ListFromDepartment "$initials Everyone" $true
            $groupTable = Convert-GroupToTable $groups
            $departmentPermission = Set-ManualPermission $manual $department.Pasta $departments.Pasta
            If ($departmentPermission) {
                $permissionTable = Convert-PermissionToTable $departmentPermission
            }
            Else {
                $permissionTable = ''
            }
            $email = Get-Layout 'e-mail' @{ 
                groups      = $groupTable
                permissions = $permissionTable
            }
            $file = "$PSScriptRoot\arquivo\$initials.html"
            $email | Out-File $file -Encoding UTF8 
            If ($sendEmail) {
                Write-Host "Enviando e-mail de confirmação para: $($department.Email)."
                Send-Email $department.Email $file $credencial
            }
        }
    }
    End {}
}

Start-Program $false