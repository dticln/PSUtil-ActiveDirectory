# Funções auxiliares do Script Checker 

<#
.SYNOPSIS
Abre arquivo do Excel

.DESCRIPTION
Abre arquivo do Excel, retornando o objeto Excel e a Pasta escolhida

.PARAMETER filename
Nome do arquivo que será aberto

.EXAMPLE
$excel, $folder = Open-Excel 'setores.xlsx'

.NOTES
O retorno de dois objetos pode ser algo perigoso, 
mas pareceu util utiliza-lo nessa função.
#>
Function Open-Excel {
	Param($filename)

    try {
        $excel = New-Object -com 'excel.application'

        $folder = $excel.Workbooks.Open("$root$filename")
        return $excel, $folder
    }
    catch {
        Write-Error "Erro ao abrir o arquivo Excel. Não foi possível localizar o arquivo $filename."
        break
    }
}

<#
.SYNOPSIS
Finaliza o programa

.DESCRIPTION
Finaliza a execução do programa mostrando uma mensagem de 
"pressione qualquer tecla para finalizar".

.EXAMPLE
Exit-Program

.NOTES
Para essa função, utiliza-se comandos cmd. Não muito aconselhado,
mas funciona e não é algo crítico: quando o programa alcança esse ponto de execução,
ele já realizou as tarefas necessárias
#>
Function Exit-Program {
    Write-Output "Pressione qualquer tecla para finalizar."
    cmd /c pause | out-null
    Exit 1
}

<#
.SYNOPSIS
Envia e-mail utilizando um cliente SMTP

.DESCRIPTION
Envia e-mail utilizando um cliente SMTP, o e-mail é configurado para estar no formato HTML.

.PARAMETER to
Destinatário

.PARAMETER title
Assunto do e-mail

.PARAMETER body
Corpo do e-mail em formato HTML

.PARAMETER user
Credenciais de utilização do e-mail, sua obtenção é feita anteriormente e enviada por parâmetro

.PARAMETER attachments
Possíveis anexos, não é um item necessário.

.EXAMPLE
Send-Email 'departamento@ufrgs.br' '<b>Olá!<b>' $credencial $file
#>
Function Send-Email {
    Param (
        $to, 
        $title,
        $html, 
        $user, 
        $attachments = $false
    )
    Try {
        $From = "[Não responda] Informações DTICLN <info-ticln@ufrgs.br>"
        $To = $to
        $Subject = "[Não responda] $title"
        $Body = Get-Content -Path $html | Out-String
        $SMTPServer = "smtp.ufrgs.br"
        $SMTPPort = "587"
        If ($attachments) {
            Send-MailMessage -From $From -to $To -Attachments $attachments -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $user –DeliveryNotificationOption OnSuccess -Encoding UTF8 
        }
        Else {
            Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $user –DeliveryNotificationOption OnSuccess -Encoding UTF8 
        }
    }
    Catch {
        Write-Error "Não foi possível enviar o e-mail para $to : $_"
        Exit-Program
    }
}

<#
.SYNOPSIS
Recupera um layout em HTML 

.DESCRIPTION
Recupera um layout em HTML em formato de string. O layout deve estar presente dentro da 
pasta "layout".
A função possibilita a substituição de preenchimento de lacunas: essas lacunas são
identificadas por uma palavra entre ":", ou seja :lacuna:. O script varre o HTML procurando 
essas lacunas e substituindo por sua respectiva cadeia de caracteres. (ver .PARAMETER yields)

.PARAMETER name
Nome do layout desejado 

.PARAMETER yields
Lista em formato dicionário ('chave' => 'valor'). A chave representa a cadeia de 
caracteres que será substituída, já o valor, o texto que será colocado no lugar.

Exemplo de processamento:

- $yields:
@{
    header = 'LEGAL :)',
    body = 'Você é um cara legal',
    name = 'Pedro da Silva'
} 

- HTML:
<body>
<h1> :header: </h1>
<p> :body: </p>
<p> Assinado: :name: </p>
</body>

- Saída:
<body>
<h1> LEGAL :) </h1>
<p> Você é um cara legal </p>
<p> Assinado: Pedro da Silva </p>
</body>

.EXAMPLE
Get-Layout 'table' @{
    description = $description
    cells       = $cells
}

.NOTES
Esse script é muito utilizado na composição de tabelas para a geração
de e-mails.
#>
Function Get-Layout {
    Param($name, $yields = @{})
    Try {
		$root = $PSScriptRoot.Substring(0, ($PSScriptRoot.Length - 9))
        $content = Get-Content -Encoding UTF8 -Path "$root\layout\$name.html" <#ajustar parâmetros #>
        Foreach ($key in $yields.Keys) {
            $content = $content -Replace ":$($key):", $yields[$key]
        }
        return $content
    }
    Catch {
        Write-Error "Não foi possível encontrar o layout: $name"
        Exit-Program
    }
    return ""
}

<#
.SYNOPSIS
Gera Hash a partir de texto

.DESCRIPTION
Gera Hash SHA256Managed a partir de texto

.PARAMETER textToHash
Cadeia de caracteres que será encriptada

.EXAMPLE
$hash = Get-Hash 'Marcelo Martins'
#>
function Get-Hash($textToHash) {
    $hasher = new-object System.Security.Cryptography.SHA256Managed
    $toHash = [System.Text.Encoding]::UTF8.GetBytes($textToHash)
    $hashByteArray = $hasher.ComputeHash($toHash)
    foreach ($byte in $hashByteArray) {
        $res += $byte.ToString()
    }
    return $res;
}

<#
.SYNOPSIS
Realiza a formatação de bytes para ser mostrado de forma amigável

.DESCRIPTION
Mostra o tamanho de arquivos utilizando B, KB, MB, GB, TB, PB, EB, ZB e YB,
tornando mais fácil a compreensão

.PARAMETER num
Número em bytes que será convertido para outras unidades de medida

.EXAMPLE
$formatedSize = Format-Bytes 12314
#>
Function Format-Bytes($num) {
    $suffix = "B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB"
    $index = 0
    while ($num -gt 1kb) {
        $num = $num / 1kb
        $index++
    } 
    return "{0:N1} {1}" -f $num, $suffix[$index]
}

<#
.SYNOPSIS
Realiza a leitura de dados na tela utilizando um cabeçalho e um valor padrão

.DESCRIPTION
Realiza a leitura de dados da tela utilizando Read-Host, permitindo a colocação de
um cabeçalho e, em caso de falha, um valor padrão que será utilizado.

.PARAMETER label
Cabeçalho que será mostrado para o usuário.

.PARAMETER default
Valor padrão que será utilizado, caso a captura de dados falhe.
Ele será mostrado ao lado do cabeçalho. Exemplo:
Digite o nome que deseja utilizar: [Pedro da Silva]

.EXAMPLE
$name = Read-HostWithDefault 'Digite o nome que deseja utilizar' 'Pedro da Silva'
#>
Function Read-HostWithDefault {
    Param (
        $label,
        $default
    )
    Process {
        $in = Read-Host -Prompt "$label [$default]"
        If ($in -like '') { 
            $in = $default
        }
        return $in
    }
}

<#
.SYNOPSIS
Diálogo para respostas com "Sim" e "Não"

.DESCRIPTION
Realiza a leitura de dados na tela de forma similar ao Read-HostWithDefault, entretanto
só permite a inserção de "S" e "N". Definindo uma resposta padrão para cada uma das alternativas.
Caso o usuário insira algo diferente disso, retorna falso e apresenta um texto personalizado pro usuário.

.PARAMETER label
Descrição da pergunta

.PARAMETER yesLabel
Texto para resposta positiva

.PARAMETER noLabel
Texto para resposta negativa

.PARAMETER defaultLabel
Texto para entrada de dados incorreta

.EXAMPLE
$sendEmail = Show-YesNoQuestion 'Enviar e-mail?' 'Enviaremos.' 'Não enviaremos' 'Resposta errada. Não enviaremos'
#>
Function Show-YesNoQuestion {
    Param (
        $label,
        $yesLabel,
        $noLabel,
        $defaultLabel = ''
    )
    Process {
        $in = Read-Host -Prompt $label
        Switch ($in) {
            'S' {
                Write-Host $yesLabel
                return $true
            }
            'N' {
                Write-Host $noLabel
                return $false
            }
            Default { 
                Write-Host $defaultLabel
                return $false
            }
        }
    }
}