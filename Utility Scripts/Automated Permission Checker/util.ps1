
Function Open-Excel {
    Param($filename)
    try {
        $excel = New-Object -com 'excel.application'
        $folder = $excel.Workbooks.Open("$PSScriptRoot\$filename")
        return $excel, $folder
    }
    catch {
        Write-Error "Erro ao abrir o arquivo Excel. Não foi possível localizar o arquivo $filename."
        break
    }
}

Function Exit-Program {
    Write-Output 'Pressione qualquer tecla para finalizar.'
    cmd /c pause | out-null
    break
}

Function Send-Email {
    Param ($to, $body, $user)
    Try {
        $From = "dti-cln@ufrgs.br"
        $To = $to
        ##$Cc = "dti-cln@ufrgs.br"
        $Subject = "Verificação de Permissões"
        $Body = Get-Content -Path $body -Encoding UTF8 | Out-String
        $SMTPServer = "smtp.ufrgs.br"
        $SMTPPort = "587"
        Send-MailMessage -From $From -to $To -Subject $Subject -Body $Body -BodyAsHtml -SmtpServer $SMTPServer -port $SMTPPort -UseSsl -Credential $user –DeliveryNotificationOption OnSuccess -Encoding UTF8 
    }
    Catch {
        Write-Error "Não foi possível enviar o e-mail para $to : $_"
        Exit-Program
    }
}

Function Get-Layout {
    Param($name, $yields = @{})
    Try {
        $content = Get-Content -Encoding UTF8 -Path "$PSScriptRoot\layout\$name.html"
        Foreach ($key in $yields.Keys) {
            $content = $content -Replace ":$($key):$", $yields[$key]
        }
        return $content
    }
    Catch {
        Write-Error "Não foi possível encontrar o layout: $name"
        Exit-Program
    }
    return ""
}

function Get-Hash($textToHash) {
    $hasher = new-object System.Security.Cryptography.SHA256Managed
    $toHash = [System.Text.Encoding]::UTF8.GetBytes($textToHash)
    $hashByteArray = $hasher.ComputeHash($toHash)
    foreach ($byte in $hashByteArray) {
        $res += $byte.ToString()
    }
    return $res;
}