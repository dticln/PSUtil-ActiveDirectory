$directories = Get-ChildItem '\\ad.ufrgs.br\LITORAL' -Directory -ErrorAction SilentlyContinue -Recurse
$array = @{}
Foreach ($directory In $directories.FullName) {
    $acl = Get-ChildItem $directory -ErrorAction SilentlyContinue | where-object {($_.PsIsContainer)} | Get-ACL -ErrorAction SilentlyContinue
    If ($acl) {
        Foreach($idRef In $acl.Access.IdentityReference){
            If ($idRef.Value -match 'AD\\[0-9]{8}' -and $acl.Owner -notcontains $idRef.Value) {
                If ($array.Contains($idRef.Value)){
                    If ($array[$idRef.Value] -notcontains $directory) {
                        $array[$idRef.Value] += $directory
                        $name = $directory.Substring($directory.LastIndexOf("\"))
                        Write-Output "Pasta `"...$name`" atribuída à `"$($idRef.Value)`"."
                    }
                } Else {
                    $list = @()
                    $list += $directory
                    $array.Add($idRef.Value, $list);
                    Write-Output "Usuário `"$($idRef.Value)`" possui acesso manual."
                }
            }
        }
    }
}


$out = @()
Foreach ($key In $array.Keys) {
    $item = $array[$key]
    $key = $key.Substring($key.Length - 8)
    Foreach ($sub in $item) {
        $obj = [PSCustomObject]@{
            Usuário = $key
            Local = $sub
        }
        $out += $obj
    }
}

$out | Export-Csv "output.csv" -Encoding UTF8 -NoTypeInformation