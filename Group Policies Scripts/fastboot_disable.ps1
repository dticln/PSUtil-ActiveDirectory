$VERSION = '1.0.0.0'
$local_path = 'C:\DTI Services\Corrections'
$file_path = "$local_path\hiberboot_version" 

Function ExecuteHiberbootCorrection
{
    REG ADD "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Power" /V HiberbootEnabled /T REG_dWORD /D 0 /F
}

If (-Not (Test-Path -Path $local_path))
{
    New-Item -ItemType Directory -Path $local_path -Force
}

If (-Not (Test-Path -Path $file_path -PathType Leaf))
{
    $VERSION | Out-File $file_path
    ExecuteHiberbootCorrection
}

$installed_version = Get-Content -Path $file_path
If ($VERSION -ne $installed_version)
{
    $VERSION | Out-File $file_path
    ExecuteHiberbootCorrection
}