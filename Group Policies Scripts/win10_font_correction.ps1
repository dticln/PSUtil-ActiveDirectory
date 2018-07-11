$VERSION = '1.0.0.0'
$local_path = 'C:\DTI Services\Corrections'
$file_path = "$local_path\font_version" 

Function ExecuteFontCorrection
{
    (Get-Item c:\Windows\Fonts).Attributes = "Normal" 
    takeown.exe /F c:\windows\fonts /A /R 
    icacls "$env:SystemRoot\fonts" /grant AD.UFRGS.BR\"Domain Users":M /t 
    $acl = Get-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" 
    $newAccess = New-Object System.Security.AccessControl.RegistryAccessRule ("AD.UFRGS.BR\Domain Users","FullControl","Allow") 
    $acl.SetAccessRule($newAccess) 
    Set-Acl "HKLM:\Software\Microsoft\Windows NT\CurrentVersion\Fonts" $acl
}

If (-Not (Test-Path -Path $local_path))
{
    New-Item -ItemType Directory -Path $local_path -Force
}

If (-Not (Test-Path -Path $file_path -PathType Leaf))
{
    $VERSION | Out-File $file_path
    ExecuteFontCorrection
}

$installed_version = Get-Content -Path $file_path
If ($VERSION -ne $installed_version)
{
    $VERSION | Out-File $file_path
    ExecuteFontCorrection
}