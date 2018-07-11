$service_name = 'DTIServices'
$service_version = '1.1.0.7'
$installed_service = Get-Service -DisplayName 'DTIServices*'
$remote_path = '\\ad.ufrgs.br\SysVol\ad.ufrgs.br\Policies\{08F8E1D0-8677-405D-88E8-D8BB7B4FD4C6}\Machine\Scripts\Startup'
$local_path = 'C:\DTI Services\bin'

Function UninstallOldService
{
    Get-Service -Name "DTIServices" | Set-Service -Status Stopped
    c:\Windows\Microsoft.NET\Framework64\v4.0.30319\InstallUtil.exe /u "$local_path\DTIService.exe"
    Remove-Item -Recurse -Force 'C:\DTI Services'
}

Function InstallService
{
    New-Item -ItemType Directory -Path 'C:\DTI Services\bin'
    New-Item -ItemType Directory -Path 'C:\DTI Services\InstalledReport'
    $folder = Get-Item 'C:\DTI Services' -Force
    $folder.Attributes="Hidden"
    Copy-Item "$remote_path\DTIService.exe" -Destination "$local_path\DTIService.exe" -Recurse
    Copy-Item "$remote_path\DTIService.exe.config" -Destination "$local_path\DTIService.exe.config" -Recurse
    Copy-Item "$remote_path\DTIService.pdb" -Destination "$local_path\DTIService.pdb" -Recurse
    c:\Windows\Microsoft.NET\Framework64\v4.0.30319\InstallUtil.exe "$local_path\DTIService.exe"
    Get-Service -Name "DTIServices" | Set-Service -StartupType Automatic
    Get-Service -Name "DTIServices" | Set-Service -Status Running
}

If ($installed_service) 
{
    $installed_version = Get-Content -Path 'C:\DTI Services\version.txt'
    If ($service_version -ne $installed_version)
    {
        UninstallOldService
        InstallService
    }
} Else 
{
    InstallService
}
