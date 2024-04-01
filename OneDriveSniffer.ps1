
If (!(Test-path "$HOME\OneDriveLib.dll" -ErrorAction SilentlyContinue))
{
    write-output "Attempting to download the OneDrive Module...";
    Invoke-WebRequest -Uri "https://github.com/rodneyviana/ODSyncService/releases/download/1.0.7796.25985/OneDriveLib.dll" -OutFile "$HOME\OneDriveLib.dll";
    Unblock-File -Path "$HOME\OneDriveLib.dll";
}

write-output "Attempting to import the OneDrive Module...";
Import-Module "$HOME\OneDriveLib.dll";

write-output "Processing current OneDrive Status Information...";
$status = Get-ODStatus
#Write-Output "$status"

if($status.DisplayName -ieq "OneDrive - Integrated Health Resources")
{
    Write-Output "OneDrive is configured properly! ($($status.DisplayName))";
    Exit;
}
else
{
    Write-Output "OneDrive status is '$($status.DisplayName)'...";
    Exit 1;
}
