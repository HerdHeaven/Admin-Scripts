param ([PARAMETER(Mandatory=$FALSE,ValueFromPipeline=$FALSE)]
[string]$FilePath,
[PARAMETER(Mandatory=$FALSE,ValueFromPipeline=$FALSE)]
[ValidateSet(“MACTripleDES”,”MD5”,”RIPEMD160”, "SHA1", "SHA256", "SHA384","SHA512")]
[string]$Algorithm="SHA256")
BEGIN {

    Add-Type -AssemblyName System.Windows.Forms

    $InitialDirectory = "$env:USERPROFILE\Downloads"

    function Get-FileName($InitialDirectory)
    {
        $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $OpenFileDialog.InitialDirectory = $InitialDirectory
        $OpenFileDialog.ShowDialog() | Out-Null
        $OpenFileDialog.FileName
    }


  }
PROCESS{

    If($FilePath -eq "")
    {$FilePath = Get-FileName $InitialDirectory}
    
    $CalcHash = get-filehash -Path $($FilePath) -Algorithm SHA256

    $Hash = Read-Host -Prompt "Please enter the $Algorithm Hash value to verify"

    if($CalcHash.Hash -eq $Hash)
    {Write-Output "The caluclated hash value matches!"}
    else
    {Write-Warning "The calculated has value does not match!`nProvided Hash:$hash`nCalculated Hash:$($calchash.hash)`n`n"}
    
  }
END {
    PAUSE
    Clear-Variable InitialDirectory, hash, calchash
}



