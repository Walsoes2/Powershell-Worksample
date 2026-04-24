param(

 [Parameter(Mandatory=$true, Position=0)]
    [string]$FilewithCHECKSUM,
 [Parameter(Mandatory=$true, Position=1)]
    [string]$Filesindottxt
)

Write-host "Usage .\GetScomFileHashes.ps1 -filewithCHECKSUM <PATH-to-txt> -Filesindottxt <PATH-to-txtlist>" -ForegroundColor Yellow

$answer = Read-Host "Continue? y/n"

if($answer.ToLower() -eq "n") { EXIT } 
else {

# Define the path to the ZIP file
$FilewithCHECKSUM = get-content $FilewithCHECKSUM
$FilePath = Get-Content $Filesindottxt


if($FilewithCHECKSUM.Length -eq $FilePath.lenght) {

Write-host "Number of checksums equal to the number of files: $FilewithCHECKSUM.Length to $FilePath.lenght. Using either length."
$compLength = $FilePath.Length
} elseif ($FilewithCHECKSUM.Length -gt $FilePath.lenght)

{
Write-host "Number of checksums are greater than number of files: $FilewithCHECKSUM.Length to $FilePath.lenght. Using the shorter one."
$compLength = $FilePath.Length
}
else
{

Write-host "Number of checksums are less than the amount of files to compare: $FilewithCHECKSUM.Length to $FilePath.lenght. Using the length of the checksumarray."

$compLength = $FilewithCHECKSUM.Length

}


    for ($i = 0; $i -lt $compLength; $i++) {
    $actualHash = (Get-FileHash -Path $filePath[$i] -Algorithm SHA256).Hash
    $expectedHash = $FilewithCHECKSUM[$i]

    if ($actualHash -eq $expectedHash) {
        Write-Host "[$($filePaths[$i])] ✅ Match"
    } else {
        Write-Host "[$($filePaths[$i])] ❌ Mismatch"
    }
}

}