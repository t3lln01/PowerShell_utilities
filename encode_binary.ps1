clear-host
 
# Encode
 
$FilePath = $args[0]
$File = [System.IO.File]::ReadAllBytes($FilePath);
 
# returns the base64 string
$Base64String = [System.Convert]::ToBase64String($File);
echo $Base64String `n
 
# Decode
 

#function Convert-StringToBinary {
#[CmdletBinding()]
#param (
#[string] $EncodedString
#, [string] $FilePath = (‘{0}\{1}’ -f $env:TEMP, [System.Guid]::NewGuid().ToString())
#)
 
#try {
#if ($EncodedString.Length -ge 1) {
 
## decodes the base64 string
#$ByteArray = [System.Convert]::FromBase64String($EncodedString);
#[System.IO.File]::WriteAllBytes($FilePath, $ByteArray);
#}
#}
#catch {
#}
# 
#Write-Output -InputObject (Get-Item -Path $FilePath);
#}