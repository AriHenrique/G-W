if (-not (Get-Module -Name AWS.Tools.Common -ListAvailable)) {
    Install-Module -Name AWS.Tools.Common -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
}
if (-not (Get-Module -Name AWS.Tools.S3 -ListAvailable)) {
    Install-Module -Name AWS.Tools.S3 -Force -AllowClobber -Scope CurrentUser -ErrorAction SilentlyContinue
}
Import-Module AWS.Tools.Common
Import-Module AWS.Tools.S3
$accessKey = "id-iam-aws"
$secretKey = "key-iam-aws"
$bucketName = "exemplo-aula-si"
$localFilePath = "C:\SHARMAQ\SHOficina\dados.mdb"
$currentYear = Get-Date -Format "yyyy"
$currentMonth = Get-Date -Format "MM"
$currentDay = Get-Date -Format "dd"
$s3FilePath = "oficina/ano=$currentYear/mes=$currentMonth/dia=$currentDay/dados.mdb"
$credentials = New-Object Amazon.Runtime.BasicAWSCredentials -ArgumentList $accessKey, $secretKey
Initialize-AWSDefaultConfiguration -AccessKey $accessKey -SecretKey $secretKey -Region "us-east-1"
Write-S3Object -BucketName $bucketName -File $localFilePath -Key $s3FilePath -CannedACLName "private"
