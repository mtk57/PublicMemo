$username = $env:TEMP_FTP_USER
$password = $env:TEMP_FTP_PASSWORD
$server = $env:TEMP_FTP_SERVER
$remoteFile = $env:TEMP_FTP_REMOTE_FILE
$localFile = $env:TEMP_FTP_LOCAL_FILE
$uri = "ftp://${username}:${password}@${server}${remoteFile}"
try {
  Invoke-WebRequest -Uri $uri -OutFile $localFile -UseBasicParsing
  Write-Host "FTP�t�@�C���̃_�E�����[�h���������܂����B"
} catch {
  Write-Error "FTP�t�@�C���̃_�E�����[�h�Ɏ��s���܂���: $($_.Exception.Message)"
  exit 1
}
