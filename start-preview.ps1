$py = (Get-Command python -ErrorAction SilentlyContinue).Source
if (-not $py) { $py = (Get-Command py -ErrorAction SilentlyContinue).Source }
if (-not $py) { Write-Output "PYTHON_NOT_FOUND"; exit 1 }
$cmd = '"' + $py + '" "C:\MyFolder\JAC\preview-server.py"'
$res = Invoke-CimMethod -ClassName Win32_Process -MethodName Create -Arguments @{ CommandLine = $cmd; CurrentDirectory = 'C:\MyFolder\JAC' }
Write-Output ("PID=" + $res.ProcessId + " RC=" + $res.ReturnValue)
