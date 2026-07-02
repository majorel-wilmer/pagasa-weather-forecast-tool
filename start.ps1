$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root
if (-not (Test-Path '.venv\Scripts\python.exe')) {
  py -3 -m venv .venv
  .\.venv\Scripts\python.exe -m pip install -r requirements.txt
}
Start-Process 'http://127.0.0.1:8877'
& .\.venv\Scripts\python.exe -m uvicorn app:app --host 127.0.0.1 --port 8877
