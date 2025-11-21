Param()

$ErrorActionPreference = "SilentlyContinue"

# Cerrar instancias de Word que puedan estar bloqueando .docx
Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue

# Opcional: espera breve para liberar locks
Start-Sleep -Milliseconds 300
