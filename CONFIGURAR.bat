@echo off
setlocal
cd /d "%~dp0"
echo.
echo Cole o endereco onde voce publicou os arquivos do Anexa
echo (exemplo: https://seuusuario.github.io/anexa)
echo.
set /p URL=Endereco: 
powershell -NoProfile -ExecutionPolicy Bypass -Command "$u=$env:URL.Trim().TrimEnd('/'); if($u -notmatch '^https://'){ Write-Host 'O endereco precisa comecar com https://'; exit 1 }; (Get-Content -Raw -Encoding UTF8 'manifest.xml').Replace('https://ENDERECO-DO-ANEXA',$u) | Set-Content -Encoding UTF8 'manifest-pronto.xml'; Write-Host ''; Write-Host 'Pronto: manifest-pronto.xml criado nesta pasta.'"
echo.
pause
