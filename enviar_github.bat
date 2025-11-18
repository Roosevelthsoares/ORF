@echo off
echo ====================================
echo   Enviando para GitHub
echo ====================================
echo.

set /p USUARIO="Digite seu usuario do GitHub: "
set /p REPO="Digite o nome do repositorio: "

echo.
echo Configurando remote...
git remote add origin https://github.com/%USUARIO%/%REPO%.git

echo.
echo Enviando codigo para o GitHub...
git push -u origin main

echo.
echo ====================================
echo   Codigo enviado com sucesso!
echo ====================================
echo.
echo Acesse: https://github.com/%USUARIO%/%REPO%
echo.
pause
