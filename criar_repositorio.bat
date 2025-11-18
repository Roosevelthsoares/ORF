@echo off
echo ====================================
echo   Inicializando Repositorio Git
echo ====================================
echo.

git init
if errorlevel 1 (
    echo ERRO: Git nao encontrado. Reinicie o terminal ou adicione Git ao PATH.
    pause
    exit /b 1
)

echo.
echo Adicionando arquivos...
git add .

echo.
echo Fazendo commit inicial...
git commit -m "Initial commit: Extrator Starlink automatizado"

echo.
echo Definindo branch principal como 'main'...
git branch -M main

echo.
echo ====================================
echo   Repositorio criado com sucesso!
echo ====================================
echo.
echo Proximos passos:
echo 1. Crie um novo repositorio no GitHub (https://github.com/new)
echo 2. Execute os comandos abaixo (substitua SEU_USUARIO e NOME_REPO):
echo.
echo    git remote add origin https://github.com/SEU_USUARIO/NOME_REPO.git
echo    git push -u origin main
echo.
pause
