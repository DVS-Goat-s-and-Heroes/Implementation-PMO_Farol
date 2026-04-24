@echo off
chcp 65001 > nul
cls
 
echo ==========================================
echo      INICIANDO ANALISADOR LEGO 4.0
echo ==========================================
echo.
echo Buscando script em: C:\Users\mikhael.jorge\OneDrive - DSV\Minhas coisas\Projetos\ProjetoFaustoImplementationPlan\teste.2.0
echo.
 
python "C:\Users\mikhael.jorge\OneDrive - DSV\Minhas coisas\Projetos\ProjetoFaustoImplementationPlan\teste.2.0\farol_pmo.py"
 
echo.
echo ==========================================
if %errorlevel% neq 0 (
    echo [ERRO] Ocorreu um problema na execucao.
) else (
    echo [SUCESSO] Script finalizado.
)
echo ==========================================
echo.
pause