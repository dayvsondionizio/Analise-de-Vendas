@echo off
echo Iniciando BI Fiscal... aguarde.
"C:\Users\Contador de Padarias\AppData\Local\Programs\Python\Python313\Scripts\streamlit.exe" run "%~dp0app.py" --server.port=8501
pause
