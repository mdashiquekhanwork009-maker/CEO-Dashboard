@echo off

set PYTHON_PATH="C:\Users\E36250444\AppData\Local\Python\pythoncore-3.14-64\python.exe"

echo Running Python scripts sequentially...
echo.

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\demand.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\submission.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\interview.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\selection.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\selction_pipeline.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\onboarding.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\exit.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\exit_pipeline.py"
if errorlevel 1 goto error

call %PYTHON_PATH% "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\activeHeadCount.py"
if errorlevel 1 goto error

echo.
echo ✅ All scripts executed successfully!
echo Closing in 5 seconds...
timeout /t 5 /nobreak >nul
exit

:error
echo.
echo ❌ Error occurred while running scripts!
pause
exit