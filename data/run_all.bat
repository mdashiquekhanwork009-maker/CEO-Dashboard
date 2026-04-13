@echo off

echo Running Python scripts sequentially...
echo.

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\demand.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\submission.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\interview.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\selection.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\selction_pipeline.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\onboarding.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\exit.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\exit_pipeline.py"
if errorlevel 1 goto error

call python "C:\Users\E36250444\OneDrive - JoulestoWatts Business Solutions Pvt Ltd\Desktop\Streamlit Dashboard\data\activeHeadCount.py"
if errorlevel 1 goto error

echo.
echo ✅ All scripts executed successfully!
echo Closing in 5 seconds...
timeout /t 5 /nobreak >nul
exit

:error
echo.
echo ❌ Error occurred while running scripts!
echo Check logs or run manually to debug.
pause
exit