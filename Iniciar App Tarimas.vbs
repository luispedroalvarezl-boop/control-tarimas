Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "cmd /c cd /d ""%USERPROFILE%\Downloads"" && python -m streamlit run app.py", 0, False