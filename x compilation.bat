start/w setversion.exe

pyinstaller -w -F -i "logo.ico" CorrectReport.py

xcopy %CD%\*.xltx %CD%\dist /H /Y /C /R
xcopy %CD%\*.ico %CD%\dist /H /Y /C /R
xcopy %CD%\*.ini %CD%\dist /H /Y /C /R

xcopy C:\vxvproj\tnnc-Excel\CorrectReportApp\CorrectReport\dist C:\vxvproj\tnnc-Excel\CorrectReportApp\ConsoleApp\ /H /Y /C /R
