@ECHO OFF
del *.vtf
cd "test cases"
copy * ..
cd ..
cls
tf2-materials.exe *
pause