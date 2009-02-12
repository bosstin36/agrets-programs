@ECHO OFF
INSTSRV.EXE "Xero" "REMOVE"
INSTSRV.EXE "Xero" "I:\_GSERVER\_HOSTED_\Xero\WGserver.exe"
REG ADD "HKLM\SYSTEM\CurrentControlSet\Services\Xero" /v "Application" /t REG_SZ /d "I:\_GSERVER\_HOSTED_\Xero\WGserver.exe"
REG ADD "HKLM\SYSTEM\CurrentControlSet\Services\Xero" /v "Description" /t REG_SZ /d "User service created by Service Manager"
