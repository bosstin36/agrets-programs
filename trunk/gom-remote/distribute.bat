@mkdir GOMRemote\
@mkdir GOMRemote\Source\
@del GOMRemote\* /Q /S
@copy *.frm GOMRemote\Source\
@copy *.frx GOMRemote\Source\
@copy *.ico GOMRemote\Source\
@copy *.vbw GOMRemote\Source\
@copy *.vbp GOMRemote\Source\
@copy GOMRemote.exe GOMRemote\
@copy example.bat GOMRemote\
@cls
@echo Done.
@echo Please compress GOMRemote folder.
@pause