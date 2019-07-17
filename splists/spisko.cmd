goto end
rem ********************************  для отдела банковского надзора **********************************

del c:\AdminDir\SpisokKO\*.html /q
del c:\AdminDir\SpisokKO\*.log /q

rem Астахов А.Б. ВТС (223) 1010
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_ko_nadz-1month.ps1'"
C:\AdminDir\SpisokKO\exe\sps_fil-nadz-1month.exe 
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_fil-nadz-1month.ps1'"
C:\AdminDir\SpisokKO\exe\sps_fil-nadz-1month.exe
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_02-predstav.ps1'"
C:\AdminDir\SpisokKO\exe\sps_02-predstav
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_03-DopOffice.ps1'"
C:\AdminDir\SpisokKO\exe\sps_03-DopOffice
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_04-OperKassyVneKassUzla.ps1'"
C:\AdminDir\SpisokKO\exe\sps_04-OperKassyVneKassUzla
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_05-OperOffice.ps1'"
C:\AdminDir\SpisokKO\exe\sps_05-OperOffice
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_06-KreditnoKassOffice.ps1'"
C:\AdminDir\SpisokKO\exe\sps_06-KreditnoKassOffice
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_07-OutOfRegion.ps1'"
C:\AdminDir\SpisokKO\exe\sps_07-OutOfRegion
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_07C_TotalCount.ps1'"
C:\AdminDir\SpisokKO\exe\sps_07C_total_count
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_10OBNINSK.ps1'"
C:\AdminDir\SpisokKO\exe\sps_10OBNINSK
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_11Kaluga.ps1'"
C:\AdminDir\SpisokKO\exe\sps_11Kaluga
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_12LicenseRevoked.ps1'"
C:\AdminDir\SpisokKO\exe\sps_12LicenseRevoked
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_32-VSP.ps1'"
C:\AdminDir\SpisokKO\exe\sps_32-VSP
pause

net use g: /del /yes 
net use g: http://s29sps.region.cbr.ru/deprts/nadz/1 
copy c:\AdminDir\SpisokKO\*.html g:
del c:\AdminDir\SpisokKO\*.html /q
net use g: /del /yes
pause


rem ********************************  для отдела платежных систем и расчетов **********************************

del F:\AdminDir\SpisokKO\*.html /q
del F:\AdminDir\SpisokKO\08spisKO_Opsr.log /q
del F:\AdminDir\SpisokKO\09spisDopOffice_Opsr.log /q
del F:\AdminDir\SpisokKO\14spisKKOffice_Opsr.log /q
del F:\AdminDir\SpisokKO\15spisOperOffice_Opsr.log /q

rem Астахов А.Б. ВТС (223) 1010

%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_08ko_opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_09DopOffice_opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_14-KreditnoKassOffice-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_15-OperOffice-opsr.ps1'"

net use g: /del /yes 
net use g: http://portal.kaluga.cbr.ru/deprts/opsr/1 
copy F:\AdminDir\SpisokKO\*.html g:
del F:\AdminDir\SpisokKO\*.html /q
net use g: /del /yes

rem ====================   Кредитные организации по регионам   =================================
del F:\AdminDir\SpisokKO\*.html /q
del F:\AdminDir\SpisokKO\28KOKirov_Opsr.log /q
del F:\AdminDir\SpisokKO\29KOKaluga_Opsr.log /q
del F:\AdminDir\SpisokKO\30KOObninsk_Opsr.log /q
del F:\AdminDir\SpisokKO\31KOMaloyar_Opsr.log /q

rem Астахов А.Б. ВТС (223) 1010

rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_28-KO_Kirov-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_29-KO_Kaluga-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_30-KO_Obninsk-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_31-KO_Maloyar-opsr.ps1'"


net use g: /del /yes 
net use g: http://portal.kaluga.cbr.ru/deprts/opsr/5
copy F:\AdminDir\SpisokKO\*.html g:
del F:\AdminDir\SpisokKO\*.html /q
net use g: /del /yes

rem ====================   Доп. офисы по регионам   =================================

del F:\AdminDir\SpisokKO\*.html /q
del F:\AdminDir\SpisokKO\24DopOfficeKirov_Opsr.log /q
del F:\AdminDir\SpisokKO\25DopOfficeKaluga_Opsr.log /q
del F:\AdminDir\SpisokKO\26DopOfficeObninsk_Opsr.log /q
del F:\AdminDir\SpisokKO\27DopOfficeMaloyar_Opsr.log /q

rem Астахов А.Б. ВТС (223) 1010

rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_24-DopOfficeKirov-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_25-DopOfficeKaluga-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_26-DopOfficeObninsk-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_27-DopOfficeMaloyar-opsr.ps1'"


net use g: /del /yes 
net use g: http://portal.kaluga.cbr.ru/deprts/opsr/4
copy F:\AdminDir\SpisokKO\*.html g:
del F:\AdminDir\SpisokKO\*.html /q
net use g: /del /yes

rem ====================   Опер. офисы по регионам   =================================

del F:\AdminDir\SpisokKO\*.html /q
del F:\AdminDir\SpisokKO\16OperOfissKaluga_Opsr.log /q
del F:\AdminDir\SpisokKO\17OperOfissObninsk_Opsr.log /q
del F:\AdminDir\SpisokKO\18OperOfissMaloyar_Opsr.log /q
del F:\AdminDir\SpisokKO\19OperOfissKirov_Opsr.log /q

rem Астахов А.Б. ВТС (223) 1010

%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_16-OperOfficeKaluga-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_17-OperOfficeObninsk-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_18-OperOfficeMaloyar-opsr.ps1'"
rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_19-OperOfficeKirov-opsr.ps1'"


net use g: /del /yes 
net use g: http://portal.kaluga.cbr.ru/deprts/opsr/2 
copy F:\AdminDir\SpisokKO\*.html g:
del F:\AdminDir\SpisokKO\*.html /q
net use g: /del /yes

rem ====================   Кред-Касс. офисы по регионам   =================================

del F:\AdminDir\SpisokKO\*.html /q
del F:\AdminDir\SpisokKO\20KKofficeKirov_Opsr.log /q
del F:\AdminDir\SpisokKO\21KKofficeKaluga_Opsr.log /q
del F:\AdminDir\SpisokKO\22KKofficeObninsk_Opsr.log /q
del F:\AdminDir\SpisokKO\23KKofficeMaloyar_Opsr.log /q

rem Астахов А.Б. ВТС (223) 1010

rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_20-KreditnoKassOfficeKirov-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_21-KreditnoKassOfficeKaluga-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_22-KreditnoKassOfficeObninsk-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_23-KreditnoKassOfficeMaloyar-opsr.ps1'"


net use g: /del /yes 
net use g: http://portal.kaluga.cbr.ru/deprts/opsr/3 
copy F:\AdminDir\SpisokKO\*.html g:
del F:\AdminDir\SpisokKO\*.html /q
net use g: /del /yes

rem ============================ Передвижные пункты касовых операций ===============================

del F:\AdminDir\SpisokKO\*.html /q
del F:\AdminDir\SpisokKO\33VSPKirov_Opsr.log /q
del F:\AdminDir\SpisokKO\34VSPKaluga_Opsr.log /q
del F:\AdminDir\SpisokKO\35VSPObninsk_Opsr.log /q
del F:\AdminDir\SpisokKO\36VSPMaloyar_Opsr.log /q

rem Астахов А.Б. ВТС (223) 1010

rem %SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_33-VSP_Kirov-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_34-VSP_Kaluga-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_35-VSP_Obninsk-opsr.ps1'"
%SystemRoot%\system32\WindowsPowerShell\v1.0\powershell.exe "& 'F:\AdminDir\SpisokKO\sps_36-VSP_Maloyar-opsr.ps1'"


net use g: /del /yes 
net use g: http://portal.kaluga.cbr.ru/deprts/opsr/6 
copy F:\AdminDir\SpisokKO\*.html g:
del F:\AdminDir\SpisokKO\*.html /q
net use g: /del /yes


rem ==============================  T H E        E N D ========================================

:end
