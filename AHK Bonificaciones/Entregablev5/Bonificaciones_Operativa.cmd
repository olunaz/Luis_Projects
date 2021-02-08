:: ---------------------------------------------------------------------
:: |			  INSTALADOR SIMPLE DE ROBOT AUTOHOTKEY   			   |
:: ---------------------------------------------------------------------
:: Verificar si D: existe - 
IF EXIST "D:" (set location=D:) ELSE (set location=C:)
set myname=%~n0
set myfile=%cd%
:: Crear scaffolding en location
mkdir %location%\AHK\%myname%
set destino=%location%\AHK\%myname%
set mypath=%~dp0
set origen=%cd%\%myname%
xcopy %origen% %destino% /s

:: Navegar a location
%location%
cd %location%\AHK\%myname%

:: Instalar el certificado
<nul call certutil -f -user -importpfx AHKLiveCertificate.p12 NoRoot

:: Llenar el archivo rutas.properties con la ruta absoluta
echo PcDiro=%USERPROFILE%\Desktop> Routes.properties
:: No hay manera de crearle acceso directo sin permisos de admin
:: usuario debe de crearlo!
set namead=%myname%.lnk
set namexe=%myname%.exe
::shortcut /F:"%USERPROFILE%\Desktop\AHK Mesa de Control Riesgos - Manual de Usuario.lnk" /A:C /T:"%location%\AHK\Mesa_Control_Riesgos\[AHK Mesa de Control] - [Riesgos] - Manual de Usuario v3.5.docx"
shortcut /F:"%USERPROFILE%\Desktop\%namead%" /A:C /T:"%location%\AHK\%myname%\%namexe%"
start %cd%\Routes.exe