# Activar-Office-2019

#abrir cmd con privilegios de administrador

# office 32 / 64 bits
cd /d %ProgramFiles%\Microsoft Office\Office16

cd /d %ProgramFiles(x86)%\Microsoft Office\Office16


# copiar y enter
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"

cscript ospp.vbs /setprt:1688

cscript ospp.vbs /unpkey:6MWKP >nul

cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP

cscript ospp.vbs /sethst:kms8.msguides.com

cscript ospp.vbs /act
