@echo off

FOR /F "tokens=1,2 delims==" %%G IN (build.properties) DO (set %%G=%%H)
set CATALINA_HOME=E:\apache-tomcat-7.0.50
set JAVA_HOME=C:\Program Files (x86)\Java\jre7
netstat -na | find "LISTENING" | find /C /I ":8080" > NUL
if %errorlevel%==0 goto :running

:running
call E:\apache-tomcat-7.0.50\bin\shutdown.bat
echo tomcat is SHUTTING DOWN end...................................
echo tomcat is running
goto :eof

:eof