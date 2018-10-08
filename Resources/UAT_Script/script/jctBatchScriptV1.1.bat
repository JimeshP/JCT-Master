@echo on
FOR /F "tokens=1,2 delims==" %%G IN (build.properties) DO (set %%G=%%H)
set CATALINA_HOME=E:\apache-tomcat-7.0.50
set JAVA_HOME=C:\Program Files (x86)\Java\jre7

netstat -na | find "LISTENING" | find /C /I ":8080" > NUL
if %errorlevel%==0 goto :filecheck
echo .............%errorlevel%.....................
:gotorun
echo tomcat is not running..............%errorlevel%.....................
echo BAT - Before Ant run
echo %ant.build.file%
ant -f %ant.build.file%
echo BAT - After Ant run ...........%errorlevel%.............
goto :eof


:filecheck  
IF EXIST %ant.whitelebel.logo%  goto :running
IF EXIST %ant.whitelebel.commonStyle%  goto :running
goto :eof

:running
call E:\apache-tomcat-7.0.50\bin\shutdown.bat
echo tomcat is SHUTTING DOWN end................%errorlevel%...................
echo tomcat is running

goto :gotorun

:eof