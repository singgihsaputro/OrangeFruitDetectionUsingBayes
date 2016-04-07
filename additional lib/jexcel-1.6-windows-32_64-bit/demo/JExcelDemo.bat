@echo off

REM Copyright (c) 2002-2007 TeamDev Ltd. All rights reserved.

set ROOT_DIR=..

set LIB_DIRECTORY=%ROOT_DIR%/lib
set BIN_DIRECTORY=%ROOT_DIR%/bin

set CORE_CLASSES=%LIB_DIRECTORY%/jniwrap-3.8.4.jar;%LIB_DIRECTORY%/winpack-3.8.4.jar;%LIB_DIRECTORY%/comfyj-2.10.jar;%LIB_DIRECTORY%/jexcel-1.6.jar;%LIB_DIRECTORY%/slf4j-api-1.5.8.jar;%LIB_DIRECTORY%/slf4j-simple-1.5.8.jar
set CUSTOM_CLASSES=%ROOT_DIR%/demo/jexceldemo.jar;

set SAMPLE_CLASSPATH=%CORE_CLASSES%;%CUSTOM_CLASSES%
set SAMPLE_MAINCLASS=com.jniwrapper.win32.jexcel.samples.demo.JExcelDemo

set PATH=%BIN_DIRECTORY%;%PATH%

start "" javaw.exe -classpath %SAMPLE_CLASSPATH% %SAMPLE_MAINCLASS%
