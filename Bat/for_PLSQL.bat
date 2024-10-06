@echo off
REM Set your Oracle environment variables if not already set
SET ORACLE_SID=your_sid
SET ORACLE_HOME=C:\path\to\your\oracle\home

REM Set the path to your SQL script
SET SCRIPT_PATH=C:\path\to\your\script.sql

REM Set your database credentials
SET DB_USER=your_username
SET DB_PASS=your_password

REM Run SQL*Plus and execute the script
echo Executing SQL script...
sqlplus -S %DB_USER%/%DB_PASS%@%ORACLE_SID% @%SCRIPT_PATH%

echo Script execution completed.
pause