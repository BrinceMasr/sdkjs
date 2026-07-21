CD /D %~dp0
call npm ci

REM set SDK_ADDONS=..\..\sdkjs-forms;..\..\sdkjs-ooxml
REM set SDK_PLATFORM=desktop
set NODE_ENV=development
call npm run build

pause
