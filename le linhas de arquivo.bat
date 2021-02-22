@echo off
for /f "tokens=*" %%a in (arc.config) do (
  echo %%a
)
pause