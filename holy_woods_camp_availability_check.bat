rem change character code to UTF-8 (When the path doesn't include Japanese this is not necessary, but will do this very just in case)
chcp 65001

rem to activate target anaconda virtual environment
call C:\Users\tomom\anaconda3\Scripts\activate.bat

rem activate target environment which is holy_woods_camp_availability
call activate holy_woods_camp_availability

rem execute python with target env's python.exe
C:\Users\tomom\anaconda3\envs\holy_woods_camp_availability\python.exe holy_woods_camp_availability_check.py True

rem to keep bat window
pause
