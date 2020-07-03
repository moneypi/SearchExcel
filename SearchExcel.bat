@echo off 
if "%1" == "h" goto begin 
mshta vbscript:createobject("wscript.shell").run("%~nx0 h",0)(window.close)&&exit 
:begin

call C:\ProgramData\Miniconda3\Scripts\activate.bat C:\ProgramData\Miniconda3
python SearchExcel.py
