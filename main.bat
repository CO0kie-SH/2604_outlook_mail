@chcp 65001>nul
@title IP 查看
@cd /d %~dp0
@set PYTHONIOENCODING=utf-8

@set path=D:\0Code2\py312;%path%
@set path=D:\job\py312\Scripts;D:\job\py312;%path%

set fei_title="播报标题"
:: set "only_work=Workday"

python main.py %*
