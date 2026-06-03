import subprocess
import sys

# 设置控制台编码为 UTF-8，避免中文路径乱码
sys.stdout.reconfigure(encoding='utf-8')

SCRIPT_1 = r"D:\_ExcelUtil\FileServer空间增长记录（10点）\OFF盘连接报错时清除注册表.bat"
# 原 bug: f-string 中 \a 被转义为响铃字符 (\x07)，要用 raw string 拼接
subprocess.run(r'runas /user:czhkc\administrator /savecred ' + SCRIPT_1, shell=True)
SCRIPT_2 = r"D:\_ExcelUtil\FileServer空间增长记录（10点）\只需要点它.bat"
subprocess.run(r'cmd /c start "" ' + SCRIPT_2, shell=True)
