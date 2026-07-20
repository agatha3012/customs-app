Set WshShell = CreateObject("WScript.Shell")

' 设置环境变量，确保能找到 node
Dim env
Set env = WshShell.Environment("PROCESS")
env("PATH") = "C:\Program Files\nodejs;" & env("PATH")

WshShell.CurrentDirectory = "C:\Users\LY\customs-app"

' 使用完整路径启动，并将日志写入文件方便排查
Dim cmd
cmd = """C:\Program Files\nodejs\node.exe"" ""C:\Users\LY\customs-app\server.js"""
WshShell.Run "cmd /c " & cmd & " > ""C:\Users\LY\customs-app\startup.log"" 2>&1", 0, False
