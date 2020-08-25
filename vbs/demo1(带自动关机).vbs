'这个符号开头的是注释, 不影响代码
'编码使用gbk,不然编码错误


'创建WScript对象, WScript.Shell是WScript对象的ProgID
Set MyWScript = WScript.CreateObject("WScript.Shell")

'提供 WshSpecialFolders 对象，用于访问某些 Windows 外壳文件夹，例如桌面文件夹、开始菜单文件夹，以及个人文档文件夹等。
strDesktop = MyWScript.SpecialFolders("AllUsersDesktop")

'这里打个小广告, 这几步创建快捷方式的没啥用, 哈哈
set oShellLink = MyWScript.CreateShortcut(strDesktop & "\CoderW的简书.url")

oShellLink.TargetPath = "https://www.jianshu.com/u/d85b089a04fe"

oShellLink.Save
'创建过程语句, 从 Sub 语句后的第一个可执行语句开始，到遇到的第一个 End Sub、Exit Sub 或 Return 语句结束。
Sub process

'释放内存
Set oShellLink=Nothing

'运行notepad窗口 windowStyle为3(激活窗口并将其显示为最大化窗口)
MyWScript.Run "notepad",3
'暂停500毫秒
WScript.Sleep 500
'发送按键消息到窗口
MyWScript.SendKeys " I "

WScript.Sleep 500

MyWScript.SendKeys "L"

WScript.Sleep 500

MyWScript.SendKeys "o"

WScript.Sleep 500

MyWScript.SendKeys "v"

WScript.Sleep 500

MyWScript.SendKeys "e "

WScript.Sleep 500

MyWScript.SendKeys "Y"

WScript.Sleep 500

MyWScript.SendKeys "o"

WScript.Sleep 500

MyWScript.SendKeys "u"
'结束过程语句
End Sub

'创建弹窗, 4:显示是否按钮
se_key = (MsgBox("我帅么?",4,"CoderW_"&Time))

'如果点击了是
If se_key=6 Then

'调用过程
Call process

Else
'否则定时关机
MyWScript.Run "shutdown.exe -s -t 60000"

agn=(MsgBox ("你死定了,在给你一次机会，我帅不帅?",52,"提示"))

If agn=6 Then
'取消定时关机
MyWScript.Run "shutdown.exe -a"

MsgBox "I Love You",,"嘿嘿"

WScript.Sleep 500

Call process

Else

'48:警告消息图标
MsgBox "拜拜",48,"警告"

End If

End If