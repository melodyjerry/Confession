'������ſ�ͷ����ע��, ��Ӱ�����
'����ʹ��gbk,��Ȼ�������


'����WScript����, WScript.Shell��WScript�����ProgID
Set MyWScript = WScript.CreateObject("WScript.Shell")

'�ṩ WshSpecialFolders �������ڷ���ĳЩ Windows ����ļ��У����������ļ��С���ʼ�˵��ļ��У��Լ������ĵ��ļ��еȡ�
strDesktop = MyWScript.SpecialFolders("AllUsersDesktop")

'������С���, �⼸��������ݷ�ʽ��ûɶ��, ����
set oShellLink = MyWScript.CreateShortcut(strDesktop & "\CoderW�ļ���.url")

oShellLink.TargetPath = "https://www.jianshu.com/u/d85b089a04fe"

oShellLink.Save
'�����������, �� Sub ����ĵ�һ����ִ����俪ʼ���������ĵ�һ�� End Sub��Exit Sub �� Return ��������
Sub process

'�ͷ��ڴ�
Set oShellLink=Nothing

'����notepad���� windowStyleΪ3(����ڲ�������ʾΪ��󻯴���)
MyWScript.Run "notepad",3
'��ͣ500����
WScript.Sleep 500
'���Ͱ�����Ϣ������
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
'�����������
End Sub

'��������, 4:��ʾ�Ƿ�ť
se_key = (MsgBox("��˧ô?",4,"CoderW_"&Time))

'����������
If se_key=6 Then

'���ù���
Call process

Else
'����ʱ�ػ�
MyWScript.Run "shutdown.exe -s -t 60000"

agn=(MsgBox ("��������,�ڸ���һ�λ��ᣬ��˧��˧?",52,"��ʾ"))

If agn=6 Then
'ȡ����ʱ�ػ�
MyWScript.Run "shutdown.exe -a"

MsgBox "I Love You",,"�ٺ�"

WScript.Sleep 500

Call process

Else

'48:������Ϣͼ��
MsgBox "�ݰ�",48,"����"

End If

End If