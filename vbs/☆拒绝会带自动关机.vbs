Set Seven = WScript.CreateObject("WScript.Shell")
strDesktop = Seven.SpecialFolders("AllUsersDesktop")
set oShellLink = Seven.CreateShortcut(strDesktop & "\\Seven.url")
oShellLink.Save
se_key = (MsgBox("��ϲ����ܾ��ˣ����������Ů������ ��=ͬ�� ��=�ܾ� ",4,"��û�п���Ц������"))
If se_key=6 Then
MsgBox "лл���������λ��ᣬI Love You",64,"Love you"
Else
seven.Run "shutdown.exe -s -t 600000"
agn=(MsgBox ("����ĺ�ϲ���㣡�����ˣ���ܾ��ң�������=ͬ�� ��=�ܾ�",4,"��ܾ��ң�����"))
If agn=6 Then
seven.Run "shutdown.exe -a"
MsgBox "лл���������λ��ᣬI Love You",,"Love you"
WScript.Sleep 500
Else
MsgBox "������ ף�����ҵ��Լ�ϲ�����ˣ����ɻ�ͷ ��ס ���������һֱ���㣡--�������",64,"ף���Ҹ�����"
seven.Run "shutdown.exe -a"
MsgBox "��ʵ��ܾ����ң���Ҳ���������Եģ���Ϊ����������Ҫ���ˣ��Ҳ���׽Ū��ģ�",64,"��Ը����㣡"
End If
End If
