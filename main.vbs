msgbox("Paste Pro 3.0 Developed by Zihan Lu in 2022")
n = inputbox("����")
t = inputbox("��Ъʱ�䣨��λΪ���룩")
info = inputbox("�Ƿ����������y��n�� ճ������3���ʼ����ȷ�������Ѹ���")
Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.AppActivate""
WScript.Sleep 3000
if info="y" then
    for i= 1 to n
    WScript.Sleep t
    Wshshell.SendKeys"^v"
    Wshshell.SendKeys i
    Wshshell.SendKeys"%s"
    Next
else
    for i= 1 to n
    WScript.Sleep t
    Wshshell.SendKeys"^v"
    Wshshell.SendKeys"%s"
    Next
end if
