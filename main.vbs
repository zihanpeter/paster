msgbox("Paste Pro 3.0 Developed by Zihan Lu in 2022")
n = inputbox("次数")
t = inputbox("间歇时间（单位为毫秒）")
info = inputbox("是否计数（输入y或n） 粘贴将在3秒后开始，请确保内容已复制")
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
