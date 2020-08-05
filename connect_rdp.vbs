Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run("""mstsc""")
WScript.Sleep 500
WshShell.SendKeys "COMPUTER_NAME_IPADDRESSS"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 5000
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{TAB}"
WshShell.SendKeys "{ENTER}"
WScript.Sleep 200
WshShell.SendKeys "USERNAME"
WshShell.SendKeys "{TAB}"
WScript.Sleep 200
WshShell.SendKeys "PASSWORD"
WScript.Sleep 200
WshShell.SendKeys "{ENTER}"
