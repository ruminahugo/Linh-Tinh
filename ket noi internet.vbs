strComputer = "." ' Sử dụng "." cho máy tính cục bộ

Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE NetConnectionID IS NOT NULL")

For Each objNetAdapter in colNetAdapters
    objNetAdapter.Disable()
    WScript.Sleep 1000 ' Tạm dừng 1 giây
    objNetAdapter.Enable()
Next
