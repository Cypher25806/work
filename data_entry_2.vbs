Set WshShell = WScript.CreateObject("WScript.Shell")
WScript.Sleep(5000)
Dim adminUnit  
        adminUnit = InputBox("Input admin unit:")
        Dim numLocal
            numLocal = InputBox("Input local:")
for i = 0 to 50
    Dim sInput 
        sInput = InputBox("Alphanumerical code of Level:")
        Dim sInput1
            sInput1 = InputBox("Quantity of Products in that level:")
    wshShell.AppActivate "grm"
    for j = 0 to (CInt(sInput1)-1)
        WScript.Sleep(100)
        WshShell.SendKeys "100171"
        WScript.Sleep(100)
        WshShell.SendKeys "{Tab}"
        WScript.Sleep(100)
        WshShell.SendKeys adminUnit
        WScript.Sleep(100)
        WshShell.SendKeys "{Tab}"
        WScript.Sleep(100)
        WshShell.SendKeys numLocal
        WScript.Sleep(100)
        WshShell.SendKeys "{Tab}"
        WScript.Sleep(100)
        WshShell.SendKeys "10000"
        WScript.Sleep(100)
        'enter product number'
        Dim sInput2
        sInput2 = InputBox("Input product number:")
        WshShell.SendKeys sInput2
        WScript.Sleep(100) 
        WshShell.SendKeys "{Tab}"
        WshShell.SendKeys "{Tab}"
        WshShell.SendKeys "{Tab}"        
        WScript.Sleep(100)
        'input quantity of products'
        Dim sInput3
            sInput3 = InputBox("Input quantity of product in loc:")
        'enter the quantity'
        WshShell.SendKeys sInput3
        WScript.Sleep(100)
        WshShell.SendKeys "{Tab}"
        WshShell.SendKeys "{Tab}"    
        WScript.Sleep(100)
        Dim sInput4 
            sInput4 = InputBox("input double or single case:(d or b)")
        WshShell.SendKeys sInput
        WshShell.SendKeys "."
        WScript.Sleep(100)
        WshShell.SendKeys sInput4
        WScript.Sleep(100)
        WScript.Echo "[ENTER]"
        WScript.Sleep(100)
        WshShell.SendKeys "{Tab}"    
    Next
Next