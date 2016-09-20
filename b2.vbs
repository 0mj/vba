Sub loopAndPaste()
  Dim notepad As Integer
  notepad = Shell("C:\windows\system32\notepad.exe", 1)
   
    Worksheets("Sheet1").Activate
    Range("B1").Select
    Do Until IsEmpty(ActiveCell)
       Worksheets("Sheet1").Activate
       ActiveCell.Offset(1, 0).Range("A1").Select
       Selection.Copy
       Call AppActivate(notepad, False)
       Call sendKeys("^v", True)
       Call sendKeys("{ENTER}", True)
    Loop
End Sub