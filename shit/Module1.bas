Attribute VB_Name = "Module1"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Public Sub dormir()
    
    Sleep (Form1.txtsleep.Text)
        
End Sub
