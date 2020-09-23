Attribute VB_Name = "Module2"
Public Function RndNum() As String
Dim sAlpha As String, iLoop As Integer
Dim iRandNum As Integer, sMatch As String
Dim tmp As String

sAlpha = "0123456789"

Randomize


        tmp = ""
        For iLoop = 1 To 50
              iRandNum = Int((11 - 1 + 1) * Rnd + 1)
              sMatch = Mid(sAlpha, iRandNum, 1)
              tmp = tmp & sMatch
              If Len(tmp) = Len(frmRndOpt.Text1.Text) Then Exit For
        Next iLoop
        RndNum = tmp
      

End Function
