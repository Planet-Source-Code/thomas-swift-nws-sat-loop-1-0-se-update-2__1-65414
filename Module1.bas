Attribute VB_Name = "Module1"
'****************************************************************
'
' Live Program Update Code
'
' Written by:  Blake B. Pell
'              blakepell@hotmail.com
'              bpell@indiana.edu
'              http://www.blakepell.com
'              December 7, 2000
'
' This code is open source, I would appreciate that anybody using
' this is a released application to e-mail or get in contact with
' me.  I hope this makes someone's day easier or helps them learn
' a bit.
'
'
'****************************************************************

Global myVer As String
Global status$
Global UpdateTime As Integer
Public Declare Function Beep Lib "kernel32" _
  (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
'*******
' Written by:  Thomas A. Swift 2006
'*******
Public Function GetNWSFileList(Inet1 As Inet, myURL As String, RootLoc As String) As Boolean
    On Local Error GoTo 100
    Dim myData As String
    Dim sParts() As String
    Dim NewStr As String
    Dim RightPos As Integer
    If Inet1.StillExecuting = True Then Exit Function
    myData = Inet1.OpenURL(myURL, icString)

       OurBuffer = Replace(myData, Chr(10), Chr(32))
        'OurBuffer = Replace(OurBuffer, Chr(10), Chr(32))
        sParts() = Split(OurBuffer, Chr(32)) 'vbCrLf
        For i = 0 To UBound(sParts) - 1
        If InStr(1, sParts(i), "img", vbTextCompare) > 0 Then
        Form1.List1.AddItem RootLoc & sParts(i)
        Debug.Print RootLoc & sParts(i)
        GetNWSFileList = True
        End If
        Next i
        Exit Function

Error handler
100:
GetNWSFileList = False
Resume 105
105 End Function
'*******

Public Function GetInternetFile(Inet1 As Inet, myURL As String, DestDIR As String) As Boolean
    ' Written by: Blake Pell
    
    On Local Error GoTo 100
    
    Dim myData() As Byte
    If Inet1.StillExecuting = True Then Exit Function
    myData() = Inet1.OpenURL(myURL, icByteArray)


    For X = Len(myURL) To 1 Step -1
        If Left$(Right$(myURL, X), 1) = "/" Then RealFile$ = Right$(myURL, X - 1)
    Next X
    myFile$ = DestDIR + "\" + RealFile$
    Open myFile$ For Binary Access Write As #1
    Put #1, , myData()
    Close #1
    
    GetInternetFile = True
    Exit Function

' error handler
100:
    GetInternetFile = False
    Resume 105
105 End Function
Public Property Get GetFileName(file As String) As String
    Dim m As Long
    Dim GetChr0 As String
    Dim GetChr1 As String
    GetFileName = ""
    For m = 1 To Len(file)
        GetChr0 = Right(file, m)
        GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then
            GetFileName = Right(GetChr0, m - 1): Exit Property
        End If
    Next m
End Property
