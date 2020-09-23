VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "NWS Sat. Image Loops 1.0 SE"
   ClientHeight    =   2760
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   2760
   ScaleWidth      =   6675
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1590
      Top             =   1530
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   1515
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   615
      Top             =   1515
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1455
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   1485
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.PictureBox Picture1 
      Height          =   1380
      Left            =   -15
      ScaleHeight     =   1320
      ScaleWidth      =   1350
      TabIndex        =   0
      Top             =   30
      Width           =   1410
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   90
      Top             =   1530
   End
   Begin VB.Menu MnuMapLoc 
      Caption         =   "Map Location"
      Begin VB.Menu MnuContenentMaps 
         Caption         =   "U.S. Contenent Map's"
         Begin VB.Menu MnuUSVisible 
            Caption         =   "U.S. Visible"
         End
         Begin VB.Menu MnuUSWaterVapor 
            Caption         =   "U.S. Water Vapor"
         End
         Begin VB.Menu MnuUSInfrared 
            Caption         =   "U.S. Infrared"
         End
         Begin VB.Menu MnuUSInfraredCh2 
            Caption         =   "U.S. Infrared Ch. 2"
         End
      End
      Begin VB.Menu MnuWestCoast 
         Caption         =   "West Coast"
      End
      Begin VB.Menu MnuEastCoast 
         Caption         =   "East Coast"
      End
      Begin VB.Menu MnuGulfofMexico 
         Caption         =   "Gulf of Mexico"
      End
      Begin VB.Menu MnuCaribbean 
         Caption         =   "Caribbean"
      End
      Begin VB.Menu MnuWestCentralPacific 
         Caption         =   "West Central Pacific"
      End
      Begin VB.Menu MnuCentralPacific 
         Caption         =   "Central Pacific"
      End
      Begin VB.Menu MnuEastPacific 
         Caption         =   "East Pacific"
      End
      Begin VB.Menu MnuWestPacific 
         Caption         =   "West Pacific"
      End
      Begin VB.Menu MnuNorthwestPacific 
         Caption         =   "Northwest Pacific"
      End
      Begin VB.Menu MnuNortheastPacific 
         Caption         =   "Northeast Pacific"
      End
      Begin VB.Menu MnuWestAtlantic 
         Caption         =   "West Atlantic"
      End
      Begin VB.Menu MnuCentralAtlantic 
         Caption         =   "Central Atlantic"
      End
      Begin VB.Menu MnuEastAtlanticMET8 
         Caption         =   "East Atlantic (MET-8)"
      End
      Begin VB.Menu MnuNorthwestAtlantic 
         Caption         =   "Northwest Atlantic"
      End
      Begin VB.Menu MnuNorthAtlantic 
         Caption         =   "North Atlantic"
      End
      Begin VB.Menu MnuNortheastAtlanticMET8 
         Caption         =   "Northeast Atlantic (MET-8)"
      End
   End
   Begin VB.Menu MnuMapType 
      Caption         =   "Map Type"
      Begin VB.Menu MnuMapTypeVisible 
         Caption         =   "Visible"
      End
      Begin VB.Menu MnuMapTypeShortwave 
         Caption         =   "Shortwave"
      End
      Begin VB.Menu MnuMapTypeWaterVapor 
         Caption         =   "Water Vapor "
      End
      Begin VB.Menu MnuMapTypeInfraredNormal 
         Caption         =   "Infrared Normal"
      End
      Begin VB.Menu MnuMapTypeInfraredAVN 
         Caption         =   "Infrared AVN"
      End
      Begin VB.Menu MnuMapTypeInfraredDvorak 
         Caption         =   "Infrared Dvorak"
      End
      Begin VB.Menu MnuMapTypeInfraredJSL 
         Caption         =   "Infrared JSL"
      End
      Begin VB.Menu MnuMapTypeInfraredRGB 
         Caption         =   "Infrared RGB"
      End
      Begin VB.Menu MnuMapTypeInfraredFunktop 
         Caption         =   "Infrared Funktop "
      End
      Begin VB.Menu MnuMapTypeInfraredRainbow 
         Caption         =   "Infrared Rainbow"
      End
   End
   Begin VB.Menu MnuAnimation 
      Caption         =   "Animation Control"
      Begin VB.Menu MnuSpeed 
         Caption         =   "Speed"
         Begin VB.Menu MnuSpeed1X 
            Caption         =   "1X"
         End
         Begin VB.Menu MnuSpeed2X 
            Caption         =   "2X"
         End
         Begin VB.Menu MnuSpeed3X 
            Caption         =   "3X"
         End
         Begin VB.Menu MnuSpeed4X 
            Caption         =   "4X"
         End
      End
      Begin VB.Menu MnuShowLatest 
         Caption         =   "Freeze On Latest Image"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Const WU_LOGPIXELSX = 88
Const WU_LOGPIXELSY = 90

Private PicX As Long
Private PicY As Long
Private RecoveryTime As Integer
Private DLCount As Integer
Private Shutdown As Boolean
Private Token As Long
Private CurPicFile As String
Private FreezeTime As Integer

Private SysTypeFiles As Boolean
Private ShowLatest As Boolean

Public MapType As String
Public MapRootLoc As String
Private Sub LoadPic()
Token = InitGDIPlus
Picture1.Picture = modGDIPlusResize.LoadPictureGDIPlus(CurPicFile, ConvertTwipsToPixels(Picture1.Width, 0), ConvertTwipsToPixels(Picture1.Height - 50, 0), , True)
FreeGDIPlus Token
End Sub
Private Sub Form_Load()
'MnuWestCoast.Checked = True
'MapRootLoc = "http://www.ssd.noaa.gov/goes/west/weus/"
MnuMapTypeVisible.Checked = True
MapType = "vis"
MnuSpeed1X.Checked = True
'Timer4.Enabled = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Shutdown = True
End Sub
Private Sub Form_Resize()
Picture1.Left = Me.ScaleLeft
Picture1.Top = Me.ScaleTop
Picture1.Width = Me.ScaleWidth
Picture1.Height = Me.ScaleHeight
If CurPicFile > "" Then LoadPic
End Sub
Private Sub Form_Unload(Cancel As Integer)
Call CleanUp
End Sub
Private Sub CleanUp()
On Error Resume Next
Kill App.Path & "\*.jpg"
End Sub
Private Sub UncheckAllMapTypes()
MnuMapTypeVisible.Checked = False
MnuMapTypeShortwave.Checked = False
MnuMapTypeWaterVapor.Checked = False
MnuMapTypeInfraredNormal.Checked = False
MnuMapTypeInfraredAVN.Checked = False
MnuMapTypeInfraredDvorak.Checked = False
MnuMapTypeInfraredJSL.Checked = False
MnuMapTypeInfraredRGB.Checked = False
MnuMapTypeInfraredFunktop.Checked = False
MnuMapTypeInfraredRainbow.Checked = False
End Sub
Private Sub MnuMapTypeInfraredAVN_Click()
UncheckAllMapTypes
MnuMapTypeInfraredAVN.Checked = True
MapType = "avn"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeInfraredDvorak_Click()
UncheckAllMapTypes
MnuMapTypeInfraredDvorak.Checked = True
MapType = "bd"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeInfraredFunktop_Click()
UncheckAllMapTypes
MnuMapTypeInfraredFunktop.Checked = True
MapType = "ft"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeInfraredJSL_Click()
UncheckAllMapTypes
MnuMapTypeInfraredJSL.Checked = True
MapType = "jsl"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeInfraredNormal_Click()
UncheckAllMapTypes
MnuMapTypeInfraredNormal.Checked = True
MapType = "ir4"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeInfraredRainbow_Click()
UncheckAllMapTypes
MnuMapTypeInfraredRainbow.Checked = True
MapType = "rb"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeInfraredRGB_Click()
UncheckAllMapTypes
MnuMapTypeInfraredRGB.Checked = True
MapType = "rgb"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeVisible_Click()
UncheckAllMapTypes
MnuMapTypeVisible.Checked = True
MapType = "vis"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeShortwave_Click()
UncheckAllMapTypes
MnuMapTypeShortwave.Checked = True
MapType = "ir2"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuMapTypeWaterVapor_Click()
UncheckAllMapTypes
MnuMapTypeWaterVapor.Checked = True
MapType = "wv"
Timer1.Enabled = False
GetPicList
End Sub

Private Sub UncheckAllMapLocs()
SysTypeFiles = False
MnuMapType.Enabled = True

MnuUSVisible.Checked = False
MnuUSWaterVapor.Checked = False
MnuUSInfrared.Checked = False
MnuUSInfraredCh2.Checked = False

MnuWestCoast.Checked = False
MnuWestCentralPacific.Checked = False
MnuCentralPacific.Checked = False
MnuEastPacific.Checked = False
MnuWestPacific.Checked = False
MnuNorthwestPacific.Checked = False
MnuNortheastPacific.Checked = False
MnuWestAtlantic.Checked = False
MnuCentralAtlantic.Checked = False
MnuEastAtlanticMET8.Checked = False
MnuGulfofMexico.Checked = False
MnuCaribbean.Checked = False
MnuEastCoast.Checked = False
MnuNorthwestAtlantic.Checked = False
MnuNorthAtlantic.Checked = False
MnuNortheastAtlanticMET8.Checked = False
End Sub
Private Sub MnuShowLatest_Click()
If MnuShowLatest.Checked = True Then
MnuShowLatest.Checked = False
ShowLatest = False
Timer2.Enabled = False
Else
MnuShowLatest.Checked = True
ShowLatest = True
FreezeTime = 0
Timer2.Enabled = True
End If
End Sub
Private Sub MnuUSInfrared_Click()
If Inet1.StillExecuting = True Then
MsgBox "The current downloads must compleate before choosing this map !"
Exit Sub
End If
Dim X As Long
UncheckAllMapLocs
MnuMapType.Enabled = False
MnuUSInfrared.Checked = True
MapRootLoc = ""
Timer1.Enabled = False
List1.Clear
For X = 5 To 20
List1.AddItem "http://www.ssd.noaa.gov/PS/PCPN/DATA/RT/NA/IR4/" & X & ".jpg"
Next X
SysTypeFiles = True
GetPicList
End Sub
Private Sub MnuUSInfraredCh2_Click()
If Inet1.StillExecuting = True Then
MsgBox "The current downloads must compleate before choosing this map !"
Exit Sub
End If
Dim X As Long
UncheckAllMapLocs
MnuMapType.Enabled = False
MnuUSInfraredCh2.Checked = True
MapRootLoc = ""
Timer1.Enabled = False
List1.Clear
For X = 5 To 20
List1.AddItem "http://www.ssd.noaa.gov/PS/PCPN/DATA/RT/NA/IR2/" & X & ".jpg"
Next X
SysTypeFiles = True
GetPicList
End Sub
Private Sub MnuUSVisible_Click()
If Inet1.StillExecuting = True Then
MsgBox "The current downloads must compleate before choosing this map !"
Exit Sub
End If
Dim X As Long
UncheckAllMapLocs
MnuMapType.Enabled = False
MnuUSVisible.Checked = True
MapRootLoc = ""
Timer1.Enabled = False
List1.Clear
For X = 5 To 20
List1.AddItem "http://www.ssd.noaa.gov/PS/PCPN/DATA/RT/NA/VIS/" & X & ".jpg"
Next X
SysTypeFiles = True
GetPicList
End Sub
Private Sub MnuUSWaterVapor_Click()
If Inet1.StillExecuting = True Then
MsgBox "The current downloads must compleate before choosing this map !"
Exit Sub
End If
Dim X As Long
UncheckAllMapLocs
MnuMapType.Enabled = False
MnuUSWaterVapor.Checked = True
MapRootLoc = ""
Timer1.Enabled = False
List1.Clear
For X = 5 To 20
List1.AddItem "http://www.ssd.noaa.gov/PS/PCPN/DATA/RT/NA/WV/" & X & ".jpg"
Next X
SysTypeFiles = True
GetPicList
End Sub
Private Sub MnuNortheastPacific_Click()
UncheckAllMapLocs
MnuNortheastPacific.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/west/nepac/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuNorthwestPacific_Click()
UncheckAllMapLocs
MnuNorthwestPacific.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/mtsat/nwpac/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuWestAtlantic_Click()
UncheckAllMapLocs
MnuWestAtlantic.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/east/watl/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuWestCentralPacific_Click()
UncheckAllMapLocs
MnuWestCentralPacific.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/mtsat/wcpac/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuWestCoast_Click()
UncheckAllMapLocs
MnuWestCoast.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/west/weus/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuCentralPacific_Click()
UncheckAllMapLocs
MnuCentralPacific.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/west/cpac/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuEastPacific_Click()
UncheckAllMapLocs
MnuEastPacific.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/west/epac/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuWestPacific_Click()
UncheckAllMapLocs
MnuWestPacific.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/mtsat/wpac/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuCentralAtlantic_Click()
UncheckAllMapLocs
MnuCentralAtlantic.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/east/catl/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuEastAtlanticMET8_Click()
UncheckAllMapLocs
MnuEastAtlanticMET8.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/met8/eatl/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuGulfofMexico_Click()
UncheckAllMapLocs
MnuGulfofMexico.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/east/gmex/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuCaribbean_Click()
UncheckAllMapLocs
MnuCaribbean.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/east/carb/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuEastCoast_Click()
UncheckAllMapLocs
MnuEastCoast.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/east/eaus/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuNorthwestAtlantic_Click()
UncheckAllMapLocs
MnuNorthwestAtlantic.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/east/nwatl/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuNorthAtlantic_Click()
UncheckAllMapLocs
MnuNorthAtlantic.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/goes/east/natl/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub MnuNortheastAtlanticMET8_Click()
UncheckAllMapLocs
MnuNortheastAtlanticMET8.Checked = True
MapRootLoc = "http://www.ssd.noaa.gov/met8/neatl/"
Timer1.Enabled = False
GetPicList
End Sub
Private Sub UncheckAllSpeeds()
MnuSpeed1X.Checked = False
MnuSpeed2X.Checked = False
MnuSpeed3X.Checked = False
MnuSpeed4X.Checked = False
End Sub
Private Sub MnuSpeed1X_Click()
UncheckAllSpeeds
MnuSpeed1X.Checked = True
Timer1.Interval = 400
End Sub
Private Sub MnuSpeed2X_Click()
UncheckAllSpeeds
MnuSpeed2X.Checked = True
Timer1.Interval = 300
End Sub
Private Sub MnuSpeed3X_Click()
UncheckAllSpeeds
MnuSpeed3X.Checked = True
Timer1.Interval = 200
End Sub
Private Sub MnuSpeed4X_Click()
UncheckAllSpeeds
MnuSpeed4X.Checked = True
Timer1.Interval = 100
End Sub
Private Sub Timer1_Timer()
If Shutdown = True Then Exit Sub
'Automatic reload
If Minute(Now) = "46" And Second(Now) = "59" Then
GoTo Reload
ElseIf Minute(Now) = "2" And Second(Now) = "59" Then
GoTo Reload
ElseIf Minute(Now) = "16" And Second(Now) = "59" Then
GoTo Reload
ElseIf Minute(Now) = "32" And Second(Now) = "59" Then
GoTo Reload
End If
'Display pic's
If ShowLatest = True Then PicY = List1.ListCount - 1
CurPicFile = App.Path & "\" & GetFileName(List1.List(PicY))
LoadPic
PicY = PicY + 1
If PicY >= List1.ListCount - 1 Then PicY = 0 'Seems to be some contreversy. OG: If PicY = List1.ListCount - 1 Then PicY = 0
Exit Sub
Reload:
Timer1.Enabled = False
GetPicList
End Sub
Private Sub GetPicList()

Call CleanUp
If SysTypeFiles = True Then
GatherPics
Exit Sub
End If

CurPicFile = ""
List1.Clear

Me.Caption = "Downloading picture list !"
TransferSuccess = GetNWSFileList(Inet1, MapRootLoc & "txtfiles/" & MapType & "_names.txt", MapRootLoc)
If Shutdown = True Then Exit Sub
If TransferSuccess = True Then
GatherPics
Else
Timer3.Enabled = True
End If
End Sub
Private Sub GatherPics()
Me.Caption = "Downloading Picture " & DLCount + 1 & " of " & List1.ListCount & " !"
TransferSuccess = GetInternetFile(Inet1, List1.List(DLCount), App.Path)
If Shutdown = True Then Exit Sub
If TransferSuccess = True Then
CurPicFile = App.Path & "\" & GetFileName(List1.List(DLCount))
LoadPic
DLCount = DLCount + 1
If DLCount = List1.ListCount Then
DLCount = 0
Timer1.Enabled = True
Me.Caption = "NWS Sat. Image Loops 1.0 SE"
Else
GatherPics
End If
Else
DLCount = 0
Timer3.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
'Automatic freeze pic animation resume
If FreezeTime = 30 Then MnuShowLatest_Click
FreezeTime = FreezeTime + 1
End Sub

Private Sub Timer3_Timer()
'Failed download recovery
If Shutdown = True Then Exit Sub
RecoveryTime = RecoveryTime + 1
If RecoveryTime = 60 Then
RecoveryTime = 0
Timer3.Enabled = False
GetPicList
End If
End Sub
Private Sub Timer4_Timer()
'Launch delay timer
If Shutdown = True Then Exit Sub
Timer4.Enabled = False
GetPicList
End Sub
Function ConvertTwipsToPixels(lngTwips As Long, lngDirection As Long) As Long
    
    'Handle to device
    Dim lngDC As Long
    Dim lngPixelsPerInch As Long
    Const nTwipsPerInch = 1440
    lngDC = GetDC(0)
    
    If (lngDirection = 0) Then 'Horizontal
        lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSX)
    Else 'Vertical
        lngPixelsPerInch = GetDeviceCaps(lngDC, WU_LOGPIXELSY)
    End If
    lngDC = ReleaseDC(0, lngDC)
    ConvertTwipsToPixels = (lngTwips / nTwipsPerInch) * lngPixelsPerInch
    
End Function
