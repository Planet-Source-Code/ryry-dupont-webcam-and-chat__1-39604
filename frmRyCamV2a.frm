VERSION 5.00
Begin VB.Form frmRyCamv2a 
   BorderStyle     =   0  'None
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   16
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   2880
      TabIndex        =   15
      Top             =   3600
      Width           =   735
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.PictureBox picD 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      ScaleHeight     =   435
      ScaleWidth      =   1635
      TabIndex        =   13
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cboFont 
      Height          =   315
      Left            =   3600
      Sorted          =   -1  'True
      TabIndex        =   12
      Text            =   "Font"
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Timer tmrMoveForm 
      Left            =   0
      Top             =   720
   End
   Begin VB.PictureBox picSel 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   3120
      ScaleHeight     =   209
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   6
      Top             =   240
      Width           =   255
      Begin VB.Image imgSel 
         Height          =   75
         Left            =   0
         Picture         =   "frmRyCamV2a.frx":0000
         Top             =   1485
         Width           =   75
      End
   End
   Begin VB.Timer tmrMove 
      Left            =   0
      Top             =   240
   End
   Begin VB.PictureBox picC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   2880
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   5
      Top             =   360
      Width           =   135
   End
   Begin VB.PictureBox picB 
      Height          =   495
      Left            =   4800
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   29
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.PictureBox picA 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   120
      Picture         =   "frmRyCamV2a.frx":003D
      ScaleHeight     =   187
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   175
      TabIndex        =   3
      Top             =   360
      Width           =   2625
   End
   Begin VB.TextBox TA 
      Height          =   285
      Index           =   2
      Left            =   4080
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TA 
      Height          =   285
      Index           =   1
      Left            =   4080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox TA 
      Height          =   285
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   1200
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      Height          =   195
      Left            =   3600
      TabIndex        =   10
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      Height          =   195
      Left            =   3600
      TabIndex        =   9
      Top             =   480
      Width           =   300
   End
   Begin VB.Label lblQuit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " X "
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   5280
      TabIndex        =   7
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblTopBar 
      BackColor       =   &H80000002&
      BorderStyle     =   1  'Fixed Single
      Height          =   135
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   5655
   End
   Begin VB.Shape shpBord 
      BorderWidth     =   2
      Height          =   4215
      Left            =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmRyCamv2a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'READ THE READMEFIRST.TXT FILE



Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Private Declare Function GetCursorPos Lib "user32.dll" (lpPoint As POINT_TYPE) As Long
Private Type POINT_TYPE
    X As Long
    Y As Long
End Type

Dim COLORsel As Boolean 'to see if mouse is held down on basecolor selector
Dim pREF(2) As Long 'temp var for longtorgb
Dim REF(2) As Long 'base color
Dim CurrentPCT 'current spot selector is at from top of colorbar
Dim LPx As Integer, LPy As Integer 'vars for moving the form

Private Sub LONGtoRGB(pColor As Long)
    pREF(2) = pColor \ 65536
    pColor = pColor - pREF(2) * 65536
    pREF(1) = pColor \ 256
    pREF(0) = pColor - pREF(1) * 256
End Sub
Private Sub UpdateText(lngColor As Long) 'update the preview text
    LONGtoRGB lngColor
    For z = 0 To 2
        TA(z) = pREF(z)
    Next z
    picD.ForeColor = picB.BackColor
    picD.Cls
    picD.FontItalic = False
    picD.Print "AaBbCcZz"
End Sub
Private Sub MakeSide()
    On Error Resume Next
    Dim pHei, pWid
    Dim PCT, PCTi(2)
    pHei = picC.Height
    pWid = picC.Width
    For z = 0 To (pHei - 1) / 2 'display top half of selector bar(black to basecolor)
        PCT = 2 * (z / (pHei - 1))
        For zz = 0 To pWid - 1
            SetPixelV picC.hdc, zz, z, RGB(Abs(REF(0) * PCT), Abs(REF(1) * PCT), Abs(REF(2) * PCT))
            'picC.PSet (zz, z), RGB(Abs(REF(0) * PCT), Abs(REF(1) * PCT), Abs(REF(2) * PCT))
        Next zz
    Next z
    
    PCTi(0) = (255 - REF(0))
    PCTi(1) = (255 - REF(1))
    PCTi(2) = (255 - REF(2))
    For z = (pHei - 1) / 2 To (pHei - 1) 'display bottom half of selector bar(basecolor to white)
        PCT = 2 * (z / (pHei - 1)) - 1
        For zz = 0 To pWid - 1
            SetPixelV picC.hdc, zz, z, RGB(REF(0) + PCT * PCTi(0), REF(1) + PCT * PCTi(1), REF(2) + PCT * PCTi(2))
            'picC.PSet (zz, z), RGB(REF(0) + PCT * PCTi(0), REF(1) + PCT * PCTi(1), REF(2) + PCT * PCTi(2))
        Next zz
    Next z
    picC.Refresh
    Erase PCTi
End Sub
Private Sub cboFont_Change()
    On Error Resume Next 'font selector
    picD.Font = cboFont.List(cboFont.ListIndex)
    UpdateText picB.BackColor
End Sub
Private Sub cboFont_Click()
    Call cboFont_Change
End Sub
Private Sub cboFont_Scroll()
    Call cboFont_Change
End Sub
Private Sub chkBold_Click() 'bold text selector
    If chkBold.Value = 0 Then
        picD.FontBold = False
    Else
        picD.FontBold = True
    End If
    UpdateText picB.BackColor
End Sub
Private Sub cmdCancel_Click()
    Unload frmRyCamv2a
End Sub
Private Sub cmdOk_Click()
    frmRyCamV2.txtSendText.Font = picD.Font
    frmRyCamV2.txtSendText.ForeColor = picD.ForeColor
    frmRyCamV2.txtSendText.FontBold = picD.FontBold
    frmRyCamV2.MakeSettings
    Unload frmRyCamv2a
End Sub
Private Sub Form_Load() 'set up default values
    Call picA_MouseDown(1, 1, 0, 0)
    COLORsel = False
    For z = 0 To Screen.FontCount - 1
        cboFont.AddItem Screen.Fonts(z)
    Next z
    cboFont.Text = frmRyCamV2.txtSendText.Font
    If frmRyCamV2.txtSendText.FontBold = True Then
        chkBold.Value = 1
    Else
        chkBold.Value = 0
    End If
    shpBord.Width = frmRyCamv2a.ScaleWidth + 1
    shpBord.Height = frmRyCamv2a.ScaleHeight + 1
    lblTopBar.Width = frmRyCamv2a.ScaleWidth
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Erase pREF
    Erase REF
End Sub
Private Sub imgSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrMove.Interval = 1 'start moving selector
End Sub
Private Sub imgSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrMove.Interval = 0 'stop moving selector
End Sub
Private Sub lblQuit_Click()
    Unload frmRyCamv2a
End Sub
Private Sub lblTopBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LPx = X 'start moving form
    LPy = Y
    tmrMoveForm.Interval = 1
End Sub
Private Sub lblTopBar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    tmrMoveForm.Interval = 0 'stop moving form
End Sub
Private Sub picA_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
     picA.Cls 'display a crosshair around selected point
     For z = X - 4 To X - 1
        picA.PSet (z, Y), vbBlack
     Next z
     For z = X + 1 To X + 4
        picA.PSet (z, Y), vbBlack
     Next z
     For z = Y - 4 To Y - 1
        picA.PSet (X, z), vbBlack
     Next z
     For z = Y + 1 To Y + 4
        picA.PSet (X, z), vbBlack
     Next z
     picA.Refresh
     LONGtoRGB picA.Point(X, Y)
     For z = 0 To 2
        REF(z) = pREF(z) 'store the selected base color
     Next z
     frmRyCamv2a.Caption = REF(0) & "  " & REF(1) & "  " & REF(2)
     If X >= 0 And X <= picA.Width - 1 And Y >= 0 And Y <= picA.Height - 1 Then MakeSide
     picB.BackColor = picC.Point(2, imgSel.Top - 6)
     UpdateText picB.BackColor
     COLORsel = True
End Sub
Private Sub picA_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If COLORsel = True Then Call picA_MouseDown(Button, Shift, X, Y)
End Sub
Private Sub picA_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    COLORsel = False
End Sub
Private Sub picSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Y < 6 Then Y = 6
    If Y > 192 Then Y = 192
    imgSel.Top = Y - 2
    picB.BackColor = picC.Point(2, imgSel.Top - 6)
    UpdateText picB.BackColor
End Sub
Private Sub tmrMove_Timer() 'move color selector on bar
    Dim CooRD As POINT_TYPE
    GetCursorPos CooRD
    CooRD.Y = CooRD.Y - frmRyCamv2a.Top / 15 - picSel.Top
    'frmRyCamv2a.Caption = CooRD.Y
    If CooRD.Y > 5 And CooRD.Y < 193 Then
        imgSel.Top = CooRD.Y
    End If
    picB.BackColor = picC.Point(2, imgSel.Top - 6)
    UpdateText picB.BackColor 'change the color selected
End Sub
Private Sub tmrMoveForm_Timer() 'move the form
    Dim MovePos As POINT_TYPE
    GetCursorPos MovePos
    frmRyCamv2a.Top = MovePos.Y * 15 - LPy
    frmRyCamv2a.Left = MovePos.X * 15 - LPx
End Sub
