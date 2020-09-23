VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmRyCamV2 
   Caption         =   "RyCam"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9330
   Icon            =   "frmRyCamV2.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   275
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   622
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMaskImg 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   6360
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox fileMask 
      Height          =   480
      Left            =   4080
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSettings 
      Caption         =   "Settings"
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton cmdRunCam 
      Caption         =   "Start Camera"
      Height          =   375
      Left            =   2640
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txtSendText 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5520
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   3000
      Width           =   3615
   End
   Begin SHDocVwCtl.WebBrowser WB 
      Height          =   2655
      Left            =   5520
      TabIndex        =   14
      Top             =   120
      Width           =   3615
      ExtentX         =   6376
      ExtentY         =   4683
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox txtApplyMask 
      Height          =   285
      Left            =   120
      TabIndex        =   13
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox txtMaskedText 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.PictureBox picTextMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      ScaleHeight     =   255
      ScaleWidth      =   3855
      TabIndex        =   11
      Top             =   3840
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox picOUT 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1920
      Left            =   120
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   10
      Top             =   1200
      Width           =   2520
   End
   Begin VB.PictureBox picFuzz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1920
      Index           =   3
      Left            =   -480
      Picture         =   "frmRyCamV2.frx":038A
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   9
      Top             =   3840
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.PictureBox picFuzz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1920
      Index           =   2
      Left            =   -480
      Picture         =   "frmRyCamV2.frx":282E
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.PictureBox picFuzz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1920
      Index           =   1
      Left            =   -480
      Picture         =   "frmRyCamV2.frx":4CA9
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   7
      Top             =   3840
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.PictureBox picIN 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1920
      Left            =   2880
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   6
      Top             =   1200
      Width           =   2520
   End
   Begin VB.PictureBox picFuzz 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   1920
      Index           =   0
      Left            =   -480
      Picture         =   "frmRyCamV2.frx":7171
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   2520
   End
   Begin VB.Timer tmrMain 
      Interval        =   500
      Left            =   3960
      Top             =   240
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtIP 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "IP Address"
      Top             =   360
      Width           =   1695
   End
   Begin MSWinsockLib.Winsock WSd 
      Left            =   3480
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock WS 
      Left            =   3000
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "READ THE READMEFIRST.TXT FILE"
      Height          =   375
      Left            =   2520
      TabIndex        =   21
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label lblConnTime 
      Height          =   255
      Left            =   2760
      TabIndex        =   18
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblSent 
      Alignment       =   2  'Center
      Caption         =   "Last Sent 00:00:00 PM"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblRec 
      Alignment       =   2  'Center
      Caption         =   "Last Received: 00:00:00 PM"
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Status: Not Connected"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2475
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuSetText 
         Caption         =   "Text Settings"
      End
      Begin VB.Menu mnuSetHandle 
         Caption         =   "Change my Handle(name)"
      End
      Begin VB.Menu mnuStats 
         Caption         =   "View Statistics"
      End
   End
   Begin VB.Menu mnuIP 
      Caption         =   "IP Finder"
      Visible         =   0   'False
      Begin VB.Menu mnuFindIP 
         Caption         =   "What's my IP?"
      End
   End
End
Attribute VB_Name = "frmRyCamV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'READ THE READMEFIRST.TXT FILE




'miscellaneous calls used in the program
Private Declare Function BmpToJpeg Lib "Bmp2Jpeg.dll" (ByVal BmpFilename As String, ByVal JpegFilename As String, ByVal CompressQuality As Integer) As Integer
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByVal nXDest As Long, ByVal nYDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'video capture calls
Private Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Private Declare Function SendMessage Lib "USER32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "USER32" () As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

'video capture constants
Private Const WM_CAP_DRIVER_CONNECT As Long = 1034
Private Const WM_CAP_DRIVER_DISCONNECT As Long = 1035
Private Const WM_CAP_GRAB_FRAME As Long = 1084
Private Const WM_CAP_EDIT_COPY As Long = 1054
Private Const WM_CAP_DLG_VIDEOFORMAT As Long = 1065
Private Const WM_CAP_DLG_VIDEOSOURCE As Long = 1066
Private Const WM_CLOSE = &H10

Dim CurrentMask As String 'the current picture mask being used
Dim SCROLLtext As String
Dim CooRD As POINTAPI
Dim SPwidth As Integer 'width of 1 space for the textmask
Dim SENDok As Boolean 'checks to see if program is ok to send another picture
Dim FUZZnum As Integer 'used when no input is coming in
Dim DATAin() As Byte 'pic data received
Dim DATAout() As Byte 'camera out data
Dim WSstates(9) As String 'camera status constants
Dim CamRunning As Boolean 'determines is camera is on or off
Dim MyName As String 'users name
Dim TimeConnected As Long 'tickcount value at time of connection
Dim BADpacketSENT() As Variant 'stores times of bad data sent
Dim BADpacketREC() As Variant 'stores times of bad data recieved
Dim numSENT As Long 'total pics sent
Dim numREC As Long 'total pics received

Private mCapHwnd As Long



Private Sub cmdConnect_Click()
    If cmdConnect.Caption = "Cancel" Then 'close the connection
        WS.SendData "9999"
        WS.Close
        WS.Listen
        WSd.Close
        WSd.Listen
        cmdConnect.Caption = "Connect"
        Exit Sub
    End If
    'reset winsocks and open the connection
    cmdConnect.Caption = "Cancel"
    WS.Close
    WSd.Close
    WS.RemoteHost = txtIP
    WSd.RemoteHost = txtIP
    WS.RemotePort = 7634
    WSd.RemotePort = 7635
    WS.Connect
    WSd.Connect
    Do Until WS.State = sckConnected
        DoEvents: DoEvents: DoEvents: DoEvents
        If WS.State = sckError Then
            MsgBox "Problem connecting! 1"
            GoTo RESconnection
            Exit Sub
        End If
    Loop
    Do Until WSd.State = sckConnected
        DoEvents: DoEvents: DoEvents: DoEvents
        If WSd.State = sckError Then
            MsgBox "Problem connecting! 2"
            GoTo RESconnection
            Exit Sub
        End If
    Loop
    SENDok = True
    WS.SendData "0000"
    TimeConnected = GetTickCount 'used to determine how long connected with someone
    
    Exit Sub
RESconnection: 'if an error occured, then reset the connection
        WS.Close
        WS.Listen
        WSd.Close
        WSd.Listen
        cmdConnect.Caption = "Connect"
End Sub
Private Sub cmdConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then frmRyCamV2.PopupMenu mnuIP
End Sub
Private Sub cmdRunCam_Click()
    If CamRunning = True Then 'stop the camera
        CamRunning = False
        cmdRunCam.Caption = "Start Camera"
        SENDok = False
        SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
        picOUT.Picture = picFuzz(0).Picture
        picOUT.Refresh
    Else 'start up the camera
        cmdRunCam.Caption = "Stop Camera"
        mCapHwnd = capCreateCaptureWindow("My Own Capture Window", 0, 0, 0, 320, 240, Me.hwnd, 0)
        SendMessage mCapHwnd, WM_CAP_DRIVER_CONNECT, 0, 0
        DoEvents
        SENDok = True
        CamRunning = True
    End If
End Sub
Private Sub cmdSettings_Click()
    frmRyCamV2.PopupMenu mnuSettings
End Sub
Private Sub Form_Load()
    WS.LocalPort = 7634 'initialize winsocks
    WSd.LocalPort = 7635
    WS.Listen
    WSd.Listen
    GetTextExtentPoint32 picTextMask.hdc, " ", 1, CooRD
    SPwidth = CooRD.X 'for use with some of the masks
    WSstates(0) = "Not Connected": WSstates(1) = "Open": WSstates(2) = "Listening": WSstates(3) = "Connection Pending": WSstates(4) = "Resolving Host": WSstates(5) = "Host Resolved": WSstates(6) = "Connecting": WSstates(7) = "Connected": WSstates(8) = "Closing": WSstates(9) = "Error"
    picOUT.Picture = picFuzz(0).Picture
    picOUT.Refresh
    picIN.Picture = picFuzz(0).Picture
    CurrentMask = "none"
    Open App.Path & "\tmpconvo.htm" For Output As #1 'create the chat window html file
        Dim tmPP As String
        tmPP = "<head><script language='javascript'>var scrollMe = window.setInterval('window.scrollBy(0,1000);', 1000);</script></head>"
        Print #1, "<html>" & tmPP & "<body bgcolor=white><center><font color=lime>RyCam v" & App.Major & "." & App.Minor & "." & App.Revision & " Chat Window</font></center>"
    Close #1
    If Dir$(App.Path & "/settings.txt") <> "" Then 'load in the users settings
        Open App.Path & "/settings.txt" For Input As #1
            Input #1, MyName
            Dim TT
            Input #1, TT
            txtSendText.Font = Trim(TT)
            Input #1, TT
            txtSendText.FontBold = Trim(TT)
            Input #1, TT
            txtSendText.ForeColor = Trim(TT)
        Close #1
    End If
    WB.Navigate App.Path & "\tmpconvo.htm" 'navigate to the chat window file
    Do Until MyName <> "" 'if user has not set a name, IE, first use, then set up a name
        MyName = InputBox("Enter your name:")
    Loop
    MakeSettings 'save the settings file
    frmRyCamV2.Caption = "RyCam v" & App.Major
    ReDim BADpacketSENT(0)
    ReDim BADpacketREC(0)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    MakeSettings
    WS.Close
    WSd.Close
    tmrMain.Interval = 0
    If CamRunning = True Then SendMessage mCapHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0
    Erase WSstates
    Erase DATAout
    Erase DATAin
End Sub
Private Sub lblRec_DblClick()
    txtApplyMask.Visible = True
End Sub
Private Sub mnuFindIP_Click() 'link to my ip address website
    ShellExecute 0&, vbNullString, "http://lutzlutz.tripod.com/iper.htm", vbNullString, vbNullString, vbNormalFocus
End Sub
Private Sub mnuSetHandle_Click()
    MyName = ""
    Do Until MyName <> ""
        MyName = InputBox("Enter your name?")
    Loop
    MakeSettings
End Sub
Private Sub mnuSetText_Click()
    frmRyCamv2a.Show
End Sub
Private Sub mnuStats_Click()
    Call ShowBadData
End Sub
Private Sub tmrMain_Timer()
    On Error Resume Next
    DoEvents
    If CamRunning = True Then 'snap the picture
        SendMessage mCapHwnd, WM_CAP_GRAB_FRAME, 0, 0
        SendMessage mCapHwnd, WM_CAP_EDIT_COPY, 0, 0
        picOUT.Picture = Clipboard.GetData
        DoMasks CurrentMask
    End If
    lblStatus = "Connection Status: " & WSstates(WS.State) 'display status of connection
    If WS.State = 7 Then lblStatus.ForeColor = vbGreen
    If WS.State = 2 Then lblStatus.ForeColor = vbBlue
    If WS.State = 0 Or WS.State = 1 Or WS.State = 8 Or WS.State = 9 Then 'reset the connection
        lblStatus.ForeColor = vbRed
        WS.Close
        WS.Listen
        WSd.Close
        WSd.Listen
        cmdConnect.Caption = "Connect"
    End If
    If WS.State <> 7 Then 'no incoming video, so make static
        FUZZnum = FUZZnum + 1
        If FUZZnum = 4 Then FUZZnum = 0
        BitBlt picIN.hdc, 0, 0, picIN.Width, picIN.Height, picFuzz(FUZZnum).hdc, 0, 0, vbSrcCopy
        picIN.Refresh
        lblConnTime = ""
    Else
        Dim DaCount As Long
        DaCount = (GetTickCount - TimeConnected) \ 1000 'display how long connected
        lblConnTime = Format((DaCount \ 3600000) - 24 * (DaCount \ 86400000), "00") & ":" & Format((DaCount \ 60) - (60 * (DaCount \ 3600)), "00") & ":" & Format(DaCount Mod 60, "00")
        If CamRunning = False Then Exit Sub
        If SENDok = False Then Exit Sub
        SENDok = False
        SavePicture picOUT.Image, App.Path & "\tmpout.bmp"
        BmpToJpeg App.Path & "\tmpout.bmp", App.Path & "\tmpout.jpg", 40
        Dim FREEout As Integer
        FREEout = FreeFile 'load the image into array to be sent
        Open App.Path & "\tmpout.jpg" For Binary As #FREEout
            ReDim DATAout(1 To LOF(FREEout))
            Get #FREEout, , DATAout
        Close #FREEout
        WSd.SendData DATAout 'send the image array
        numSENT = numSENT + 1
        lblSent.Caption = "Last Sent: " & Time
    End If
End Sub
Private Sub txtApplyMask_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 13 Then
        If txtApplyMask = "dotext" Then 'display text
            CurrentMask = "text"
            txtMaskedText.Visible = True
        End If
        If txtApplyMask = "doscroll" Then 'display scrolling text
            GetTextExtentPoint32 picTextMask.hdc, txtMaskedText, Len(txtMaskedText), CooRD
            SCROLLtext = txtMaskedText & Space((picOUT.Width - CooRD.X) \ SPwidth)
            CurrentMask = "scroll"
            txtMaskedText.Visible = True
        End If
        If txtApplyMask = "dotime" Then 'display time
            CurrentMask = "time"
            txtMaskedText.Visible = False
        End If
        If txtApplyMask = "doreset" Then 'reset all the things
            CurrentMask = "none"
            txtMaskedText.Visible = False
        End If
        If Left(txtApplyMask, 7) = "docolor" And Len(txtApplyMask) = 16 Then 'change the text color
            picTextMask.ForeColor = RGB(Val(Mid(txtApplyMask, 8, 3)), Val(Mid(txtApplyMask, 11, 3)), Val(Mid(txtApplyMask, 14, 3)))
            txtMaskedText.Visible = False
        End If
        If txtApplyMask = "dolist" Then MsgBox "Commands available in this program:" & vbCrLf & vbTab & "dotext  -  displays a text message" & vbCrLf & vbTab & "doscroll  -  scrolls the text" & vbCrLf & vbTab & "dotime  -  displays the time in the image" & vbCrLf & vbTab & "doreset  -  resets the masks", vbOKOnly, "Command Menu"
        If Left(txtApplyMask, 8) = "docustom" Then
            fileMask.Path = App.Path & "\masks"
            fileMask.Refresh
            For z = 0 To fileMask.ListCount - 1
                If Mid(txtApplyMask, 10, Len(txtApplyMask) - 9) = fileMask.List(z) Then
                    picMaskImg.Picture = LoadPicture(App.Path & "\masks\" & fileMask.List(z))
                    CurrentMask = "image"
                End If
            Next z
        End If
        
        txtApplyMask.Visible = False
        txtApplyMask = ""
    End If
End Sub
Private Sub txtIP_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If txtIP = "IP Address" Then txtIP = ""
End Sub
Private Sub txtMaskedText_Change()
    On Error Resume Next
    picTextMask.Cls
    GetTextExtentPoint32 picTextMask.hdc, txtMaskedText, Len(txtMaskedText), CooRD
    If CurrentMask = "scroll" Then 'set up in scroll mode
        SCROLLtext = txtMaskedText & Space((picOUT.Width - CooRD.X) \ SPwidth)
    Else 'normal text mode
        picTextMask.Print Space((picOUT.Width - CooRD.X) \ (2 * SPwidth)) & txtMaskedText
    End If
    picTextMask.Refresh
End Sub
Private Sub txtSendText_KeyDown(KeyCode As Integer, Shift As Integer) 'sending messages
    Dim DoBold As String, DoBold2 As String, ConvoText As String
    If KeyCode = 13 Then
        If txtSendText = "" Then Exit Sub
        If txtSendText.FontBold = True Then
            DoBold = "<b>"
            DoBold2 = "</b>"
        Else
            DoBold = ""
            DoBold2 = ""
        End If
        ConvoText = "<font face='" & txtSendText.Font & "' color='#" & GetHex(txtSendText.ForeColor) & "'>" & DoBold & MyName & ":  " & txtSendText & DoBold2 & "</font><br>"
        If WS.State = 7 Then WS.SendData "5555" & ConvoText
        Dim FF As Integer
        FF = FreeFile
        Open App.Path & "\tmpconvo.htm" For Append As #FF
            Print #FF, ConvoText
        Close #FF
        WB.Refresh2
        txtSendText = ""
    End If
End Sub
Private Sub txtSendText_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then txtSendText = "" 'clear out the text
End Sub
Private Sub WS_ConnectionRequest(ByVal requestID As Long)
    WS.Close
    WS.Accept requestID
End Sub
Private Sub WS_DataArrival(ByVal bytesTotal As Long)
    Dim tmpINcmd As String
    WS.GetData tmpINcmd
    If tmpINcmd = "0000" Then SENDok = True 'command to send a pic
    If Left(tmpINcmd, 4) = "5555" Then 'command for incoming text
        Dim FF As Integer, ConvoText As String
        FF = FreeFile
        ConvoText = Right(tmpINcmd, Len(tmpINcmd) - 4)
        Open App.Path & "\tmpconvo.htm" For Append As #FF
            Print #FF, ConvoText
        Close #FF
        WB.Refresh
    End If
    If InStr(tmpINcmd, "<!--6666-->") > 0 Then 'command for error statistics
        ReDim Preserve BADpacketSENT(0 To UBound(BADpacketSENT) + 1)
        BADpacketSENT(UBound(BADpacketSENT)) = Time
        SENDok = True
    End If
End Sub
Private Sub WSd_ConnectionRequest(ByVal requestID As Long)
    WSd.Close
    WSd.Accept requestID
End Sub
Private Sub WSd_DataArrival(ByVal bytesTotal As Long) 'incoming picture data
    Erase DATAin
    ReDim DATAin(1 To 50000)
    On Error GoTo eRRoRiN
    WSd.GetData DATAin
    Dim FREEin As Integer
    If Dir$(App.Path & "\tmpin.jpg") <> "" Then Kill App.Path & "\tmpin.jpg" 'delete old file if it exists
    FREEin = FreeFile 'open file to save incoming pic data to
    Open App.Path & "\tmpin.jpg" For Binary As #FREEin
        Put #FREEin, , DATAin
    Close #FREEin
    numREC = numREC + 1
    lblRec.Caption = "Last Received: " & Time
    picIN.Picture = LoadPicture(App.Path & "\tmpin.jpg")
    WS.SendData "0000"
    Exit Sub
eRRoRiN: 'if error occured, set error statistics
    ReDim Preserve BADpacketREC(0 To UBound(BADpacketREC) + 1)
    BADpacketREC(UBound(BADpacketREC)) = Time
    WS.SendData "<!--6666-->"
End Sub
Private Sub DoMasks(MaskType As String) 'masking the image
    Select Case MaskType
        Case "text" 'adds text to the top of image
            BitBlt picOUT.hdc, 0, 0, picOUT.Width, picTextMask.Height, picTextMask.hdc, 0, 0, vbSrcCopy
        Case "scroll" 'scrolls text at top of image
            SCROLLtext = Right(SCROLLtext, Len(SCROLLtext) - 2) & Left(SCROLLtext, 2)
            picTextMask.Cls
            picTextMask.Print SCROLLtext
            picTextMask.Refresh
            BitBlt picOUT.hdc, 0, 0, picOUT.Width, picTextMask.Height, picTextMask.hdc, 0, 0, vbSrcCopy
        Case "time" 'adds time to top of image
            txtMaskedText = Time
            BitBlt picOUT.hdc, 0, 0, picOUT.Width, picTextMask.Height, picTextMask.hdc, 0, 0, vbSrcCopy
        Case "image"
            BitBlt picOUT.hdc, 0, 0, picOUT.Width, picOUT.Height, picMaskImg.hdc, 0, 0, vbSrcAnd
        Case "none"
            'do nothing
    End Select
    picOUT.Refresh
End Sub
Private Function GetHex(theColor As Long) 'converts a long color to a hex color
    Dim HexMake(2)
    If theColor < 0 Then theColor = 0
    HexMake(2) = theColor \ 65536
    theColor = theColor - HexMake(2) * 65536
    HexMake(1) = theColor \ 256
    HexMake(0) = theColor - HexMake(1) * 256
    For z = 0 To 2
        HexMake(z) = Hex(RGB(HexMake(z), 0, 0))
        If Len(HexMake(z)) = 1 Then
            HexMake(z) = "0" & HexMake(z)
        Else
            HexMake(z) = HexMake(z)
        End If
    Next z
    GetHex = HexMake(0) & HexMake(1) & HexMake(2)
    Erase HexMake
End Function
Public Sub MakeSettings() 'saves user settings
    Dim FreeF As Integer
    FreeF = FreeFile
    Open App.Path & "/settings.txt" For Output As #FreeF
        Print #FreeF, MyName
        Print #FreeF, txtSendText.Font
        Print #FreeF, txtSendText.FontBold
        Print #FreeF, txtSendText.ForeColor
    Close #FreeF
End Sub
Private Sub ShowBadData() 'shows error statistics
    frmRyCamV2b.Show
    DoEvents
    With frmRyCamV2b
        .lstA.Clear
        .lstS.Clear
        .lstR.Clear
        .lstA.AddItem "Total Sent:      " & numSENT
        .lstA.AddItem "Total Received:  " & numREC
        .lstA.AddItem "Bad Sent:        " & UBound(BADpacketSENT)
        .lstA.AddItem "Bad Received:    " & UBound(BADpacketREC)
        For z = 0 To UBound(BADpacketSENT)
            .lstS.AddItem BADpacketSENT(z)
        Next z
        For z = 0 To UBound(BADpacketREC)
            .lstR.AddItem BADpacketREC(z)
        Next z
    End With
End Sub
