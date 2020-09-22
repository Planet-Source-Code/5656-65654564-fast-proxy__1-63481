VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Fast Proxy"
   ClientHeight    =   5715
   ClientLeft      =   1875
   ClientTop       =   1620
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Impact"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":492A
   ScaleHeight     =   5715
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6000
      Left            =   0
      Picture         =   "Form1.frx":D2234
      ScaleHeight     =   6000
      ScaleMode       =   0  'User
      ScaleWidth      =   11070.08
      TabIndex        =   0
      Top             =   0
      Width           =   11040
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   4880
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   0
         Top             =   4490
      End
      Begin VB.PictureBox Pictureicon 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         Picture         =   "Form1.frx":19FB3E
         ScaleHeight     =   240
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   160
         Width           =   255
      End
      Begin VB.ListBox lstResponse 
         BackColor       =   &H00000000&
         ForeColor       =   &H00404040&
         Height          =   3885
         ItemData        =   "Form1.frx":1A4468
         Left            =   7800
         List            =   "Form1.frx":1A446A
         TabIndex        =   9
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00F6C944&
         Caption         =   "Clear"
         Height          =   375
         Left            =   4920
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   4680
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00F6C944&
         Caption         =   "Clear"
         Height          =   375
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4680
         Width           =   2775
      End
      Begin VB.ListBox lstIP 
         BackColor       =   &H00000000&
         ForeColor       =   &H00404040&
         Height          =   3885
         ItemData        =   "Form1.frx":1A446C
         Left            =   2040
         List            =   "Form1.frx":1A446E
         TabIndex        =   6
         Top             =   720
         Width           =   2775
      End
      Begin VB.ListBox lstWin 
         BackColor       =   &H00000000&
         ForeColor       =   &H00404040&
         Height          =   3885
         ItemData        =   "Form1.frx":1A4470
         Left            =   4920
         List            =   "Form1.frx":1A4472
         TabIndex        =   5
         Top             =   720
         Width           =   2775
      End
      Begin MSComctlLib.ProgressBar PB1 
         Height          =   255
         Left            =   2160
         TabIndex        =   11
         Top             =   5160
         Visible         =   0   'False
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSWinsockLib.Winsock Sck1 
         Left            =   480
         Top             =   4880
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Response Time"
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   7800
         TabIndex        =   24
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Proxies: 0"
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Working Proxies: 0"
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "List To Scan"
         ForeColor       =   &H00F6C944&
         Height          =   285
         Left            =   240
         TabIndex        =   21
         Top             =   930
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fast Proxy"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   8836
         TabIndex        =   20
         Top             =   5160
         Width           =   975
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   160
         Width           =   255
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   4080
         TabIndex        =   18
         Top             =   160
         Width           =   255
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   220
         TabIndex        =   17
         Top             =   2000
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scan"
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   1290
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scroll Working Proxies"
         Enabled         =   0   'False
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Stop All Scanning"
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   230
         TabIndex        =   14
         Top             =   1650
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
         ForeColor       =   &H00F6C944&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Label8"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   9840
         TabIndex        =   12
         Top             =   5160
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "Command1"
      Height          =   255
      Left            =   11400
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000000&
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F6C944&
      Height          =   515
      Left            =   13440
      TabIndex        =   2
      Top             =   960
      Width           =   1215
      Begin VB.TextBox txtPort 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00F6C944&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "80"
         Top             =   165
         Width           =   975
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00000000&
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F6C944&
      Height          =   515
      Left            =   14760
      TabIndex        =   1
      Top             =   960
      Width           =   4800
      Begin VB.Shape ShapeLead 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00F6C944&
         Height          =   255
         Left            =   0
         Shape           =   2  'Oval
         Top             =   240
         Width           =   4575
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim ss As Integer
Dim GoIP As String
Dim X1, X2, i As Integer
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long
Const WM_USER = &H400
Const CCM_FIRST = &H2000&
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETBARCOLOR = (WM_USER + 9)


'Api Functions Declaration
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Selected As Integer
Dim Focused As Boolean
Private Declare Function SendMessagePBChange Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Private Sub Form_ResizeSysTray()

    '********************************************************************
    'Add this to resize event to hide in tray on minimize

        If Me.WindowState = vbMinimized Then
            Call SystrayOn(Me, "Double Click to Restore Me back to the screen")
            Call ChangeSystrayToolTip(Me, "Double Click to Restore Me back to the screen")
            Call PopupBalloon(Me, "App is now hidden in the Systray !" + vbCrLf + "Double click Icon to restore", "Balloon Tool Tip")
            Me.Hide
        End If

End Sub

Private Sub Form_Terminate()
'Added to keep program from crashing
    On Error Resume Next
 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyEnable", "0", REG_DWORD
 '********************************************************************
    'If you don't remove icon from tray on double click show, add this
    'good idea
    Call SystrayOff(Me)
    Unload Me
    End
End Sub
Private Sub SetTrans()
'Free the memory set
If hRgn Then DeleteObject hRgn

'Scan the Bitmap and remove all transparent pixels from it, creating a new region
hRgn = GetBitmapRegion(Me.Picture, vbWhite)

'Set the Forms new Region
SetWindowRgn Me.hWnd, hRgn, True

End Sub

Public Function Change_pb_ForeColor(ByVal hWnd As Long, ByVal lColor As Long)
SendMessagePBChange hWnd, PBM_SETBARCOLOR, 0, ByVal lColor
End Function
Public Function Change_pb_Color(ByVal hWnd As Long, ByVal lColor As Long)
          SendMessagePBChange hWnd, PBM_SETBKCOLOR, 0, ByVal lColor
End Function
Public Function Txt_FileExists(ByVal WhatFile As String) As Boolean
        Txt_FileExists = (0 < Len(Trim$(Dir$(WhatFile))))
End Function
Public Sub Pause(ByVal seconds As Single)
Call Sleep(Int(seconds * Form2.Text1.Text))
End Sub

Private Sub cmdFile_Click()
    'Added to keep program from crashing
    On Error Resume Next
    Form2.CD1.Filter = "All Files (*.*) | *.*"
    Form2.CD1.ShowOpen
    Form2.TxtFile.Text = Form2.CD1.FileName
    X1 = -1
    X2 = -1
End Sub
Private Sub cmdFile2_Click()
'Added to keep program from crashing
    On Error Resume Next
    Form2.CD1.Filter = "All Files (*.*) | *.*"
    Form2.CD1.ShowOpen
    Form2.TxtFile.Text = Form2.CD1.FileName
    X1 = -1
    X2 = -1
End Sub
Private Sub Command3_Click()
'Added to keep program from crashing
    On Error Resume Next
lstIP.Clear
End Sub
Private Sub Command6_Click()
'Added to keep program from crashing
    On Error Resume Next
lstWin.Clear
lstResponse.Clear
End Sub

Private Sub Label1_Click()
'Added to keep program from crashing
   
cmdFile2_Click
End Sub
Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim ReturnVal%
    ReleaseCapture
    ReturnVal% = SendMessage(hWnd, &HA1, 2, 0)
End Sub



'See's If File Exists
Private Sub Form_Load()
 '********************************************************************
    'If you want the form to be in the tray on startup add this

    'Call SystrayOn(Me, "Fast Proxy Ver." & App.Major & "." & App.Minor & "." & App.Revision & " Current Proxy " & Form2.Text2)
    Call SystrayOn(Me, "Fast Proxy " & Form2.Text2)
 'Call PopupBalloon(Me, "Put your Message Here !", "Balloon Tool Tip")

'Center form - Cant use the forms properties, because when restored
' from the tray, it will go back to the center!
    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2

'Creates the Form
FakeForm Me, Picture1, RGB(255, 255, 255)
'your bmp's outline color must be 255,255,255, which is White
'make outline transparent
SetTrans

'gets your version from version.dat
Open "version.dat" For Input As #1
        If Not EOF(1) Then
            Line Input #1, CurrentVersion
        
Label8.Caption = "Ver." & CurrentVersion
End If
        Close #1

Dim file_path As String
Dim file_name As String

'reads saved options
Dim MyText As Integer
Form2.Text1.Text = GetSetting(App.EXEName, "MyString1", "MyText1", "")
Form2.TxtFile.Text = GetSetting(App.EXEName, "MyString2", "MyText2", "")

    file_path = App.Path
    If Right$(file_path, 1) <> "\" Then file_path = file_path & "\"

    file_name = file_path & "Working Proxies.txt"
    If Txt_FileExists(file_name) Then
       If Form2.Check1.Value = 1 Then
    Form1.Hide
   
End If
        
        lstIP.Refresh
        
    Else
        MsgBox "Working Proxies.txt Not Found,  Please Create It! Then Try Me Again."
        Unload Me
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Added to keep program from crashing
    On Error Resume Next
 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyEnable", "0", REG_DWORD
 '********************************************************************
    'If you don't remove icon from tray on double click show, add this
    'good idea
    Call SystrayOff(Me)
    Unload Me
End
End Sub
Private Sub Label10_Click()
'Added to keep program from crashing
    On Error Resume Next
'msgbox warning
Dim vbYesNoButton
vbYesNoButton = MsgBox("All Fast Proxy Features will be disabled, Are you sure you want to do this?", vbYesNo, "Warning!!!")
If vbYesNoButton = 6 Then
 SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyEnable", "0", REG_DWORD
    Unload Me
ElseIf vbYesNoButton = 7 Then
    Form1.Show
End If


End Sub
Private Sub Label11_Click()
'Added to keep program from crashing
    On Error Resume Next
Me.Hide
End Sub





Private Sub Label4_Click()
MsgBox "This Program Will Help Hide Your IP While Your On The Internet." & vbCrLf _
& "I Will Be Adding More Options and Settings." & vbCrLf _
& "Feedback is Welcome Email troyh@frontiernet.net", vbOKOnly, "About Fast Proxy"
End Sub

Private Sub Label5_Click()
'Added to keep program from crashing
    

lstWin.Clear
lstResponse.Clear
Timer1.Enabled = True
    Label5.Enabled = False
    Label6.Enabled = False
    Form2.TxtFile.Enabled = False

    lstIP.Clear  'clear the list
    X1 = -1
    X2 = -1
      
      If Form2.TxtFile.Text = "" Then
      Form2.TxtFile.Text = "Results.tmp"
      MsgBox "You havent picked a proxy list, so we'll rescan your last scan reults."
      
      End If
      Open Form2.TxtFile.Text For Input As #1  'open the file
      Do Until EOF(1) = True  'go until the end of the file
        Input #1, GoIP
        X1 = X1 + 1
          If GoIP = "" Then
          Else
            lstIP.AddItem GoIP, X1  'add all the lines into the lstip listbox
          Label12.Caption = "Proxies: " & lstIP.ListCount
          End If
      Loop
      
      Close #1 'close the file


    Call ScanProxy
    
    
    Label6.Visible = True
    Form2.TxtFile.Enabled = True
    Label6.Enabled = True
End Sub
Private Sub Label6_Click()
'Added to keep program from crashing
    On Error Resume Next
lstWin.ListIndex = 0
lstResponse.ListIndex = 0
Timer2.Enabled = True
End Sub
Private Sub Label7_Click()
'Added to keep program from crashing
    On Error Resume Next
Close #1
Form2.TxtFile.Enabled = False
Timer1.Enabled = False
Timer2.Enabled = False
PB1.Enabled = False
Timer1.Enabled = True

End Sub

Private Sub Label9_Click()
Me.Hide
Form2.Show
End Sub





Private Sub Sck1_Connect()
    'Added to keep program from crashing
    On Error Resume Next
    ShapeLead.BackColor = &H0&
    DoEvents
End Sub
Sub ScanProxy()
'Added to keep program from crashing
    On Error Resume Next
Dim p As Integer
Dim IPAdd As String
Dim StartTime As Long
Dim StopTime As Long
Dim TimeElapsed As Long
Dim StringTime As String

'This is the scan
      X2 = 1
    PB1.Max = lstIP.ListCount - 1
    PB1.Min = 0
    
For i = 1 To lstIP.ListCount - 1

         IPAdd = Trim(lstIP.List(X2))
         lstIP.Selected(X2) = True    'Highlight the current selection
         p = InStr(1, IPAdd, ":", vbTextCompare)      'Pull the port number off
         txtPort.Text = Right(IPAdd, Len(IPAdd) - p)  'and put it in the textbox
         IPAdd = Left(IPAdd, p - 1)
         Label2.Caption = "Working Proxies: " & lstWin.ListCount
         Debug.Print X2 & "  :  " & IPAdd & " : " & txtPort.Text
         PB1.Visible = True
         PB1.Value = PB1.Value + 1
         'change progressbar colors
         Change_pb_ForeColor PB1.hWnd, &HF6C944
         Change_pb_Color PB1.hWnd, &H0&
         
         
         StartTime = timeGetTime   'Start our timeout counter
         Sck1.Connect IPAdd, txtPort.Text
    Do
         Select Case Sck1.State
             Case 7, 8, 9, 0
                 Exit Do
         End Select
         DoEvents
         'Timeout check, timeGetTime works in milliseconds
         If timeGetTime - StartTime > 1500 Then Exit Do ' 10 seconds max 10000
    Loop
    
    
       
       
    If Sck1.State = 7 Then
    'scrolls lstwin
    lstWin.AddItem lstIP.List(X2)
    lstWin.Selected(lstWin.NewIndex) = True
       
       'uncomment the next line to hear beep when proxy found
       'Beep
       ShapeLead.BackColor = &HF6C944
       StopTime = timeGetTime
       TimeElapsed = StopTime - StartTime
       'Pad the elapsed time so the listboxes sort property can work
       StringTime = Trim(Str(TimeElapsed))
       Select Case Len(StringTime)
        Case Is = 1
          StringTime = "0000" & StringTime
        Case Is = 2
          StringTime = "000" & StringTime
        Case Is = 3
          StringTime = "00" & StringTime
        Case Is = 4
          StringTime = "0" & StringTime
       End Select
       Debug.Print "It Responded in " & TimeElapsed
       
       lstResponse.AddItem "It Responded in " & TimeElapsed
       lstResponse.ListIndex = lstResponse.ListIndex + 1
       'THE FOLLOWINGS SAVES LISTWIN TO A TEMP FILE THEN READS IT

Dim Retval As Long
Dim X As Integer
    'Build a temporary file
    Open App.Path & "\Results.tmp" For Output As #1
      For X = 1 To lstWin.ListCount - 1
         IPAdd = Trim(lstWin.List(X))
    'comment out the next line if you want the times kept
    '     IPAdd = Right(IPAdd, Len(IPAdd) - 0)
         Print #1, IPAdd
      Next X
    Close #1
        
        
        DoEvents
        
    End If
       
       X2 = X2 + 1
       Sck1.Close
       ShapeLead.BackColor = &H0&
 If PB1.Value = lstIP.ListCount - 1 Then
 MsgBox "Scan Finished"
 PB1.Value = PB1.Min
 PB1.Visible = False
 Label5.Enabled = True
 End If
Next
     
    
    
    lstWin.Refresh
    
    lstResponse.Refresh
    
End Sub
Private Sub Timer1_Timer()
'Added to keep program from crashing
    On Error Resume Next
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyEnable", "1", REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyServer", lstWin.Text, REG_SZ
End Sub
Private Sub Timer2_Timer()
'Added to keep program from crashing
    On Error Resume Next

Sleep Form2.Text1.Text
'scroll thru lstwin
  If lstWin.ListIndex <> lstWin.ListCount - 1 Then
lstWin.Text = lstWin.Text + 1
lstWin.ListIndex = lstWin.ListIndex + 1
lstResponse.ListIndex = lstResponse.ListIndex + 1
ElseIf lstWin.ListIndex = lstWin.ListCount - 1 Then
lstResponse.ListIndex = lstResponse.ListIndex - 1
lstWin.ListIndex = 0
lstResponse.ListIndex = 0
  End If
  
End Sub
Private Sub FormDrag(TheForm As Form)
Dim X As Integer
ReleaseCapture
Call SendMessage(TheForm.hWnd, &HA1, 2, 0&)
'SNAP ON CODE
If Me.Left < 300 Then Me.Left = 0
If Me.Left > Screen.Width - (300 + Me.Width) Then Me.Left = Screen.Width - Me.Width
If Me.Top < 300 Then Me.Top = 0
If Me.Top > Screen.Height - (300 + Me.Height) Then Me.Top = Screen.Height - Me.Height
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Static lngMsg            As Long
    Dim blnflag              As Boolean

    lngMsg = X / Screen.TwipsPerPixelX

        If blnflag = False Then

            blnflag = True
        
                Select Case lngMsg
                    Case WM_RBUTTONCLK      'to popup menu on right-click
                        Call SetForegroundWindow(Me.hWnd)
                        Call RemoveBalloon(Me)
                        'Reference the menu object of the form below for popup
                        'PopupMenu Me.mnufile

                    Case WM_LBUTTONDBLCLK   'SHow form on left-dblclick
                        'Use line below if you want to remove tray icon on dbclick show form.
                        'If not, be sure to put Systrayoff in form unload and terminate events.
                        'Call SystrayOff(Me)
                        
                        Call SetForegroundWindow(Me.hWnd)
                        Call RemoveBalloon(Me)
                        Me.WindowState = vbNormal
                        Me.Show
                        Me.SetFocus
            
                End Select
        
            blnflag = False
        
        End If

Err.Clear

Static bBusy As Boolean
    If bBusy = False Then           'Do one thing at a time
        bBusy = True
        Select Case CLng(X \ 15)
            Case WM_LBUTTONDBLCLK   'Double-click left mouse button: same as selecting About
            'frmAbout.Show 1
            Case WM_LBUTTONDOWN     'Left mouse button pressed: change traffic light icon & tip
                
            Case WM_LBUTTONUP       'Left mouse button released
                Form1.Visible = True
                DoEvents
                AppActivate "BORDER"
                'Restore hidden windows:



            Case WM_RBUTTONDBLCLK   'Double-click right mouse button
            
            Case WM_RBUTTONDOWN     'Right mouse button pressed
            
            Case WM_RBUTTONUP       'Right mouse button released: display popup menu
                'PopupMenu frmmen.frmpop
        End Select
        bBusy = False
    End If
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Call the Form to Drag
FormDrag Me
End Sub

Private Function FakeForm(Form As Form, Titlebar As PictureBox, Forecolor As String)
'This is the Form1-Function
'It calls the other needed Functions and sets the neede Controls to needed Values
'Remember: Borderstyle must be 0 and u mustn't use menus
'If Form.BorderStyle <> 0 Then Exit Function
'MakeBorder Form
'Titlebar.Left = 60
'Titlebar.Top = 60
'Titlebar.Height = 270
'Titlebar.Width = Form1.Width - 125
Titlebar.AutoRedraw = True
'If Title = "" Then Title = Form.Caption
Titlebar.Forecolor = Forecolor
Titlebar.CurrentX = 3
Titlebar.CurrentY = (Titlebar.ScaleHeight - Titlebar.TextHeight(Title)) / 2
Titlebar.Print Title

End Function
