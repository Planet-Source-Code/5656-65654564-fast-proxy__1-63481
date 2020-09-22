VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Fast Proxy:  Options"
   ClientHeight    =   5715
   ClientLeft      =   0
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
   ForeColor       =   &H00F6C944&
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":492A
   ScaleHeight     =   5715
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   0
      Picture         =   "Form2.frx":D2234
      ScaleHeight     =   5775
      ScaleMode       =   0  'User
      ScaleWidth      =   11040
      TabIndex        =   0
      Top             =   0
      Width           =   11040
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Caption         =   "Proxy"
         ForeColor       =   &H00F6C944&
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   555
         Width           =   2535
         Begin VB.OptionButton Option1 
            BackColor       =   &H00000000&
            Caption         =   "Enable"
            ForeColor       =   &H00F6C944&
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00000000&
            Caption         =   "Disable"
            ForeColor       =   &H00F6C944&
            Height          =   255
            Left            =   1560
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00000000&
         Caption         =   "Change Proxy Every Millisecond"
         ForeColor       =   &H00F6C944&
         Height          =   735
         Left            =   360
         TabIndex        =   15
         Top             =   1275
         Width           =   2535
         Begin VB.TextBox Text1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BeginProperty Font 
               Name            =   "Impact"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00F6C944&
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Text            =   "700"
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00000000&
         Caption         =   "Auto Update"
         ForeColor       =   &H00F6C944&
         Height          =   615
         Left            =   3000
         TabIndex        =   11
         Top             =   555
         Width           =   2535
         Begin VB.OptionButton Option3 
            BackColor       =   &H00000000&
            Caption         =   "On"
            ForeColor       =   &H00F6C944&
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00000000&
            Caption         =   "Off"
            ForeColor       =   &H00F6C944&
            Height          =   255
            Left            =   960
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00000000&
            Caption         =   "Look"
            ForeColor       =   &H00F6C944&
            Height          =   255
            Left            =   1800
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00000000&
         Caption         =   "Default Proxy List To Scan"
         ForeColor       =   &H00F6C944&
         Height          =   735
         Left            =   3000
         TabIndex        =   8
         Top             =   1275
         Width           =   6135
         Begin VB.CommandButton cmdFile 
            Cancel          =   -1  'True
            Caption         =   "Browse"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5280
            TabIndex        =   10
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox TxtFile 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00F6C944&
            Height          =   405
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   5040
         End
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
         Left            =   361
         Picture         =   "Form2.frx":19FB3E
         ScaleHeight     =   240
         ScaleWidth      =   255
         TabIndex        =   7
         Top             =   160
         Width           =   255
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00000000&
         Caption         =   "Run Program At Startup"
         ForeColor       =   &H00F6C944&
         Height          =   615
         Left            =   5640
         TabIndex        =   4
         Top             =   555
         Width           =   3495
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Caption         =   "Start Minimized"
            ForeColor       =   &H00F6C944&
            Height          =   255
            Left            =   1920
            TabIndex        =   22
            Top             =   240
            Width           =   1455
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00000000&
            Caption         =   "Enable"
            ForeColor       =   &H00F6C944&
            Height          =   225
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00000000&
            Caption         =   "Disable"
            ForeColor       =   &H00F6C944&
            Height          =   225
            Left            =   960
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00000000&
         Caption         =   "Current Proxy"
         ForeColor       =   &H00F6C944&
         Height          =   735
         Left            =   360
         TabIndex        =   2
         Top             =   2115
         Width           =   8775
         Begin VB.TextBox Text2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            ForeColor       =   &H00F6C944&
            Height          =   375
            Left            =   120
            Locked          =   -1  'True
            ScrollBars      =   1  'Horizontal
            TabIndex        =   3
            Top             =   240
            Width           =   8535
         End
      End
      Begin VB.CommandButton cmdFile2 
         Height          =   405
         Left            =   10440
         TabIndex        =   1
         Top             =   1515
         Visible         =   0   'False
         Width           =   435
      End
      Begin MSComDlg.CommonDialog CD1 
         Left            =   10440
         Top             =   915
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         CancelError     =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "< - - Back"
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Left            =   3600
         TabIndex        =   21
         Top             =   165
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Options"
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
         Left            =   8640
         TabIndex        =   20
         Top             =   5115
         Width           =   2175
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
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

Private Declare Function SendMessagePBChange Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Sub SetTrans()
'Free the memory set
If hRgn Then DeleteObject hRgn

'Scan the Bitmap and remove all transparent pixels from it, creating a new region
hRgn = GetBitmapRegion(Me.Picture, vbWhite)

'Set the Forms new Region
SetWindowRgn Me.hWnd, hRgn, True

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
SaveSetting App.EXEName, "MyString8", "MyText8", Check1.Value
End If
If Check1.Value = 0 Then
SaveSetting App.EXEName, "MyString8", "MyText8", Check1.Value
End If
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Local Error Resume Next
Err.Clear

Static bBusy As Boolean
    If bBusy = False Then           'Do one thing at a time
        bBusy = True
        Select Case CLng(X \ 15)
            Case WM_LBUTTONDBLCLK   'Double-click left mouse button: same as selecting About
            'frmAbout.Show 1
            Case WM_LBUTTONDOWN     'Left mouse button pressed: change traffic light icon & tip
                
            Case WM_LBUTTONUP       'Left mouse button released
                Form2.Visible = True
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


Private Sub Form_Terminate()
Call SystrayOff(Me)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim ReturnVal%
    ReleaseCapture
    ReturnVal% = SendMessage(hWnd, &HA1, 2, 0)
End Sub
Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal%
    ReleaseCapture
    ReturnVal% = SendMessage(hWnd, &HA1, 2, 0)
End Sub
Private Sub cmdFile_Click()
    'Added to keep program from crashing
    On Error Resume Next
    CD1.Filter = "All Files (*.*) | *.*"
    CD1.ShowOpen
    Form2.TxtFile.Text = CD1.FileName
    X1 = -1
    X2 = -1
End Sub

Private Sub Form_Load()
On Error Resume Next

'Center form - Cant use the forms properties, because when restored
' from the tray, it will go back to the center!

    Top = Screen.Height / 2 - Height / 2
    Left = Screen.Width / 2 - Width / 2
'Creates the Form
FakeForm Me, Picture1, RGB(255, 255, 255)
'your bmp's outline color must be 255,255,255, which is White
'make outline transparent
SetTrans

Dim MyText As Integer
CD1.InitDir = App.Path
Text1.Text = GetSetting(App.EXEName, "MyString1", "MyText1", "")
TxtFile.Text = GetSetting(App.EXEName, "MyString2", "MyText2", "")
Option1.Value = GetSetting(App.EXEName, "MyString3", "MyText3", "")
Option2.Value = GetSetting(App.EXEName, "MyString4", "MyText4", "")
Option3.Value = GetSetting(App.EXEName, "MyString5", "MyText5", "")
Option4.Value = GetSetting(App.EXEName, "MyString6", "MyText6", "")
Option5.Value = GetSetting(App.EXEName, "MyString7", "MyText7", "")
Option6.Value = GetSetting(App.EXEName, "MyString8", "MyText8", "")
Text2.Text = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyServer")
Check1.Value = GetSetting(App.EXEName, "MyString9", "MyText9", "")
End Sub

Private Sub Label10_Click()
Dim MyText As Integer
SaveSetting App.EXEName, "MyString1", "MyText1", Text1.Text
SaveSetting App.EXEName, "MyString2", "MyText2", TxtFile.Text
SaveSetting App.EXEName, "MyString3", "MyText3", Option1.Value
SaveSetting App.EXEName, "MyString4", "MyText4", Option2.Value
SaveSetting App.EXEName, "MyString5", "MyText5", Option3.Value
SaveSetting App.EXEName, "MyString6", "MyText6", Option4.Value
SaveSetting App.EXEName, "MyString7", "MyText7", Option5.Value
SaveSetting App.EXEName, "MyString8", "MyText8", Option6.Value
SaveSetting App.EXEName, "MyString9", "MyText9", Check1.Value





Form1.Show
Unload Me
End Sub



Private Sub Option1_Click()
'Added to keep program from crashing
    On Error Resume Next
Option1.Value = True
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyEnable", "1", REG_DWORD
Dim MyText As Integer
SaveSetting App.EXEName, "MyString3", "MyText3", Option1.Value
End Sub

Private Sub Option2_Click()
'Added to keep program from crashing
    On Error Resume Next
Option2.Value = True
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyEnable", "0", REG_DWORD
SetKeyValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\", "ProxyEnable", "0", REG_DWORD
Dim MyText As Integer
SaveSetting App.EXEName, "MyString4", "MyText4", Option2.Value

End Sub

Private Sub Option3_Click()
'Added to keep program from crashing
    On Error Resume Next
Option3.Value = True
Dim MyText As Integer
SaveSetting App.EXEName, "MyString5", "MyText5", Option3.Value
End Sub

Private Sub Option4_Click()
'Added to keep program from crashing
    On Error Resume Next
Option4.Value = True
Dim MyText As Integer
SaveSetting App.EXEName, "MyString6", "MyText6", Option4.Value

End Sub

'check the option buttons
     'Option2.Value = True
     'Option1.Value = False
     'CheckBox.Value = Checked
     'CheckBox.Value = Unchecked
'end
Private Sub Option5_Click()
'Added to keep program from crashing
    On Error Resume Next
Option5.Value = True
SetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run\", "Fast Proxy", App.Path & "\Fast Proxy.exe", REG_SZ
Dim MyText As Integer
SaveSetting App.EXEName, "MyString7", "MyText7", Option5.Value

End Sub

Private Sub Option6_Click()
'Added to keep program from crashing
    On Error Resume Next
Option6.Value = True
DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Run", "Fast Proxy"
Dim MyText As Integer
SaveSetting App.EXEName, "MyString8", "MyText8", Option6.Value

End Sub

Private Sub Option7_Click()
On Error Resume Next
Dim MyText As Integer
GetSetting App.EXEName, "MyString5", "MyText5", Option3.Value
GetSetting App.EXEName, "MyString6", "MyText6", Option4.Value



Shell App.Path & "\Updater.exe", vbNormalFocus
Call SystrayOff(Me)
End

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

