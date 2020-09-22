VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form FormUpdater 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Fast Proxy Live Update"
   ClientHeight    =   5835
   ClientLeft      =   105
   ClientTop       =   0
   ClientWidth     =   6885
   ForeColor       =   &H00F6C944&
   Icon            =   "Updater.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox TaskText 
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Text            =   "Fast Proxy"
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   5040
   End
   Begin VB.ListBox List 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F6C944&
      Height          =   1200
      Left            =   1680
      TabIndex        =   4
      Top             =   3600
      Width           =   3375
   End
   Begin VB.Timer Timer2 
      Left            =   5400
      Top             =   5040
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   8070
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5460
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Updater.frx":0442
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
            Picture         =   "Updater.frx":B45C
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "Updater.frx":16476
         EndProperty
      EndProperty
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4320
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      URL             =   "http://"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F6C944&
      Height          =   255
      Left            =   6480
      TabIndex        =   6
      Top             =   25
      Width           =   255
   End
   Begin VB.Label lblConnect 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Click To Begin Live Update"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F6C944&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label latver 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00F6C944&
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Label curver 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00F6C944&
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   4320
      Width           =   855
   End
End
Attribute VB_Name = "FormUpdater"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal iparam As Long) As Long

Private Sub Label1_Click()
TerminateTask TaskText.Text
Shell App.Path & "\Fast Proxy.exe", vbNormalFocus
End
End Sub

Private Sub StatusBar1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal%
    ReleaseCapture
    ReturnVal% = SendMessage(hwnd, &HA1, 2, 0)
End Sub

Private Sub form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ReturnVal%
    ReleaseCapture
    ReturnVal% = SendMessage(hwnd, &HA1, 2, 0)
End Sub
Private Sub FormDrag(TheForm As Form)
Dim X As Integer
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
'SNAP ON CODE
If Me.Left < 300 Then Me.Left = 0
If Me.Left > Screen.Width - (300 + Me.Width) Then Me.Left = Screen.Width - Me.Width
If Me.Top < 300 Then Me.Top = 0
If Me.Top > Screen.Height - (300 + Me.Height) Then Me.Top = Screen.Height - Me.Height
End Sub
Private Sub Form_Load()
'Navigate to a file with instructions, or just an oprning page.
WebBrowser1.Navigate App.Path & "\Updater.html"
'On Local Error GoTo 200
If Dir("latver.dat") <> "" Then Kill "latver.dat"
 myVer = App.Major & "." & App.Minor & "." & App.Revision


' this is where the updated program needs to write it's current version
' number to.  The above commented out line puts the version number in
' the correct format.

status$ = "Idle"
UpdateTime = 0


'Open App.Path & "\version.dat" For Input As #1
'    Input #1, myVer
'Close #1



Path = "http://secure.servequake.com/Update/Fast_Proxy/"

    FormUpdater.Show
    DoEvents
    If Dir("version.dat") = "" Then
        CurrentVersion = ""
    Else
        Open "version.dat" For Input As #1
        If Not EOF(1) Then
            Line Input #1, CurrentVersion
        End If
        Close #1
    End If


DownloadFile "Ver.dat", "latver.dat"

    DoEvents
    If Dir("latver.dat") = "" Then
        LatestVersion = ""
    Else
        Open "latver.dat" For Input As #21
        If Not EOF(21) Then
            Line Input #21, LatestVersion
        End If
        Close #21
    End If
    
    
     DoEvents


    If CurrentVersion <> "" Then
        FormUpdater.curver = "Version " & CurrentVersion
        FormUpdater.List.AddItem "Your version is " & CurrentVersion
        
     If LatestVersion <> "" Then
        FormUpdater.List.AddItem "Latest version is " & LatestVersion
    If LatestVersion = CurrentVersion Then
        FormUpdater.List.AddItem "You have the latest version."
 End If
    
    
    
    
    Else
        FormUpdater.curver = "No version info"
    End If

End If

Exit Sub

200 myVer = "1.0.0.0"
X = MsgBox("Version information has not been found, Live Update will assume it's Version 1.0.0.0")

Resume 205

205
End Sub


Private Sub Timer1_Timer()
If Inet1.StillExecuting = False Then
    StatusBar1.Panels(1).Text = "Status: Idle"
Else
    StatusBar1.Panels(1).Text = "Status: " & status$
End If

End Sub

Private Sub Timer2_Timer()
    UpdateTime = UpdateTime + 1
    StatusBar1.Panels(2).Text = "Download Time:" & Str$(UpdateTime) & " Seconds"
End Sub

Private Sub lblConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

lblConnect.BackColor = &HC0C0C0
End Sub

Private Sub lblConnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lblConnect.BackColor = &H808080
Dim St As String
Dim fSt As String
Dim i As Long
i = 0
List.AddItem "Retrieving version info."
DownloadFile "Ver.dat", "latver.dat"

Open "latver.dat" For Input As #21
    Input #21, St
    latver = St
    List.AddItem "Latest version is " & St
    If CurrentVersion = latver Then
        List.AddItem "You have the latest version."
    'heres where ya load form 1 when no update found
        Call Pause(3)
   TerminateTask TaskText.Text
   Shell App.Path & "\Fast Proxy.exe", vbNormalFocus
   End
   
    Else
   'closes form one so it dont interfere with update
        
        List.AddItem "Retrieving file list."
        DownloadFile "files.dat", "update.dat"
        DoEvents
        List.AddItem "Comparing file lists."
        DoFiles
        DoEvents
        
        Open "update.dat" For Input As #20
            Do Until EOF(20)
                Line Input #20, fSt
                GetSections fSt, ","
                CheckFile Section(1), Section(2)
            Loop
        Close #20
        CreateFileList
        If Dir("version.dat") <> "" Then Kill ("version.dat")
        Open "version.dat" For Output As #24
            Print #24, latver
        Close #24
        List.AddItem "Version has been updated to " & latver & "."
        curver = "Version " & latver
        List.AddItem "Update complete."
        
       'heres where ya shell and stop after file downloads
        Call Pause(2)
       
    X = MsgBox("Live Update Complete! Fast Proxy Will Now Close, Please Restart It", vbInformation)
   
   Shell App.Path & "\Fast Proxy.exe", vbNormalFocus
   End
        
    End If
    DoEvents
Close #21
If Dir("vers.dat") <> "" Then Kill "vers.dat"
If Dir("update.dat") <> "" Then Kill "update.dat"
End Sub



