VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4575
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   4575
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Display images"
      Height          =   255
      Left            =   6360
      TabIndex        =   7
      Top             =   3720
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   4080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   240
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   6240
      TabIndex        =   3
      Top             =   0
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Download"
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   3345
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   6015
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   810
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   135
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Height          =   735
      Left            =   6360
      TabIndex        =   6
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Menu mnumenu 
      Caption         =   "menu"
      Begin VB.Menu mnuadd 
         Caption         =   "add url"
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu mnudownload 
         Caption         =   "Download Folder"
      End
      Begin VB.Menu space1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ParseFilename   As String
Private i               As Integer
Dim displayBool As Boolean
Dim strTotalSize As Long
Private Sub Check1_Click()
    displayBool = True
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    List2.Clear
    For i = 0 To List1.ListCount - 1
        WebBrowser1.Navigate List1.List(i)
        Do Until WebBrowser1.ReadyState = READYSTATE_COMPLETE
            DoEvents
            Loop
    Next i
    If Dir(App.Path & "\downloads", vbDirectory) = vbNullString Then
        MkDir App.Path & "\downloads"
    End If
    With List2
        For i = 0 To .ListCount - 1
            Label3.Caption = "Downloading file " & i + 1
            ParseFilename = Mid$(.List(i), InStrRev(.List(i), "/") + 1, Len(.List(i)) - InStrRev(.List(i), "/") + 1)
            Inet1.Execute Trim$(.List(i)), "GET"
            Do While Inet1.StillExecuting
                DoEvents
                Loop
                DoEvents
        Next i
    End With
    Label4.Caption = vbNullString
    Label1.Caption = vbNullString
    Label3.Caption = vbNullString
    strTotalSize = 0
    On Error GoTo 0
End Sub
Private Sub Form_Load()
    With List1
        .AddItem "http://www.3dtotal.com/home2/gallery/gallery.asp?cat=character"
        .AddItem "http://news.google.co.uk/nwshp?hl=en&tab=wn&q="
        .AddItem "www.pscode.com"
    End With
End Sub
Private Sub Inet1_StateChanged(ByVal State As Integer)
    Dim sByte() As Byte
    Dim ff       As Integer
    ff = FreeFile()
    On Error Resume Next
    Select Case State
        Case icResponseCompleted
            If Dir(App.Path & "\downloads\" & ParseFilename) <> vbNullString Then
                ParseFilename = Split(ParseFilename, ".")(0) & i & "." & Split(ParseFilename, ".")(1)
            End If
            Open App.Path & "\downloads\" & ParseFilename For Binary Access Write As ff
                Do
                    DoEvents
                    sByte = Inet1.GetChunk(4096, icByteArray)
                    If UBound(sByte) = -1 Then
                        Exit Do
                    End If
                    Put #ff, , sByte
                    ProgressBar1.Value = Seek(1) - 1
                    strTotalSize = strTotalSize + ProgressBar1.Value
                    Label4.Caption = "Total size " & Format(strTotalSize, "###,####") & " KB"
                    Label1.Caption = ProgressBar1.Value & " bytes " & FormatPercent(ProgressBar1.Value / ProgressBar1.Max)
                    Loop
                Close #ff
                Select Case Check1.Value
                    Case 0
                        Unload Form2
                    Case 1
                        Form2.Picture1.Picture = LoadPicture(App.Path & "\downloads\" & ParseFilename)
                        Form2.Show
                End Select
        Case icResponseReceived
            If LenB(Inet1.GetHeader("Content-Length")) > 0 Then
                ProgressBar1.Max = CLng(Inet1.GetHeader("Content-Length"))
            End If
    End Select
    ProgressBar1.Value = 0
    On Error GoTo 0
End Sub
Private Sub List1_MouseUp(Button As Integer, _
        Shift As Integer, _
        X As Single, _
        Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnumenu
    End If
End Sub
Private Sub mnuadd_Click()
    Dim strUrl As String
    strUrl = InputBox("add url")
    If Not strUrl = vbNullString Then
        List1.AddItem strUrl
    End If
End Sub
Private Sub mnudownload_Click()
    If Dir(App.Path & "\downloads", vbDirectory) <> vbNullString Then
        Shell "explorer " & App.Path & "\downloads", vbNormalFocus
    End If
End Sub
Private Sub mnuexit_Click()
    Unload Me
End Sub
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, _
        URL As Variant)
    Dim img As MSHTML.HTMLImg
    On Error Resume Next
    For Each img In WebBrowser1.Document.images
        If InStr(1, img.src, "http") Then
            List2.AddItem img.src
        End If
    Next img
    Label2.Caption = List2.ListCount & " Files found"
    On Error GoTo 0
End Sub
