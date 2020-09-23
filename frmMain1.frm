VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTuTMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "P.A.D."
   ClientHeight    =   7005
   ClientLeft      =   3105
   ClientTop       =   4335
   ClientWidth     =   12975
   Icon            =   "frmMain1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   12975
   Begin VB.ListBox lstNamesSearch 
      Height          =   2985
      ItemData        =   "frmMain1.frx":08CA
      Left            =   120
      List            =   "frmMain1.frx":08CC
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   6375
   End
   Begin MSComctlLib.ImageList imglMain 
      Left            =   6720
      Top             =   6360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":08CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0BE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":0F02
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":121C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":1536
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":1850
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":1B6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":1E84
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain1.frx":219E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   6240
      Top             =   6480
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      ToolTipText     =   "Delete selected record"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      ToolTipText     =   "Edit selected record"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      ToolTipText     =   "Select next record"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Frame fraFullInfo 
      Caption         =   "Additional Information"
      Height          =   6015
      Left            =   6720
      TabIndex        =   26
      Top             =   600
      Width           =   6135
      Begin VB.TextBox txtComments 
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   37
         Top             =   4200
         Width           =   5895
      End
      Begin VB.TextBox txtBDAYYYYY 
         Height          =   285
         Left            =   3360
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   36
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtBDAYDD 
         Height          =   285
         Left            =   2640
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   35
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtBDAYMM 
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   34
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtCode 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   3480
         Width           =   1455
      End
      Begin VB.TextBox txtCountry 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtProv 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   2040
         Width           =   2655
      End
      Begin VB.TextBox txtCity 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtStreet 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   2655
      End
      Begin VB.TextBox txtPhoto 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtFax 
         Height          =   285
         Left            =   4080
         TabIndex        =   27
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Image picPhoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   3360
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblComments 
         Caption         =   "Additional Comments:"
         Height          =   255
         Left            =   120
         TabIndex        =   48
         Top             =   3960
         Width           =   1815
      End
      Begin VB.Label lblSlash2 
         Caption         =   "  /"
         Height          =   255
         Left            =   3120
         TabIndex        =   47
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label lblSlash1 
         Caption         =   "  /"
         Height          =   255
         Left            =   2400
         TabIndex        =   46
         Top             =   3480
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Birthday (MM/DD/YYYY):"
         Height          =   255
         Left            =   1920
         TabIndex        =   45
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label lblCode 
         Caption         =   "Area/Zip Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label lblCountry 
         Caption         =   "Country:"
         Height          =   255
         Left            =   120
         TabIndex        =   43
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label lblProvinceState 
         Caption         =   "Province/State:"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblCity 
         Caption         =   "City:"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblAddress 
         Caption         =   "Street Address:"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblPhoto 
         Caption         =   "Photo:"
         Height          =   255
         Left            =   3360
         TabIndex        =   39
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax:"
         Height          =   255
         Left            =   4080
         TabIndex        =   38
         Top             =   3240
         Width           =   1095
      End
   End
   Begin VB.Frame fraEntry 
      Caption         =   "Entries at a glance"
      Height          =   2535
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   6375
      Begin VB.CommandButton cmdExpand 
         Caption         =   "Expand to Full >>"
         Height          =   255
         Left            =   4560
         TabIndex        =   19
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtMiddleName 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   50
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton cmdCompact 
         Caption         =   "&Compact <<"
         Height          =   255
         Left            =   4560
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   5895
      End
      Begin VB.CommandButton cmdURL 
         Caption         =   "&Web Page:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Visit this person's web page"
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox txtPager 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtCell 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtPhone 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox txtEmail 
         Height          =   285
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.CommandButton cmdEmail 
         Caption         =   "E&mail:"
         Height          =   255
         Left            =   4200
         TabIndex        =   12
         ToolTipText     =   "Send this person an email"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox txtLastName 
         Height          =   285
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtFirstName 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblMiddleName 
         Caption         =   "M.I."
         Height          =   255
         Left            =   1680
         TabIndex        =   51
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lblPager 
         Caption         =   "Pager Number:"
         Height          =   255
         Left            =   4200
         TabIndex        =   25
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblCell 
         Caption         =   "Cell Number:"
         Height          =   255
         Left            =   2160
         TabIndex        =   24
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblPhone 
         Caption         =   "Phone Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblLastName 
         Caption         =   "Last Name:"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblFName 
         Caption         =   "First Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Add new record"
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<"
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      ToolTipText     =   "Select previous record"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      ToolTipText     =   "Search records"
      Top             =   3600
      Width           =   1095
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imglMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Previous"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Next"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Web Browser"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Email"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   49
      Top             =   6735
      Width           =   12975
      _ExtentX        =   22886
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   20214
            Text            =   "Personal Address Directory v2.0.0"
            TextSave        =   "Personal Address Directory v2.0.0"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   2593
            MinWidth        =   2593
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstNames 
      Columns         =   1
      Height          =   2985
      ItemData        =   "frmMain1.frx":24B8
      Left            =   120
      List            =   "frmMain1.frx":24BA
      TabIndex        =   8
      Top             =   600
      Width           =   6375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileAdd 
         Caption         =   "Add Entry"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileEdit 
         Caption         =   "Edit Entry"
         Enabled         =   0   'False
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Delete Entry"
         Enabled         =   0   'False
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuFileSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuUtil 
      Caption         =   "&Utilities"
      Begin VB.Menu mnuUtilEmail 
         Caption         =   "Compose New Email"
      End
      Begin VB.Menu mnuUtilWeb 
         Caption         =   "Launch Web Browser"
      End
      Begin VB.Menu mnuUtilSperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilSort 
         Caption         =   "Sort"
         Begin VB.Menu mnuUtilSortAZ 
            Caption         =   "A to Z"
         End
         Begin VB.Menu mnuUtilSortZA 
            Caption         =   "Z to A"
         End
         Begin VB.Menu mnuUtilSortSeperator1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUtilSortFirstLast 
            Caption         =   "First Name Last"
         End
         Begin VB.Menu mnuUtilSortLastLast 
            Caption         =   "Last Name Last"
         End
      End
   End
   Begin VB.Menu mnuClear 
      Caption         =   "&Clear Current Search"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
      End
      Begin VB.Menu mnuHelpSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuSystray 
      Caption         =   "Systray"
      Visible         =   0   'False
      Begin VB.Menu mnuSystrayLoad 
         Caption         =   "Load"
      End
      Begin VB.Menu mnuSystraySeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystrayAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuSystrayExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmTuTMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
   (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Option Explicit
Option Compare Text
Dim MyData As Database
Dim MyRecord As Recordset
Dim SQL As String



Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

Private Sub cmdAdd_Click()
    Unload Me
    frmAdd.Show
End Sub

Private Sub cmdBack_Click()
    If lstNames.ListIndex = 0 Then
        lstNames.ListIndex = (lstNames.ListCount - 1)
    Else
        lstNames.ListIndex = (lstNames.ListIndex - 1)
    End If
End Sub

Private Sub cmdCompact_Click()

    frmMain.Width = 6705
    
    cmdCompact.Visible = False
    cmdExpand.Visible = True


End Sub

Private Sub cmdDelete_Click()
    If MsgBox("You Really want to delete " & lstNames.Text, vbQuestion + vbYesNo, "Delete Record") = vbYes Then
        Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
        SQL = "SELECT * FROM PAD"
        Set MyRecord = MyData.OpenRecordset(SQL)
        MyRecord.Delete
        
        lstNames.RemoveItem (lstNames.ListIndex)
    Else
        Exit Sub
    End If
End Sub

Private Sub cmdEdit_Click()
    Unload Me
    frmEdit.Show
End Sub

Private Sub cmdEmail_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "mailto:" + txtEmail.Text
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL
End Sub

Private Sub cmdExpand_Click()
    
    frmMain.Width = 13080
    
    cmdExpand.Visible = False
    cmdCompact.Visible = True
    

End Sub

Private Sub cmdNext_Click()
    If lstNames.ListIndex = lstNames.ListCount - 1 Then
        lstNames.ListIndex = 0
    Else
        lstNames.ListIndex = (lstNames.ListIndex + 1)
    End If
End Sub

Private Sub cmdSearch_Click()
lstNamesSearch.Visible = True
mnuClear.Visible = True

Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
Dim LName As String
Dim strFind As String
lstNamesSearch.Clear
strFind = Trim(InputBox("Enter Text for search.", "Search Box"))
LName = Trim(UCase(strFind))

Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
Set MyRecord = MyData.OpenRecordset("PAD")

With MyRecord
    If .EOF Then
        MsgBox "No matching Names found, try again please", vbCritical, "Result"
    Else
        Do Until .EOF
            lstNamesSearch.AddItem MyRecord.Fields("FirstName") & vbTab & MyRecord.Fields("MiddleName") & vbTab & MyRecord.Fields("LastName")
            .MoveNext
        Loop
    End If
End With
End Sub

Private Sub cmdURL_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = txtURL.Text
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL
End Sub

Private Sub Form_Load()
    sbMain.Panels.Item(2) = Time
    Me.Width = 6705
    
Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
Set MyRecord = MyData.OpenRecordset("PAD")

    If MyRecord.EOF Then
        MsgBox "PAD Data not found.  Possible cause: Database is blank.  This will correct once you add data to the databse.", vbInformation, "Notice"
    Else
        MyRecord.MoveFirst
        Do Until MyRecord.EOF
            lstNames.AddItem MyRecord.Fields("FirstName") & vbTab & MyRecord.Fields("MiddleName") & vbTab & MyRecord.Fields("LastName")
            MyRecord.MoveNext
        Loop
    End If
    
tlbMain.Buttons(5).Enabled = False
tlbMain.Buttons(6).Enabled = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Sys As Long
Sys = X / Screen.TwipsPerPixelX
    Select Case Sys
        Case WM_LBUTTONDOWN:
            Me.PopupMenu mnuSystray
    End Select
End Sub

Private Sub Form_Resize()
If WindowState = vbMinimized Then
Me.Hide
Me.Refresh
With nid
.cbSize = Len(nid)
.hwnd = Me.hwnd
.uId = vbNull
.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
.uCallBackMessage = WM_MOUSEMOVE
.hIcon = Me.Icon
.szTip = Me.Caption & vbNullChar
End With
Shell_NotifyIcon NIM_ADD, nid
Else
Shell_NotifyIcon NIM_DELETE, nid
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub lstNames_Click()
On Error Resume Next
        Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
        Set MyRecord = MyData.OpenRecordset("PAD")
        MyRecord.MoveFirst
    Do Until MyRecord.EOF
            If lstNames.Text = MyRecord!FirstName & vbTab & MyRecord!MiddleName & vbTab & MyRecord!LastName Then
                txtFirstName.Text = MyRecord!FirstName
                txtMiddleName.Text = MyRecord!MiddleName
                txtLastName.Text = MyRecord!LastName
                txtEmail.Text = MyRecord!Email
                txtPhone.Text = MyRecord!Phone
                txtCell.Text = MyRecord!Cell
                txtPager.Text = MyRecord!Pager
                txtStreet.Text = MyRecord!Street
                txtCity.Text = MyRecord!City
                txtProv.Text = MyRecord!Prov
                txtCountry.Text = MyRecord!Country
                txtCode.Text = MyRecord!Zip
                txtBDAYMM.Text = MyRecord!BDAYMon
                txtBDAYDD.Text = MyRecord!BDAYDay
                txtBDAYYYYY.Text = MyRecord!BDAYYear
                txtComments.Text = MyRecord!Comments
                txtURL.Text = MyRecord!Url
                txtFax.Text = MyRecord!Fax
                txtPhoto.Text = MyRecord!Picture
            End If
            MyRecord.MoveNext
    Loop


If txtPhoto.Text = "" Then
    picPhoto.Picture = LoadPicture("")
Else
    picPhoto.Picture = LoadPicture("")
    
        Dim sExtension As String
            sExtension = UCase(Right$(txtPhoto, 3))
 
        If Dir$(txtPhoto) = "" Then
            Exit Sub
        End If
 
        Select Case sExtension
            Case "JPG", "GIF", "BMP"
            picPhoto.Picture = LoadPicture(txtPhoto.Text)
        Case Else
            MsgBox "Invaild Path", vbExclamation, "Link Field"
        End Select
    
End If

tlbMain.Buttons(5).Enabled = True
tlbMain.Buttons(6).Enabled = True
mnuFileEdit.Enabled = True
mnuFileDelete.Enabled = True

End Sub

Private Sub lstNamesSearch_Click()
On Error Resume Next
        Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
        Set MyRecord = MyData.OpenRecordset("PAD")
        MyRecord.MoveFirst
    Do Until MyRecord.EOF
            If lstNames.Text = MyRecord!FirstName & vbTab & MyRecord!MiddleName & vbTab & MyRecord!LastName Then
                txtFirstName.Text = MyRecord!FirstName
                txtMiddleName.Text = MyRecord!MiddleName
                txtLastName.Text = MyRecord!LastName
                txtEmail.Text = MyRecord!Email
                txtPhone.Text = MyRecord!Phone
                txtCell.Text = MyRecord!Cell
                txtPager.Text = MyRecord!Pager
                txtStreet.Text = MyRecord!Street
                txtCity.Text = MyRecord!City
                txtProv.Text = MyRecord!Prov
                txtCountry.Text = MyRecord!Country
                txtCode.Text = MyRecord!Zip
                txtBDAYMM.Text = MyRecord!BDAYMon
                txtBDAYDD.Text = MyRecord!BDAYDay
                txtBDAYYYYY.Text = MyRecord!BDAYYear
                txtComments.Text = MyRecord!Comments
                txtURL.Text = MyRecord!Url
                txtFax.Text = MyRecord!Fax
                txtPhoto.Text = MyRecord!Picture
            End If
            MyRecord.MoveNext
    Loop

If txtPhoto.Text = "" Then
    picPhoto.Picture = LoadPicture("")
Else
    picPhoto.Picture = LoadPicture("")
    
        Dim sExtension As String
            sExtension = UCase(Right$(txtPhoto, 3))
 
        If Dir$(txtPhoto) = "" Then
            Exit Sub
        End If
 
        Select Case sExtension
            Case "JPG", "GIF", "BMP"
            picPhoto.Picture = LoadPicture(txtPhoto.Text)
        Case Else
            MsgBox "Invaild Path", vbExclamation, "Link Field"
        End Select
    
End If
tlbMain.Buttons(5).Enabled = True
tlbMain.Buttons(6).Enabled = True
mnuFileEdit.Enabled = True
mnuFileDelete.Enabled = True
End Sub

Private Sub mnuClear_Click()
    lstNamesSearch.Visible = False
    mnuClear.Visible = False
End Sub

Private Sub mnuFileAdd_Click()
    Call cmdAdd_Click
End Sub

Private Sub mnuFileDelete_Click()
    Call cmdDelete_Click
End Sub

Private Sub mnuFileEdit_Click()
    Call cmdEdit_Click
End Sub

Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileSearch_Click()
    Call cmdSearch
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpContents_Click()
    MsgBox "No help Contents yet, please read the Help.doc file included."
End Sub

Private Sub mnuSystrayAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuSystrayExit_Click()
Shell_NotifyIcon NIM_DELETE, nid
End
End Sub

Private Sub mnuSystrayLoad_Click()
WindowState = vbNormal
Me.Show
End Sub

Private Sub mnuUtilEmail_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "mailto:"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL
End Sub

Private Sub mnuUtilSortAZ_Click()
Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim temp As String
 iCount = lstNames.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With lstNames
        If .List(i) > .List(i + 1) Then
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
        End If
     End With
    Next i
Next j
Call cmdFirst_Click
End Sub

Private Sub mnuUtilSortFirstLast_Click()
lstNames.Clear

Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
Set MyRecord = MyData.OpenRecordset("PAD")


        MyRecord.MoveFirst
        Do Until MyRecord.EOF
            lstNames.AddItem MyRecord.Fields("FirstName") & vbTab & MyRecord.Fields("MiddleName") & vbTab & MyRecord.Fields("LastName")
            MyRecord.MoveNext
        Loop
End Sub

Private Sub mnuUtilSortLastLast_Click()
lstNames.Clear

Set MyData = OpenDatabase(App.Path + "\Address.mdb")
Set MyRecord = MyData.OpenRecordset("PAD")


        MyRecord.MoveFirst
        Do Until MyRecord.EOF
            lstNames.AddItem MyRecord.Fields("LastName") & vbTab & MyRecord.Fields("MiddleName") & vbTab & MyRecord.Fields("FirstName")
            MyRecord.MoveNext
        Loop
End Sub

Private Sub mnuUtilSortMidd_Click()

End Sub

Private Sub mnuUtilSortZA_Click()
Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim temp As String
 iCount = lstNames.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With lstNames
        If .List(i) < .List(i + 1) Then
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
        End If
     End With
    Next i
Next j
Call cmdFirst_Click
End Sub

Private Sub mnuUtilWeb_Click()
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As String
    Dim sDirectory As String
    sTopic = "Open"
    sFile = "about:blank"
    sParams = 0&
    sDirectory = 0&
    
    RunShellExecute sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL
End Sub

Private Sub picPhoto_Click()
    frmPicFull.imgPic.Picture = frmMain.picPhoto.Picture
    frmPicFull.Show
End Sub

Private Sub sbMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    sbMain.ToolTipText = Date
End Sub

Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call cmdBack_Click
    Case 2
        Call cmdNext_Click
    Case 4
        Call mnuFileAdd_Click
    Case 5
        Call mnuFileEdit_Click
    Case 6
        Call mnuFileDelete_Click
    Case 8
        Call mnuFileSearch_Click
    Case 10
        Call mnuUtilWeb_Click
    Case 11
        Call mnuUtilEmail_Click
    Case 14
        Call mnuHelpContents_Click
End Select
End Sub

Private Sub tmrTime_Timer()
    sbMain.Panels.Item(2) = Time
End Sub

Public Sub RunShellExecute(sTopic As String, sFile As Variant, _
                           sParams As Variant, sDirectory As Variant, _
                           nShowCmd As Long)

   Dim hWndDesk As Long
   Dim success As Long
  
   hWndDesk = GetDesktopWindow()
  
   success = ShellExecute(hWndDesk, sTopic, sFile, sParams, sDirectory, nShowCmd)

  If success = SE_ERR_NOASSOC Then
    MsgBox "Couldn't load the default application"
    Call Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sFile, vbNormalFocus)
  End If
End Sub

