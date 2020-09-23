VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Entry"
   ClientHeight    =   8685
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6135
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox txtCountry 
      Height          =   315
      ItemData        =   "frmAdd.frx":08CA
      Left            =   120
      List            =   "frmAdd.frx":0A21
      TabIndex        =   10
      Top             =   5160
      Width           =   3135
   End
   Begin VB.ComboBox txtProv 
      Height          =   315
      ItemData        =   "frmAdd.frx":0F11
      Left            =   120
      List            =   "frmAdd.frx":0FD8
      TabIndex        =   9
      Top             =   4440
      Width           =   3135
   End
   Begin VB.ComboBox txtBDAYDD 
      Height          =   315
      ItemData        =   "frmAdd.frx":12EF
      Left            =   4800
      List            =   "frmAdd.frx":1350
      TabIndex        =   14
      Top             =   5880
      Width           =   735
   End
   Begin VB.ComboBox txtBDAYMM 
      Height          =   315
      ItemData        =   "frmAdd.frx":13C7
      Left            =   3240
      List            =   "frmAdd.frx":13EF
      TabIndex        =   13
      Top             =   5880
      Width           =   1575
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3975
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   4200
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox txtPhone 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtCell 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtPager 
      Height          =   285
      Left            =   4200
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   5895
   End
   Begin VB.TextBox txtStreet 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox txtCity 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   3135
   End
   Begin VB.TextBox txtCode 
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   5880
      Width           =   855
   End
   Begin VB.TextBox txtBDAYYYYY 
      Height          =   285
      Left            =   5520
      MaxLength       =   4
      TabIndex        =   15
      Top             =   5880
      Width           =   495
   End
   Begin VB.TextBox txtComments 
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   16
      Top             =   6600
      Width           =   5895
   End
   Begin VB.Timer tmrTime 
      Interval        =   1000
      Left            =   0
      Top             =   8040
   End
   Begin VB.TextBox txtPhoto 
      Height          =   285
      Left            =   4560
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer tmrTimer 
      Left            =   2520
      Top             =   4680
   End
   Begin VB.CommandButton cmdPhoto 
      Caption         =   "&Add Photo"
      Height          =   255
      Left            =   3480
      TabIndex        =   17
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txtFax 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Top             =   5880
      Width           =   1815
   End
   Begin MSComctlLib.ImageList imgladd 
      Left            =   3120
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":1455
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":176F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":1A89
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAdd.frx":1DA3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbAdd 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgladd"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Clear"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insert Picture"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Help"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comdia1 
      Left            =   2880
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   19
      Top             =   8415
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8149
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
   Begin VB.Image picPhoto 
      BorderStyle     =   1  'Fixed Single
      Height          =   2535
      Left            =   3480
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   600
      Width           =   975
   End
   Begin VB.Label lblPhone 
      Caption         =   "Phone Number:"
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblCell 
      Caption         =   "Cell Number:"
      Height          =   255
      Left            =   2160
      TabIndex        =   31
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblPager 
      Caption         =   "Pager Number:"
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblEmail 
      Caption         =   "Email:"
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label lblUrl 
      Caption         =   "URL:"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label lblAddress 
      Caption         =   "Street Address:"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblCity 
      Caption         =   "City:"
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label lblProvinceState 
      Caption         =   "Province/State:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label lblCountry 
      Caption         =   "Country:"
      Height          =   255
      Left            =   120
      TabIndex        =   24
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Label lblCode 
      Caption         =   "Zip Code:"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   5640
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Birthday (MM/DD/YYYY):"
      Height          =   255
      Left            =   3240
      TabIndex        =   22
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label lblComments 
      Caption         =   "Additional Comments:"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   6360
      Width           =   1815
   End
   Begin VB.Label lblFax 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   1200
      TabIndex        =   20
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAND 
         Caption         =   "Save and Add another"
      End
      Begin VB.Menu mnuFilePic 
         Caption         =   "Insert Picture"
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClear 
         Caption         =   "Clear"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuHelpSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyData As Database
Dim MyRecord As Recordset


Private Sub cmdPhoto_Click()
On Error GoTo DialogError
With comdia1
        .CancelError = True
        .Filter = "JPG File (*.jpg)|*.jpg|Bitmap File (*.bmp)|*.bmp|GIF File(*.gif)|*.gif|All Files(*.*)|*.*"
        .FilterIndex = 1
        .DialogTitle = "Select a Picture File"
        .ShowOpen
   txtPhoto.Text = .FileName
   
   End With

Call picupdateADD

DialogError:
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Show
End Sub

Private Sub mnuFileClear_Click()
    txtName.Text = ""
    txtEmail.Text = ""
    txtPhone.Text = ""
    txtCell.Text = ""
    txtPager.Text = ""
    txtStreet.Text = ""
    txtCity.Text = ""
    txtProv.Text = ""
    txtCountry.Text = ""
    txtCode.Text = ""
    txtBDAYMM.Text = ""
    txtBDAYDD.Text = ""
    txtBDAYYYYY.Text = ""
    txtComments.Text = ""
    txtURL.Text = ""
    txtFax.Text = ""
    txtPhoto.Text = ""
End Sub

Private Sub mnuFilePic_Click()
    Call cmdPhoto_Click
End Sub

Private Sub mnuFileSave_Click()
Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
Set MyRecord = MyData.OpenRecordset("PAD")

If txtName.Text = "" Then
    MsgBox "Please Enter a Name"
    Exit Sub
End If


With MyRecord
        .AddNew
        !Name = Trim(txtName.Text)
        !Email = Trim(txtEmail.Text)
        !Phone = Trim(txtPhone.Text)
        !Cell = Trim(txtCell.Text)
        !Pager = Trim(txtPager.Text)
        !Url = Trim(txtURL.Text)
        !Street = Trim(txtStreet.Text)
        !City = Trim(txtCity.Text)
        !Prov = Trim(txtProv.Text)
        !Country = Trim(txtCountry.Text)
        !Zip = Trim(txtCode.Text)
        !BDAYMon = Trim(txtBDAYMM.Text)
        !BDAYDay = Trim(txtBDAYDD.Text)
        !BDAYYear = Trim(txtBDAYYYYY.Text)
        !Comments = Trim(txtComments.Text)
        !Fax = Trim(txtFax.Text)
        !Picture = Trim(txtPhoto.Text)
        .Update
        

End With

    frmMain.Show
    Unload Me
    
    frmMain.cmdEdit.Enabled = False
    frmMain.cmdDelete.Enabled = False
    frmMain.mnuFileEdit.Enabled = False
    frmMain.mnuFileDelete.Enabled = False
    frmMain.cmdNext.Enabled = False
    frmMain.cmdBack.Enabled = False
    frmMain.tlbMain.Buttons(5).Enabled = False
    frmMain.tlbMain.Buttons(6).Enabled = False
    

End Sub

Private Sub mnuFileSaveAND_Click()
Set MyData = OpenDatabase(App.Path + "\PAD.mdb")
Set MyRecord = MyData.OpenRecordset("PAD")

If txtName.Text = "" Then
    MsgBox "Please Enter a Name"
    Exit Sub
End If


With MyRecord
        .AddNew
        !Name = Trim(txtName.Text)
        !Email = Trim(txtEmail.Text)
        !Phone = Trim(txtPhone.Text)
        !Cell = Trim(txtCell.Text)
        !Pager = Trim(txtPager.Text)
        !Url = Trim(txtURL.Text)
        !Street = Trim(txtStreet.Text)
        !City = Trim(txtCity.Text)
        !Prov = Trim(txtProv.Text)
        !Country = Trim(txtCountry.Text)
        !Zip = Trim(txtCode.Text)
        !BDAYMon = Trim(txtBDAYMM.Text)
        !BDAYDay = Trim(txtBDAYDD.Text)
        !BDAYYear = Trim(txtBDAYYYYY.Text)
        !Comments = Trim(txtComments.Text)
        !Fax = Trim(txtFax.Text)
        !Picture = Trim(txtPhoto.Text)
        .Update
        

End With

    txtName.Text = ""
    txtEmail.Text = ""
    txtPhone.Text = ""
    txtCell.Text = ""
    txtPager.Text = ""
    txtStreet.Text = ""
    txtCity.Text = ""
    txtProv.Text = ""
    txtCountry.Text = ""
    txtCode.Text = ""
    txtBDAYMM.Text = ""
    txtBDAYDD.Text = ""
    txtBDAYYYYY.Text = ""
    txtComments.Text = ""
    txtURL.Text = ""
    txtFax.Text = ""
    txtPhoto.Text = ""

End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpContents_Click()
    MsgBox "No help Contents yet, please read the Help.doc file included."
End Sub

Private Sub tlbAdd_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        Call mnuFileSave_Click
    Case 3
        Call mnuFileClear_Click
    Case 4
        Call cmdPhoto_Click
    Case 6
        Call mnuHelpContents_Click
End Select
End Sub

