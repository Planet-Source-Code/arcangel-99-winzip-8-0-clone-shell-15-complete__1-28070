VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "WinZip"
   ClientHeight    =   4635
   ClientLeft      =   30
   ClientTop       =   660
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList 
      Left            =   6000
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0CCA
            Key             =   "&New Archive..."
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0FE6
            Key             =   "&Open Archive... "
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CC2
            Key             =   "Exit..."
         EndProperty
      EndProperty
   End
   Begin VB.Timer RemClip 
      Interval        =   500
      Left            =   6720
      Top             =   3240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   53
      ImageHeight     =   35
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":299E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F8A
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35D6
            Key             =   "Favorites"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3BCE
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4282
            Key             =   "Extract"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":490E
            Key             =   "View"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4EDE
            Key             =   "CheckOut"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":55C6
            Key             =   "Wizard"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5C1A
            Key             =   "AddDis"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":61E6
            Key             =   "ExtractDis"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6796
            Key             =   "ViewDis"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6CC2
            Key             =   "CheckOutDis"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   4320
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12435
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3360
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5927
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3352
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Modified"
         Object.Width           =   2470
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   1623
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Ratio"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Packed"
         Object.Width           =   1413
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Path"
         Object.Width           =   2822
      EndProperty
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   -70
      Width           =   7335
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   825
         Left            =   105
         TabIndex        =   3
         Top             =   165
         Width           =   7200
         _ExtentX        =   12700
         _ExtentY        =   1455
         ButtonWidth     =   1588
         ButtonHeight    =   1455
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Key             =   "New"
               Object.ToolTipText     =   "Create New Archive"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open"
               Key             =   "Open"
               Object.ToolTipText     =   "Open an Existing Archive"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Favorites"
               Key             =   "Favorites"
               Object.ToolTipText     =   "List Archives in your Favorite Zip Folders"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Add"
               Key             =   "Add"
               Object.ToolTipText     =   "Add files to the Archive"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Extract"
               Key             =   "Extract"
               Object.ToolTipText     =   "Extract files from the Archive"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "View"
               Key             =   "View"
               Object.ToolTipText     =   "View Files inside the Archive"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CheckOut"
               Key             =   "CheckOut"
               Object.ToolTipText     =   "Creat Icons for Files in the Archive"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Wizard"
               Key             =   "Wizard"
               Object.ToolTipText     =   "Activate the Winzip Wizard"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label LblFont 
      Caption         =   "LblFont"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   5640
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   " &File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Archive..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open Archive... "
      End
      Begin VB.Menu mnuFavorite 
         Caption         =   "&Favorite Zip Folders..."
      End
      Begin VB.Menu mnuCloseArchive 
         Caption         =   "&Close Archive..."
      End
      Begin VB.Menu mnuBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProperties 
         Caption         =   "Properties..."
      End
      Begin VB.Menu mnuCreateShortcut 
         Caption         =   "Create Shortcut"
      End
      Begin VB.Menu mnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveArchive 
         Caption         =   "Move Archive..."
      End
      Begin VB.Menu mnuCopyArchive 
         Caption         =   "Copy Archive..."
      End
      Begin VB.Menu mnuRenameArchive 
         Caption         =   "Rename Archive..."
      End
      Begin VB.Menu mnuDeleteArchive 
         Caption         =   "Delete Archive..."
      End
      Begin VB.Menu mnuBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print..."
      End
      Begin VB.Menu mnuWizard 
         Caption         =   "Wizard..."
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit..."
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuAdd 
         Caption         =   "Add.."
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete..."
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "Extract..."
      End
      Begin VB.Menu mnuView 
         Caption         =   "View..."
      End
      Begin VB.Menu mnuBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "Invert Selection"
      End
      Begin VB.Menu mnuBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuComment 
         Caption         =   "Comment..."
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuConfig 
         Caption         =   "Configuration..."
      End
      Begin VB.Menu mnuPassword 
         Caption         =   "Password..."
      End
      Begin VB.Menu mnuSort 
         Caption         =   "Sort"
         Begin VB.Menu mnuByName 
            Caption         =   "by Name"
         End
         Begin VB.Menu mnuByType 
            Caption         =   "by Type"
         End
         Begin VB.Menu mnuByDate 
            Caption         =   "by Date"
         End
         Begin VB.Menu mnuBySize 
            Caption         =   "by Size"
         End
         Begin VB.Menu mnuByCompression 
            Caption         =   "by Compression Ratio"
         End
         Begin VB.Menu mnuByPath 
            Caption         =   "by Path"
         End
         Begin VB.Menu mnuOrigionalOrder 
            Caption         =   "by Origional Order"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuBar6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveExit 
         Caption         =   "Save Setting on Exit..."
      End
      Begin VB.Menu mnuSaveNow 
         Caption         =   "Save Setting Now..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuContents 
         Caption         =   "Contents..."
      End
      Begin VB.Menu mnuTips 
         Caption         =   "Tip of the Day..."
      End
      Begin VB.Menu mnuBar7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFAQ 
         Caption         =   "FAQ..."
      End
      Begin VB.Menu mnuBar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Winzip..."
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "PopupMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuShowText 
         Caption         =   "Show Text"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

On Error Resume Next


'\\ Make sure ListView is in Report Style.
ListView1.View = lvwReport




'\\ Make sure Menu Icons are made.
  Set CoolMenuObj = New CoolMenu
   
  Call CoolMenuObj.Install(Me.hwnd, ImageList, True, True)

  LblFont.Caption = Me.FontName
  LblFont.Font = Me.Font
  
  
  
  
  
  
  '\\ Set the Wift and Height the for is allow it move to
  SetClipVars 5325, 7455
  
  
  
   '\\ Get the settings for where the form should move to.
  
    With App
        Me.Left = GetSetting(.Title, "Settings", "MainLeft", 0)
        Me.Top = GetSetting(.Title, "Settings", "MainTop", 0)
        Me.Width = GetSetting(.Title, "Settings", "MainWidth", 0)
        Me.Height = GetSetting(.Title, "Settings", "MainHeight", 0)
        Me.WindowState = GetSetting(.Title, "Settings", "WindowState", Me.WindowState)
    End With
  
  
   
  

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

'\\ Remove the Icons in the menu to save memory

  Call CoolMenuObj.Install(0&)
  
  Set CoolMenuObj = Nothing

End Sub

Private Sub Form_Resize()

On Error Resume Next

'\\ Set the Left most Property of the Controls.
Frame1.Left = "0"
ListView1.Left = "0"

'\\ Set the Width of the Controls while being sized.
Frame1.Width = frmMain.Width - 120
ListView1.Width = frmMain.Width - 120

'\\ Set the Top most Property of the Controls.
Frame1.Top = -70
ListView1.Top = 995

'\\ Set the Height of the controls while being sized.
ListView1.Height = frmMain.Height - 2000


'\\ Make it so that the Height and Width do not go below a certain number.
'If Me.Width <= 7455 Then Me.Width = 7455

'If Me.Height <= 5325 Then Me.Height = 5325

ClipForForm Me, 5325, 7455


End Sub


Private Sub Form_Unload(Cancel As Integer)

'\\ Save the Form Position and Setting upon exit.

With App
    
        If Me.WindowState <> vbMinimized Then
            
            SaveSetting .Title, "Settings", "WindowState", Me.WindowState
            
            If Me.WindowState <> vbMaximized Then
                SaveSetting .Title, "Settings", "MainLeft", Me.Left
                SaveSetting .Title, "Settings", "MainTop", Me.Top
                SaveSetting .Title, "Settings", "MainWidth", Me.Width
                SaveSetting .Title, "Settings", "MainHeight", Me.Height
            End If
            
        End If
        
     
    End With
    
    
RemoveClipping

End Sub


Private Sub mnuAbout_Click()

'\\ Show the About Box.
frmAbout.Show

End Sub


Private Sub mnuExit_Click()

'\\ Close all Open Windows to Save Memory then End the Application.
Call Exit_App

End Sub


Private Sub Refresh_Timer()
Me.Refresh
End Sub

Private Sub RemClip_Timer()
'\\ This Removes the automatic clipping that occurs when the form loads

RemoveClipping
RemClip.Enabled = False
End Sub

