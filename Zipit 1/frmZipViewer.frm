VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E8889BD8-5764-11D4-BDD7-E09052C10310}#3.0#0"; "RichsoftZipit.ocx"
Begin VB.Form frmZipViewer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zip Viewer"
   ClientHeight    =   4020
   ClientLeft      =   990
   ClientTop       =   2175
   ClientWidth     =   7800
   Icon            =   "frmZipViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7800
   Begin RichsoftZipit.Zipit Zipit1 
      Left            =   6720
      Top             =   2520
      _ExtentX        =   1746
      _ExtentY        =   1746
      UseDirectoryInfo=   0   'False
      CompressionLevel=   0
      AddAction       =   0
      ExtrAction      =   0
      IncludeSystemFiles=   0   'False
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog cdlZip 
      Left            =   5760
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".zip"
      DialogTitle     =   "Open Zip Archive"
      Filter          =   "Zip Files|*.zip"
      MaxFileSize     =   256
   End
   Begin ComctlLib.ListView lvwZip 
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "iglIcons"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Filename"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Date/Time"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Packed"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Ratio"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblWorkFile 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   7575
   End
   Begin VB.Label lblEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "richard@richsoftcomputing.cjb.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmZipViewer.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "mailto:richard@richsoftcomputing.cjb.net?subject=Zipit Control 1.0"
      ToolTipText     =   "Contact me"
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Label lblWeb 
      Caption         =   "www.richsoftcomputing.cjb.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmZipViewer.frx":045C
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Go to Richsoft Computing Website"
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Label lblRichsoft 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft Computing 2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
   End
   Begin VB.Label lblFiles 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   6600
      TabIndex        =   1
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open/Create Archive"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Archive"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuAdd 
         Caption         =   "A&dd File"
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "&Extract File"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete File"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Zipit Control"
      End
   End
   Begin VB.Menu mnuZipActions 
      Caption         =   "Zip Actions"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Delete..."
      End
      Begin VB.Menu mnuFileExtract 
         Caption         =   "Extract..."
      End
      Begin VB.Menu mnuFileView 
         Caption         =   "View..."
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu mnuInvertSelection 
         Caption         =   "Invert Selection"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "frmZipViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==============================================================================
'Richsoft Computing 2000
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'Please visit my website at www.geocities.com/richardsouthey.
'If you would like to make any comments/suggestions then please e-mail them to
'richardsouthey@hotmail.com.
'==============================================================================

'API Call which drives the Hyperlink
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Open File Status
Dim FileOpen As Boolean





Public Sub HyperJump(ByVal URL As String)
    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub









Private Sub Form_Load()
    'Set the open file status to false
    FileOpen = False
    'Set the title
    Me.Caption = App.ProductName
    'Get values from registry
    Zipit1.CompressionLevel = CInt(GetSetting("Richsoft Computing", "ZipViewer", "CompLevel", zipNormal))
    Zipit1.AddAction = CInt(GetSetting("Richsoft Computing", "ZipViewer", "AddAction", zipDefault))
    Zipit1.UseDirectoryInfo = CBool(GetSetting("Richsoft Computing", "ZipViewer", "UseDirectoryInfo", True))
    Zipit1.UseDOS83Format = CBool(GetSetting("Richsoft Computing", "ZipViewer", "DOS83Format", False))
    Zipit1.Overwrite = CBool(GetSetting("Richsoft Computing", "ZipViewer", "Overwrite", False))
    Zipit1.ExtrAction = CInt(GetSetting("Richsoft Computing", "ZipViewer", "ExtrAction", zipNormal))
    Zipit1.ExtractDir = CStr(GetSetting("Richsoft Computing", "ZipViewer", "ExtractDir", CurDir$))
End Sub


Private Sub Form_Unload(Cancel As Integer)
    'Save the Zipit Properties to the registry for next time
    SaveSetting "Richsoft Computing", "ZipViewer", "CompLevel", CStr(Zipit1.CompressionLevel)
    SaveSetting "Richsoft Computing", "ZipViewer", "AddAction", CStr(Zipit1.AddAction)
    SaveSetting "Richsoft Computing", "ZipViewer", "UseDirectoryInfo", CStr(Zipit1.UseDirectoryInfo)
    SaveSetting "Richsoft Computing", "ZipViewer", "DOS83Format", CStr(Zipit1.UseDOS83Format)
    SaveSetting "Richsoft Computing", "ZipViewer", "Overwrite", CStr(Zipit1.Overwrite)
    SaveSetting "Richsoft Computing", "ZipViewer", "ExtrAction", CStr(Zipit1.ExtrAction)
    SaveSetting "Richsoft Computing", "ZipViewer", "ExtractDir", Zipit1.ExtractDir

End Sub


Private Sub lblEmail_Click()
    HyperJump lblEmail.Tag
End Sub

Private Sub lblWeb_Click()
    'Active the Hyperlink
    HyperJump lblWeb.Caption
End Sub

Private Sub lvwZip_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'When right button is pressed check to see
    'if its over an item, and then show the actions
    'popup menu
    If Button <> vbRightButton Then Exit Sub
    
    'Check to see if over an item
    lvwZip.SelectedItem = lvwZip.HitTest(x, y)
    If lvwZip.SelectedItem Is Nothing Then Exit Sub
    
    'An item is selected so pop up the menu
    mnuFileOpen.Caption = "Open " & lvwZip.SelectedItem
    PopupMenu mnuZipActions, , , , mnuFileOpen
End Sub


Private Sub mnuAbout_Click()
    'Show the control's about box
    Zipit1.About
End Sub


Private Sub mnuAdd_Click()
    'Add a file to the archive
    'Check a archive is open
    If Not FileOpen Then Exit Sub
    
    'Show the Add dialog
    frmAdd.Show 1
End Sub

Private Sub mnuClose_Click()
    'Close the open archive
    Zipit1.Filename = ""
    FileOpen = False
    Me.Caption = App.Title
    
End Sub

Private Sub mnuExit_Click()
    'Exit the program
    Unload Me
End Sub





Private Sub mnuFileOpen_Click()
Dim Path As String
    Dim File As New Collection
    Dim OldDirInfo As Boolean
    Dim OldAction As ZipAction
    Dim OldCurDir As String
    Dim OldOverwrite As Boolean
    
    'Open the selected file with default app
    
    'Save current property values
    OldDirInfo = Zipit1.UseDirectoryInfo
    OldAction = Zipit1.ExtrAction
    OldOverwrite = Zipit1.Overwrite
    OldCurDir = CurDir$
    
    'First extract the file to the temp directory
    File.Add lvwZip.SelectedItem.SubItems(5) & lvwZip.SelectedItem
    Zipit1.ExtractDir = Environ$("TEMP") & "\"
    Zipit1.UseDirectoryInfo = False
    Zipit1.Overwrite = True
    Zipit1.ExtrAction = zipDefault
    r = Zipit1.Extract(File)
    If r > 0 Then
        'File extracted so view it
        HyperJump Environ$("TEMP") & "\" & lvwZip.SelectedItem
    Else
        MsgBox "Error Viewing"
    End If
    'Set old property values
    Zipit1.UseDirectoryInfo = OldDirInfo
    Zipit1.ExtrAction = OldAction
    Zipit1.Overwrite = OldOverwrite
    ChDir OldCurDir
End Sub


Private Sub mnuInvertSelection_Click()
    'Inverts the file lists selection
    Dim i As ListItem
    For Each i In lvwZip.ListItems
        i.Selected = Not i.Selected
    Next i
End Sub

Private Sub mnuOpen_Click()
    'Open an archive
    'Using a filename that does not exist will create a new archive
    On Error Resume Next
    cdlZip.ShowSave
    'Check if cancel was pressed
    If Err = cdlCancel Then Exit Sub
    
    Zipit1.Filename = cdlZip.Filename
    FileOpen = True
    'Set the caption
    Me.Caption = App.Title & " - " & Zipit1.Filename
End Sub


Private Sub mnuSelectAll_Click()
    'Selects the entire file list
    Dim i As ListItem
    For Each i In lvwZip.ListItems
        i.Selected = True
    Next i
End Sub

Private Sub Zipit1_OnArchiveUpdate()
    'The archive has been updated so refresh the list
    Dim itmX As ListItem
    Dim r As Long
    Dim num As Long
    Dim i As Long
    Dim Filename As String
    Dim Path As String
    Dim ret As Long
    Dim FileSections() As String
    Dim Files As New ZipFileEntry
    Dim Location As Long
    
    'Get the number of files in the archive
    r = Zipit1.vFiles.Count
    num = r
       
    'Clear the list
    lvwZip.ListItems.Clear
    
    'Loop through each file in the archive
    For i = 1 To r
        'Store file info in a variable for ease of use
        'because the intellisense will give help
        Set Files = Zipit1.vFiles.Item(i)
        With Files
            'Check for a folder entry
            If Right$(.Filename, 1) <> "/" Then
                'It's a file so add it to the list
                'Add a item to the list
                
                'Find the filename from the path, using the
                'reverse find string command
                'Uses this function for compatibility with VB5
                'ret =
                
                'This is the VB6 function
                Location = InStrRev(.Filename, "/", -1)
                Path = Left$(.Filename, Location)
                Filename = Mid$(.Filename, Location + 1)
                'The last subscript holds the filename
                
                Set itmX = lvwZip.ListItems.Add(, , Filename)
                'Add the info
                itmX.Tag = i ' this index is required for so archive operations
                itmX.SubItems(1) = .FileDateTime
                itmX.SubItems(2) = .CompressedSize
                itmX.SubItems(3) = .UncompressedSize
                itmX.SubItems(5) = Path
                'Trap div by zero
                If .UncompressedSize <> 0 Then
                    itmX.SubItems(4) = Format(CInt((1 - (.CompressedSize / .UncompressedSize)) * 100)) & "%"
                Else
                    itmX.SubItems(4) = "0%"
                End If
            Else
                'It's a folder so don't add it and decrease the number of files by one
                num = num - 1
            End If
        End With
    Next i
    
    'Show the amount of files in the archive
    lblFiles.Caption = Format(num) & " file(s) in archive"

End Sub


Private Sub Zipit1_OnDeleteProgress(Percentage As Integer, Filename As String)
    ProgressBar1.Value = Percentage
    lblWorkFile.Caption = Filename
End Sub


Private Sub Zipit1_OnUnzipProgress(Percentage As Integer, Filename As String)
    ProgressBar1.Value = Percentage
    lblWorkFile.Caption = Filename
End Sub


Private Sub Zipit1_OnZipComplete(Successful As Long)
    lblWorkFile.Caption = ""
    ProgressBar1.Value = 0
    MsgBox Successful & "files added to archive", vbInformation, App.ProductName
End Sub

Private Sub Zipit1_OnZipProgess(Percentage As Integer, Filename As String)
    ProgressBar1.Value = Percentage
    lblWorkFile.Caption = Filename
End Sub


