VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add"
   ClientHeight    =   4920
   ClientLeft      =   240
   ClientTop       =   765
   ClientWidth     =   6735
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "Go!"
      Default         =   -1  'True
      Height          =   375
      Left            =   5520
      TabIndex        =   17
      Top             =   2160
      Width           =   1095
   End
   Begin ComctlLib.Slider sldCompression 
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3600
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      _Version        =   327682
      Max             =   9
      SelStart        =   5
      Value           =   5
   End
   Begin VB.CheckBox chkDirNames 
      Caption         =   "&Save full path info"
      Height          =   255
      Left            =   3480
      TabIndex        =   7
      Top             =   4440
      Width           =   2655
   End
   Begin VB.CheckBox chkDOS83 
      Caption         =   "Store filenames in &8.3 format"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   4440
      Width           =   3015
   End
   Begin VB.ComboBox cboAction 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   3840
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "Add &Folder"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddFile 
      Caption         =   "&Add File"
      Height          =   375
      Left            =   3360
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox lstAdd 
      Height          =   1035
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   3
      Top             =   2400
      Width           =   3135
   End
   Begin VB.FileListBox filListing 
      Height          =   1845
      Left            =   3360
      MultiSelect     =   1  'Simple
      TabIndex        =   2
      Top             =   120
      Width           =   3255
   End
   Begin VB.DirListBox dirListing 
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.DriveListBox drvListing 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label lblNormal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      Height          =   255
      Left            =   4920
      TabIndex        =   16
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblMax 
      BackStyle       =   0  'Transparent
      Caption         =   "Max"
      Height          =   255
      Left            =   6360
      TabIndex        =   15
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblStore 
      BackStyle       =   0  'Transparent
      Caption         =   "Store"
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label lblCompression 
      Caption         =   "&Compression Ratio:"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label lblAction 
      Caption         =   "Ac&tion:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label lblSelected 
      Caption         =   "The files below will be added to the archive:"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2160
      Width           =   3135
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdAddFile_Click()
    'Add selected file(s) to the list of included files
    Dim i As Long
    'Loop through the list and find the selected items
    For i = 0 To filListing.ListCount - 1
        If filListing.Selected(i) Then
            'Add it to the list, checking for path
            If Right$(filListing.Path, 1) = "\" Then
                lstAdd.AddItem filListing.Path & filListing.List(i)
            Else
                lstAdd.AddItem filListing.Path & "\" & filListing.List(i)
            End If
            'Cancel the selection
            filListing.Selected(i) = False
        End If
    Next i

End Sub

Private Sub cmdAddFolder_Click()
    'Add the entire folder to the list
    
    'Check for the final back slash and then add
    '*.* to the list
    If Right$(filListing.Path, 1) = "\" Then
        lstAdd.AddItem filListing.Path & "*.*"
    Else
        lstAdd.AddItem filListing.Path & "\*.*"
    End If
        
End Sub


Private Sub cmdCancel_Click()
    'Cancel this dialog box
    Unload Me
End Sub


Private Sub cmdOK_Click()
    Dim Files As New Collection
    Dim i As Long
    On Error GoTo errorHandler
    'Start the process
    'First check if there is something to do
    If lstAdd.ListCount < 1 Then
        MsgBox "There is nothing to do!", vbInformation, App.ProductName
        Exit Sub
    End If
    
    'Now set all the ZipIt properties
    With frmZipViewer.Zipit1
        .AddAction = cboAction.ItemData(cboAction.ListIndex)
        .CompressionLevel = sldCompression.Value
        .UseDirectoryInfo = -chkDirNames.Value
        .UseDOS83Format = -chkDOS83.Value
        
        'Set the files to process
        For i = 0 To lstAdd.ListCount - 1
            Files.Add lstAdd.List(i)
        Next i
        
        'Close this dialog box
        Unload Me
        
        'Start the processing
        .Add Files
    End With

Exit Sub
    
errorHandler:
    LogError "frmAdd.cmdOK_Click", Err, False
    Resume Next
End Sub

Private Sub dirListing_Change()
    'Change the file list with this directory list
    filListing.Path = dirListing.Path
End Sub


Private Sub drvListing_Change()
    'Change the directory lsiting to the current directory
    'of the select drive
    'Trap errors
    On Error GoTo errorHandler
    
    dirListing.Path = CurDir$(drvListing.Drive)
    
    Exit Sub
    
errorHandler:
    'The error handler
    MsgBox "The drive selected cannot be accessed", vbCritical, APP_TITLE
    Exit Sub
    Resume Next
End Sub


Private Sub filListing_DblClick()
    'If a file is double clicked add it to the list
    Dim i As Long
    'Loop through the list and find the selected items
    For i = 0 To filListing.ListCount - 1
        If filListing.Selected(i) Then
            'Add it to the list, checking for path
            If Right$(filListing.Path, 1) = "\" Then
                lstAdd.AddItem filListing.Path & filListing.List(i)
            Else
                lstAdd.AddItem filListing.Path & "\" & filListing.List(i)
            End If
            'Cancel the selection
            filListing.Selected(i) = False
        End If
    Next i
End Sub


Private Sub Form_Load()
    'Set all the options
    With frmZipViewer.Zipit1
        'Set the the store dir info check box
        chkDirNames.Value = -.UseDirectoryInfo
        'Set the store in DOS 8.3 format box
        chkDOS83.Value = -.UseDOS83Format
        
        'Add text to the action combo box
        cboAction.AddItem "Add (and replace) files"
        cboAction.ItemData(cboAction.NewIndex) = zipDefault
        cboAction.AddItem "Update (and add) files"
        cboAction.ItemData(cboAction.NewIndex) = zipUpdate
        cboAction.AddItem "Freshen (existing) files"
        cboAction.ItemData(cboAction.NewIndex) = zipFreshen
        
        
        'Set the action
        cboAction.ListIndex = (.AddAction - 1)
        'Set the compression level slider
        sldCompression.Value = .CompressionLevel
                
    End With
End Sub



Private Sub lstAdd_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    'Check to see if the file(s) are to be removed
    If KeyCode <> vbKeyDelete And KeyCode <> vbKeyBack Then Exit Sub

    'Remove the file(s) from the list
    'Loop through the list and check to see if it is selected
    For i = lstAdd.ListCount - 1 To 0 Step -1
        If lstAdd.Selected(i) Then
            'If selected then remove the item
            lstAdd.RemoveItem (i)
        End If
    Next i
End Sub


