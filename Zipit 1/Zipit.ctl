VERSION 5.00
Begin VB.UserControl Zipit 
   BackStyle       =   0  'Transparent
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   960
   InvisibleAtRuntime=   -1  'True
   PropertyPages   =   "Zipit.ctx":0000
   ScaleHeight     =   975
   ScaleWidth      =   960
   ToolboxBitmap   =   "Zipit.ctx":0016
   Begin VB.Frame fra3D 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   735
      Begin VB.Image imgPic 
         Height          =   480
         Left            =   120
         Picture         =   "Zipit.ctx":0328
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "Zipit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'==============================================================================
'Richsoft Computing 2000
'Richard Southey
'This code is e-mailware, if you use it please e-mail me and tell me about
'your program.
'Please visit my website at www.geocities.com/richardsouthey.
'If you would like to make any comments/suggestions then please e-mail them to
'richardsouthey@hotmail.com.
'==============================================================================

'Zip archive collection
Public vFiles As New Collection
'Archive Filename
Private ZipFilename As String
'Compression Level
Private CompLevel As ZipLevel
'Extraction Directory
Private ExtrDir As String
'Use/store directory info
Private UseDirInfo As Boolean
'Add Action
Private AddFileAction As ZipAction
'Extract Action
Private ExtrFileAction As ZipAction
'Overwrite files
Private OverwriteFiles As Boolean
'Use DOS 8.3 format
Private DOS83Format As Boolean
'Recurse Subdirectories
Private RecurseSubs As Boolean
'Include System/Hidden Files
Private IncludeSysFiles As Boolean


'Compression Level values
Public Enum ZipLevel
    zipStore = 0
    zipLevel1 = 1
    zipSuperFast = 2
    zipFast = 3
    zipLevel4 = 4
    zipNormal = 5
    zipLevel6 = 6
    zipLevel7 = 7
    zipLevel8 = 8
    zipMax = 9
End Enum

'Actions
Public Enum ZipAction
    zipDefault = 1
    zipFreshen = 2
    zipUpdate = 3
End Enum

'Events
Event OnArchiveUpdate()
Event OnZipProgess(Percentage As Integer, Filename As String)
Event OnUnzipProgress(Percentage As Integer, Filename As String)
Event OnZipComplete(Successful As Long)
Event OnUnzipComplete(Successful As Long)
Event OnDeleteProgress(Percentage As Integer, Filename As String)
Event OnDeleteComplete(Successful As Long)


Private Function ConvertWildcards(Files As Collection, ByVal IncludeSysFiles As Boolean, ByVal Recurse As Boolean) As Collection
    'Checks the file list that will be added to the archive
    'and converts any wildcards into a list of files.
    Dim i As Long
    Dim r As String
    Dim ret As New Collection
    Dim Path As String
    Dim Buffer As String * MAX_PATH
    Dim Attributes As Integer
    Attributes = vbNormal Or vbReadOnly
    
    'Check to see if system/hidden files are to be included
    If IncludeSysFiles = True Then
        'Add these attributes to the search
        Attributes = Attributes Or vbSystem Or vbHidden
    End If
    
    'Loop through the file list collection one item at a time
    For i = 1 To Files.Count
        Debug.Print Files(i)
        'Parse the file specification to find the path
        Path = ParsePath(Files(i))
        'Find the files matching the pattern
        Debug.Print Files(i) & ":"
        r = Dir$(Files(i), Attributes)
        Do Until r = ""
            'Add the file to the new file list collection
            ret.Add Path & r
            Debug.Print Path & r
            'Move on to next file, if one exists
            r = Dir$()
        Loop
    Next i
    'Return the new file list collection
    Set ConvertWildcards = ret
End Function


Public Function Delete(Filenames As Collection) As Long
    'Extracts files from the archive
    Dim r As Long
    Dim i As Long
    Dim ZipFile As String
    Dim Filename As String
    On Error Resume Next
    'Set return value to 0
    Delete = 0
    
    'Store local copy of relevant properties.
    ZipFile = ZipFilename
    
    'Make sure their is an archive set
    If ZipFile = "" Then
        'Return a failed response
        Delete = 0
        Exit Function
    End If
    
    'Loop through the collection and extract each file
    For i = 1 To Filenames.Count
        'Check to see if it is the last file in the archive
        'Re-read the archive
        Read
        'Check the number of files
        If vFiles.Count <> 1 Then
            'Delete the file
            r = DeleteFile(ZipFile, Filenames(i))
            If r = True Then
                'Add one to the return value if success
                Delete = Delete + 1
            End If
        Else
            'Delete the archive so the 'no valid zip entries'
            'message is not shown
            Kill ZipFile
            'Add one to the success count
            Delete = Delete + 1
        End If
        'Trigger the progess event and make sure it can happen
        RaiseEvent OnDeleteProgress((i / Filenames.Count) * 100, Filenames(i))
        DoEvents
    Next i
    'Trigger the completed event
    RaiseEvent OnDeleteComplete(Delete)
    'Re-read the archive and trigger the refresh event
    Read
    RaiseEvent OnArchiveUpdate
End Function

Public Sub About()
    'Show the about box
    frmAbout.Show 1
End Sub

Public Function Add(ByVal Filenames As Collection) As Long
    'Adds files to the archive, it returns the amount of
    'file successfully added
    Dim Filename As String
    Dim locZipFile As String
    Dim locUseDirInfo As String
    Dim locDOS83Format As Boolean
    Dim locAction As ZipAction
    Dim locCompLevel As ZipLevel
    Dim locRecurse As Boolean
    Dim locIncludeSysFiles As Boolean
    Dim i As Long
    Dim r As Boolean
    'On Error Resume Next
    'Set return value to 0
    Add = 0
    
    'Store local copy of relevant properties.
    locZipFile = ZipFilename
    locUseDirInfo = UseDirInfo
    locDOS83Format = DOS83Format
    locAction = AddFileAction
    locCompLevel = CompLevel
    locRecurse = RecurseSubs
    locIncludeSysFiles = True 'IncludeSysFiles
    
    'Make sure a archive has been set
    If locZipFile = "" Then
        'Show function has failed
        Add = 0
        Exit Function
    End If
    'Check to see if their are any files in the archive
    If vFiles.Count = 0 Then
        'Delete the archive so the 'no valid zip entries found'
        'message is not shown
        Kill locZipFile
    End If
    
    'Update the file list converting wildcards into the files they represent
    Set Filenames = ConvertWildcards(Filenames, locIncludeSysFiles, locRecurse)
    'Loop through the filename collection adding each file
    For i = 1 To Filenames.Count
        'Add the file to the archive
        Debug.Print "File: " & Filenames(i)
        r = AddFile(locZipFile, Filenames(i), locUseDirInfo, locDOS83Format, locAction, locCompLevel)
        'If true then add one to the return value
        If r = True Then
            Add = Add + 1
        End If
        'Trigger the progress event and make sure it can happen
        RaiseEvent OnZipProgess((i / Filenames.Count) * 100, Filenames(i))
        DoEvents
    Next i
    'Tigger the complete event
    RaiseEvent OnZipComplete(Add)
    'Re-read the archive and trigger the refresh event
    Read
    RaiseEvent OnArchiveUpdate
End Function

Private Sub AddEntry(zFile As ZipFile)
    Dim xFile As New ZipFileEntry
    'Adds a file from the archive into the collection
    
    xFile.Version = zFile.Version
    xFile.Flag = zFile.Flag
    xFile.CompressionMethod = zFile.CompressionMethod
    xFile.CRC32 = zFile.CRC32
    xFile.FileDateTime = GetDateTime(zFile.Date, zFile.Time)
    xFile.CompressedSize = zFile.CompressedSize
    xFile.UncompressedSize = zFile.UncompressedSize
    xFile.FileNameLength = zFile.FileNameLength
    xFile.Filename = zFile.Filename
    xFile.ExtraFieldLength = zFile.ExtraFieldLength
    
    vFiles.Add xFile
End Sub

Public Function Extract(Filenames As Collection) As Long
    'Extracts files from the archive
    Dim r As Boolean
    Dim i As Long
    Dim Filename As String
    Dim locZipFile As String
    Dim locUseDirInfo As Boolean
    Dim locOverwrite As Boolean
    Dim locAction As ZipAction
    Dim locExtrDir As String
    'On Error Resume Next
    'Set return value to 0
    Extract = 0
    
    'Store local copy of relevant properties.
    locZipFile = ZipFilename
    locUseDirInfo = UseDirInfo
    locOverwrite = OverwriteFiles
    locAction = ExtrFileAction
    locExtrDir = ExtractDir
    
    
    'Make sure their is an archive set
    If locZipFile = "" Then
        'Return a failed response
        Extract = 0
        Exit Function
    End If
        
    'Loop through the collection and extract each file
    For i = 1 To Filenames.Count
        'Extract the file
        Debug.Print "File: " & Filenames(i)
        r = ExtractFile(locZipFile, Filenames(i), locExtrDir, locUseDirInfo, locOverwrite, locAction)
        If r = True Then
            'Add one to the return value if success
            Extract = Extract + 1
        End If
        'Trigger the progess event and make sure it can happen
        RaiseEvent OnUnzipProgress((i / Filenames.Count) * 100, Filenames(i))
        DoEvents
    Next i
    'Trigger the completed event
    RaiseEvent OnUnzipComplete(Extract)
End Function

Public Property Get ExtractDir() As String
Attribute ExtractDir.VB_ProcData.VB_Invoke_Property = "ZipitProperties"
    'Return the extraction directory
    ExtractDir = ExtrDir
End Property

Private Function ParsePath(Path As String)
    'Takes a full file specification and returns the path
    For a = Len(Path) To 1 Step -1
        If Mid$(Path, a, 1) = "\" Then
            ParsePath = Left$(Path, a - 1) & "\"
            Exit Function
        End If
    Next a
End Function

Public Property Get UseDirectoryInfo() As Boolean
    'Return the use directory info status
    UseDirectoryInfo = UseDirInfo
End Property
Public Property Let ExtractDir(New_Extractdir As String)
    'Update the extraction directory
    ExtrDir = New_Extractdir
    PropertyChanged "ExtractDir"
End Property

Public Property Let UseDirectoryInfo(New_UseDirectoryInfo As Boolean)
    'Update the use directory info status
    UseDirInfo = New_UseDirectoryInfo
    PropertyChanged "UseDirectoryInfo"
End Property
Public Property Let CompressionLevel(New_CompressionLevel As ZipLevel)
    'Update the compression level directory
    CompLevel = New_CompressionLevel
    PropertyChanged "CompressionLevel"
End Property
Public Property Let Filename(New_Filename As String)
    Dim r As Long
    Dim i As Long
    'Called when the filename is updated
    ZipFilename = New_Filename
    PropertyChanged "Filename"
    'Read in the contents of the file
    r = Read
    'Raise the update event
    RaiseEvent OnArchiveUpdate
End Property

Public Property Get Filename() As String
Attribute Filename.VB_ProcData.VB_Invoke_Property = "ZipitProperties"
    'Called when the filename is read
    Filename = ZipFilename
End Property

Public Property Get CompressionLevel() As ZipLevel
    'Called when the compression level is read
    CompressionLevel = CompLevel
End Property
Private Function GetDateTime(ZipDate As Integer, ZipTime As Integer) As Date
    'Converts the file date/time dos stamp from the archive
    'in to a normal date/time string
    
    Dim r As Long
    Dim FTime As FileTime
    Dim Sys As SYSTEMTIME
    Dim ZipDateStr As String
    Dim ZipTimeStr As String
    
    'Convert the dos stamp into a file time
    r = DosDateTimeToFileTime(CLng(ZipDate), CLng(ZipTime), FTime)
    'Convert the file time into a standard time
    r = FileTimeToSystemTime(FTime, Sys)

    ZipDateStr = Sys.wDay & "/" & Sys.wMonth & "/" & Sys.wYear
    ZipTimeStr = Sys.wHour & ":" & Sys.wMinute & ":" & Sys.wSecond

    GetDateTime = ZipDateStr & " " & ZipTimeStr
End Function
Public Function Read() As Long
    'Reads the archive and places each file into a collection
    Dim Sig As Long
    Dim ZipStream As Integer
    Dim Res As Long
    Dim zFile As ZipFile
    Dim Name As String
    Dim i As Integer
    
    'If the filename is empty return a empty file list
    If ZipFilename = "" Then
        Read = 0
        'Remove any files still in the list
        For i = vFiles.Count To 1 Step -1
            vFiles.Remove i
        Next i
        Exit Function
    End If
    
    'Clears the collection
    'begin
    'vFiles.Clear;
    For i = vFiles.Count To 1 Step -1
        vFiles.Remove i
    Next i
    
    'Opens the archive for binary access
    ZipStream = FreeFile
    Open ZipFilename For Binary As ZipStream
    'Loop through archive
    Do While True
        Get ZipStream, , Sig
        'See if the file header has been found
              If Sig = LocalFileHeaderSig Then
                    'Read each part of the file header
                    Get ZipStream, , zFile.Version
                    Get ZipStream, , zFile.Flag
                    Get ZipStream, , zFile.CompressionMethod
                    Get ZipStream, , zFile.Time
                    Get ZipStream, , zFile.Date
                    Get ZipStream, , zFile.CRC32
                    Get ZipStream, , zFile.CompressedSize
                    Get ZipStream, , zFile.UncompressedSize
                    Get ZipStream, , zFile.FileNameLength
                    Get ZipStream, , zFile.ExtraFieldLength
                    'Get the filename
                    'Set up a empty string so the right number of
                    'bytes is read
                    Name = String$(zFile.FileNameLength, " ")
                    Get ZipStream, , Name
                    zFile.Filename = Mid$(Name, 1, zFile.FileNameLength)
                    'Move on through the archive
                    'Skipping extra space, and compressed data
                    Seek ZipStream, (Seek(ZipStream) + zFile.ExtraFieldLength)
                    Seek ZipStream, (Seek(ZipStream) + zFile.CompressedSize)
                    'Add the fileinfo to the collection
                    AddEntry zFile
              Else
              Debug.Print Sig
                If Sig = CentralFileHeaderSig Or Sig = 0 Then
                    'All the filenames have been found so
                    'exit the loop
                    Exit Do
                'End
                Else
                If Sig = EndCentralDirSig Then
                    'Exit the loop
                    Exit Do
                End If
                End If
            End If
        Loop
        'Close the archive
        Close ZipStream
        'Return the number of files in the archive
        Read = vFiles.Count

    'Fire the update event
    RaiseEvent OnArchiveUpdate
End Function




Public Property Let UseDOS83Format(New_UseDOS83Format As Boolean)
    'Update the use DOS 8.3 Format property
    DOS83Format = New_UseDOS83Format
    PropertyChanged "UseDOS83Format"
End Property

Public Property Let RecurseSubFolders(New_RecurseSubFolders As Boolean)
    'Update the recurse sub folder property
    '**RECURSE NOT YET IMPLEMENTED**
    RecurseSubs = New_RecurseSubFolders
    PropertyChanged "RecurseSubFolders"
End Property
Public Property Let AddAction(New_AddAction As ZipAction)
    'Update the add action
    AddFileAction = New_AddAction
    PropertyChanged "AddAction"
End Property

Public Property Let ExtrAction(New_ExtrAction As ZipAction)
    'Update the extraction action
    ExtrFileAction = New_ExtrAction
    PropertyChanged "ExtrAction"
End Property


Public Property Let Overwrite(New_Overwrite As Boolean)
    'Update overwrite files property
    OverwriteFiles = New_Overwrite
    PropertyChanged "Overwrite"
End Property

Public Property Get UseDOS83Format() As Boolean
    'Return store in 8.3 format status
    UseDOS83Format = DOS83Format
End Property
Public Property Get RecurseSubFolders() As Boolean
    'Return recurse sub folder status
    '**RECURSE FUNCTION NOT YET IMPLEMENTED**
    RecurseSubFolders = RecurseSubs
End Property
Public Property Get AddAction() As ZipAction
    'Return add file action i.e update/freshen
    AddAction = AddFileAction
End Property

Public Property Get ExtrAction() As ZipAction
    'Return extract action - freshen/update etc
    ExtrAction = ExtrFileAction
End Property
Public Property Get Overwrite() As Boolean
    'Return store in 8.3 format status
    Overwrite = OverwriteFiles
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Get properties out of storage
    ExtrDir = PropBag.ReadProperty("ExtractDir", "")
    CompLevel = PropBag.ReadProperty("CompressionLevel", zipMax)
    ZipFilename = PropBag.ReadProperty("Filename", "")
    UseDirInfo = PropBag.ReadProperty("UseDirectoryInfo", True)
    OverwriteFiles = PropBag.ReadProperty("Overwrite", False)
    DOS83Format = PropBag.ReadProperty("UseDOS83Format", False)
    AddFileAction = PropBag.ReadProperty("AddAction", zipDefault)
    ExtrFileAction = PropBag.ReadProperty("ExtrAction", zipDefault)
    RecurseSubs = PropBag.ReadProperty("RecurseSubFolders", False)
    IncludeSysFiles = PropBag.ReadProperty("IncludeSystemFiles", True)
End Sub

Private Sub UserControl_Resize()
    'Fix the control's size
    UserControl.Size 990, 990
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Put properties into storage
    PropBag.WriteProperty "Filename", ZipFilename, ""
    PropBag.WriteProperty "ExtractDir", ExtrDir, ""
    PropBag.WriteProperty "UseDirectoryInfo", UseDirInfo, True
    PropBag.WriteProperty "CompressionLevel", CompLevel, zipMax
    PropBag.WriteProperty "Overwrite", OverwriteFiles, False
    PropBag.WriteProperty "UseDOS83Format", DOS83Format, False
    PropBag.WriteProperty "AddAction", AddFileAction, zipDefault
    PropBag.WriteProperty "ExtrAction", ExtrFileAction, zipDefault
    PropBag.WriteProperty "RecurseSubFolders", RecurseSubs, False
    PropBag.WriteProperty "IncludeSystemFiles", IncludeSysFiles, True
End Sub


