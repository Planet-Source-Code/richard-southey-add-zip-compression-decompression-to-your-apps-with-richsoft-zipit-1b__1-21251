VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Richsoft Zipit - Beta Test"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4815
   FillColor       =   &H8000000F&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblWebsite 
      BackStyle       =   0  'Transparent
      Caption         =   "www.richsoftcomputing.btinternet.co.uk"
      DragIcon        =   "frmAbout.frx":030A
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Tag             =   "www.richsoftcomputing.btinternet.co.uk"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.Label lblRichsoft 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Richsoft Computing © 2000"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Image imgPic 
      Height          =   720
      Left            =   240
      Picture         =   "frmAbout.frx":045C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   720
   End
   Begin VB.Label lblZipIt 
      BackStyle       =   0  'Transparent
      Caption         =   "ZipIt 1.0 beta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   480
      Width           =   2655
   End
End
Attribute VB_Name = "frmAbout"
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

Public Sub HyperJump(ByVal URL As String)
    'Function to execute the Hyperlink
    Call ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Sub


Private Sub cmdOK_Click()
    'Close the about box
    Unload Me
End Sub






Private Sub lblWebsite_DragDrop(Source As Control, X As Single, Y As Single)
    'If the mouse is over the label, the control
    'must be in drag mode. In this case, the
    'DragDrop event occurs when the mouse is
    'clicked.
    If Source Is lblWebsite Then
        With lblWebsite
            Call HyperJump(.Tag)
            .Font.Underline = False
            .ForeColor = vbButtonText
        End With
    End If
End Sub


Private Sub lblWebsite_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    'If the control is in drag mode, you can detect
    'MouseLeave easily by observing the State parameter.
    
    If State = vbLeave Then
        With lblWebsite
            .Drag vbEndDrag
            .FontUnderline = False
            .ForeColor = vbButtonText
        End With
    End If
End Sub


Private Sub lblWebsite_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Enter drag mode on the first MouseMove
    'allows easy detection of MouseLeave.
    
    With lblWebsite
        .ForeColor = vbHighlight
        .Font.Underline = True
        .Drag vbBeginDrag
    End With
End Sub


