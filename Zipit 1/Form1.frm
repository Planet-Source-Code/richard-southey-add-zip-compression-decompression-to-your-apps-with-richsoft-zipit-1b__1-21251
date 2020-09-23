VERSION 5.00
Object = "*\AZipitControl.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin RichsoftZipit.Zipit Zipit1 
      Left            =   1440
      Top             =   600
      _ExtentX        =   1746
      _ExtentY        =   1746
      Filename        =   "hh"
      UseDirectoryInfo=   0   'False
      CompressionLevel=   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
