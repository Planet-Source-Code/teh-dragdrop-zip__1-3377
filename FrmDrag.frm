VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form FrmDrag 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drag & Drop Demonstration"
   ClientHeight    =   2505
   ClientLeft      =   240
   ClientTop       =   1545
   ClientWidth     =   7230
   Icon            =   "FrmDrag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox rt1 
      Height          =   2415
      Left            =   3120
      TabIndex        =   1
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   10000
      TextRTF         =   $"FrmDrag.frx":0442
   End
   Begin VB.ListBox List1 
      Height          =   2400
      ItemData        =   "FrmDrag.frx":04F0
      Left            =   0
      List            =   "FrmDrag.frx":04F2
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
End
Attribute VB_Name = "FrmDrag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'The settings for format are:
'Constant       Value       Description
'vbCFText       1           Text (.txt files)
'vbCFBitmap     2           Bitmap (.bmp files)
'vbCFMetafile   3           metafile (.wmf files)
'vbCFEMetafile  14          Enhanced metafile (.emf files)
'vbCFDIB        8           Device-independent bitmap (DIB)
'vbCFPalette    9           Color palette
'vbCFFiles      15          List of files
'vbCFRTF        -16639      Rich text format (.rtf files)

Private Sub Form_Load()
List1.OLEDropMode = 1
List1.OLEDragMode = 1
rt1.OLEDropMode = 1
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i%
If Data.GetFormat(vbCFFiles) Then
    Caption = Data.Files.Count & " object(s) selected"
    List1.Clear
    For i = 1 To Data.Files.Count
        List1.AddItem Data.Files(i%)
    Next i
End If
''' if drag text from rt1
'If Data.GetFormat(vbCFText) Then
'    List1.AddItem "Text : " & Data.GetData(vbCFText)
'End If
End Sub



Private Sub rt1_OLEDragDrop(Data As RichTextLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
If Data.GetFormat(vbCFText) Then
    rt1.LoadFile Data.GetData(vbCFText), rtfText
End If
If Data.GetFormat(vbCFFiles) Then
    rt1.LoadFile Data.Files(1), rtfText  'Load one file for demo
End If
Caption = rt1.FileName
End Sub

