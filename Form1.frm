VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "File Binder - New Document"
   ClientHeight    =   6945
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   463
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFile 
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   3300
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8205
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   3375
      Left            =   105
      TabIndex        =   1
      Top             =   3465
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   5953
      _Version        =   393217
      BackColor       =   -2147483633
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   105
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuBlank_File_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdd 
         Caption         =   "&Add"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "&Remove"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "&Extract"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuBlank_File_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save As..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuBlank_File_2 
         Caption         =   "-"
      End
      Begin VB.Menu mniExit 
         Caption         =   "&Exit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================
'= FILE PACKER             =
'=  By James J. Kelly Jr.  =
'===========================

'By using this code\application you agree to the following:
'I provide no warrentys for this application or anything else
'I take no resposibility for anything that happens to your or anybody elses computer

'-Thank you for your time

'©Backwoods Interactive 2003. All rights reserved
'All code in this application is by Backwoods Interactive

Private Type STRUCT_FILE
 Version As String
 Filename() As String
 FileContent() As String
End Type

Dim sFile As STRUCT_FILE

Private Function GetName(ByVal sPath As String)

Dim i As Long

i = Len(sPath)

Do Until Mid(sPath, i, 1) = "\"
i = i - 1
GetName = GetName + Mid(sPath, i, 1)
Loop

GetName = ""

For h = i + 1 To Len(sPath)
GetName = GetName + Mid(sPath, h, 1)
Next h

End Function

Private Sub Form_Load()

Dim Free As Long
Free = FreeFile

sFile.Version = Hex(100)
ReDim sFile.FileContent(0)
ReDim sFile.Filename(0)

On Error Resume Next
Dim Width As Long
Dim Height As Long

Me.Width = GetSetting("FileBind", "GUI", "Width")
Me.Height = GetSetting("FileBind", "GUI", "Height")
Me.Left = GetSetting("FileBind", "GUI", "XCoord")
Me.Top = GetSetting("FileBind", "GUI", "YCoord")

If Command <> "" Then
'If LCase(Left(Command, 6)) = "-open " Then

Open Command For Binary As #Free
 
 Get #Free, , sFile
 
 If sFile.Version = Hex(100) Then
 
 'Put any code you would want here
  Me.Caption = "File Binder - " + Command
 
 Else
 
 MsgBox ("This file is corrupt or made by a newer or older version of file bind!")
 ReDim sFile.FileContent(0)
 ReDim sFile.Filename(0)
 sFile.Version = ""
 
 Close #Free
 
 Exit Sub
 
 End If


ListFile

End If
'End If

End Sub

Private Sub Form_Resize()

On Error Resume Next
lstFile.Height = Me.ScaleHeight / 2
RichTextBox1.Top = lstFile.Top + lstFile.Height + 6
RichTextBox1.Height = Me.ScaleHeight - RichTextBox1.Top - RichTextBox1.Left
RichTextBox1.Width = Me.ScaleWidth - RichTextBox1.Left - 6.5
lstFile.Width = RichTextBox1.Width

End Sub

Private Sub Form_Unload(Cancel As Integer)

SaveSetting "FileBind", "GUI", "Width", Me.Width
SaveSetting "FileBind", "GUI", "Height", Me.Height
SaveSetting "FileBind", "GUI", "XCoord", Me.Left
SaveSetting "FileBind", "GUI", "YCoord", Me.Top

 ReDim sFile.FileContent(0)
 
 ReDim sFile.Filename(0)
 
 sFile.Version = ""
 
 Unload Me
 
 End

End Sub

Private Sub lstFile_Click()

RichTextBox1.Text = sFile.FileContent(lstFile.ListIndex + 1)

End Sub

Private Sub mniExit_Click()
 
 ReDim sFile.FileContent(0)
 
 ReDim sFile.Filename(0)
 
 sFile.Version = ""
 
 Unload Me
 
 End
 
End Sub

Private Sub mnuAdd_Click()

On Error Resume Next

Dim Free As Long
Dim Buffer As String
Free = FreeFile

cd.Filter = "All Files (*.*)|*.*"
cd.DialogTitle = "Add File"
cd.ShowOpen

If cd.Filename <> "" Then

Dim Result As Integer

For h = 1 To UBound(sFile.Filename)

If sFile.Filename(h) = GetName(cd.Filename) Then Result = MsgBox("This file already exists!", vbOKOnly + vbExclamation): Exit Sub

Next h

Me.MousePointer = vbHourglass

Open cd.Filename For Binary As #Free

ReDim Preserve sFile.FileContent(UBound(sFile.FileContent) + 1)
Buffer = Space(LOF(Free))

Get #Free, , Buffer

sFile.FileContent(UBound(sFile.FileContent, 1)) = Buffer
Buffer = ""

ReDim Preserve sFile.Filename(UBound(sFile.Filename, 1) + 1)
sFile.Filename(UBound(sFile.Filename, 1)) = GetName(cd.Filename)
Close #Free

ListFile

Me.MousePointer = vbArrow

End If
'MsgBox (Str(UBound(sFile.FileName)))

End Sub



Private Sub mnuExtract_Click()

On Error Resume Next

Dim Free As Long
Free = FreeFile

cd.Filename = sFile.Filename(lstFile.ListIndex + 1)
cd.Filter = "All Files (*.*)|*.*"
cd.DialogTitle = "Extract As..."
cd.ShowSave

If cd.Filename <> "" Then

Me.MousePointer = vbHourglass

Kill cd.Filename

Open cd.Filename For Binary As #Free
 Put #Free, , sFile.FileContent(lstFile.ListIndex + 1)
Close #Free

Me.MousePointer = vbArrow

End If

End Sub

Private Sub mnuNew_Click()

 ReDim sFile.FileContent(0)
 
 ReDim sFile.Filename(0)
 
 sFile.Version = ""
 
 Me.Caption = "File Binder - New Document"
 
 ListFile
 
End Sub

Private Sub mnuOpen_Click()

On Error Resume Next

Dim Free As Long
Free = FreeFile

cd.Filter = "File Package (*.pck)|*.pck"
cd.DialogTitle = "Open"
cd.ShowOpen

If cd.Filename <> "" Then

Me.MousePointer = vbHourglass

Open cd.Filename For Binary As #Free
 
 Get #Free, , sFile
 
 If sFile.Version = Hex(100) Then
 
 'Put any code you would want here
 Me.Caption = "File Binder - " + cd.Filename
 
 Else
 
 MsgBox ("This file is corrupt or made by a newer or older version of file bind!")
 ReDim sFile.FileContent(0)
 ReDim sFile.Filename(0)
 sFile.Version = ""
 
 Close #Free
 
 Me.MousePointer = vbArrow
 
 Exit Sub
 
 End If


ListFile

 Me.MousePointer = vbArrow

End If

End Sub

Private Sub mnuRemove_Click()

Dim Result As Integer

If lstFile.Text <> "" Then
Result = MsgBox("Are you sure you want to remove " + lstFile.Text + "?", vbYesNo + vbQuestion, "Remove")

If Result = vbYes Then

Me.MousePointer = vbHourglass

DeleteValue lstFile.ListIndex + 1, sFile.Filename
DeleteValue lstFile.ListIndex + 1, sFile.FileContent

ReDim Preserve sFile.Filename(UBound(sFile.Filename, 1) - 1)
ReDim Preserve sFile.FileContent(UBound(sFile.FileContent, 1) - 1)

ListFile

End If

Me.MousePointer = vbArrow

End If

End Sub

Private Function ListFile()

lstFile.Clear

For i = 1 To UBound(sFile.Filename)
lstFile.AddItem "-" + sFile.Filename(i)
Next i

End Function

Private Sub mnuSave_Click()

On Error Resume Next

Dim Free As Long
Free = FreeFile

cd.Filter = "File Package (*.pck)|*.pck"
cd.DialogTitle = "Save as..."
cd.ShowSave

If cd.Filename <> "" Then

 Me.MousePointer = vbHourglass

Kill cd.Filename

If LCase(Right(cd.Filename, 4)) <> ".pck" Then cd.Filename = cd.Filename + ".pck"

Open cd.Filename For Binary As #Free
 Put #Free, , sFile
Close #Free

 Me.Caption = "File Binder - " + cd.Filename

 Me.MousePointer = vbArrow

End If

End Sub
