VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HTML Creator"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About..."
      Height          =   375
      Left            =   3240
      TabIndex        =   26
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdOptions 
      Caption         =   "&Options"
      Height          =   375
      Left            =   4560
      TabIndex        =   25
      Top             =   7080
      Width           =   1215
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   3000
      TabIndex        =   3
      Text            =   "Untitled"
      Top             =   360
      Width           =   2775
   End
   Begin VB.CheckBox chkNoBreakReturn 
      Caption         =   "Use &no break after a file"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Frame Frame 
      Caption         =   "Advanced Options"
      Height          =   1335
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   5400
      Width           =   5655
      Begin VB.CheckBox chkTotalImage 
         Caption         =   "&Show total file(s)"
         Height          =   195
         Left            =   2880
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkThumbDirectory 
         Caption         =   "'&Thumb' directory exist"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cmbThumb 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label txtThumbnails 
         AutoSize        =   -1  'True
         Caption         =   "&Create Thumbnails for images:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   2130
      End
      Begin VB.Label lblPixels 
         AutoSize        =   -1  'True
         Caption         =   "pixels"
         Height          =   195
         Left            =   1800
         TabIndex        =   20
         Top             =   600
         Width           =   390
      End
   End
   Begin VB.Frame Frame 
      Height          =   25
      Index           =   1
      Left            =   -720
      TabIndex        =   23
      Top             =   6840
      Width           =   8775
   End
   Begin VB.TextBox txtOut 
      Height          =   615
      Left            =   2880
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   22
      Top             =   5880
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdWrite 
      Caption         =   "&Write"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CheckBox chkShowImages 
      Caption         =   "Show images on &page (supported format are: JPG, BMP, GIF && PNG)"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5040
      Width           =   5535
   End
   Begin VB.CheckBox chkOpenNewWindow 
      Caption         =   "Open links in &new window"
      Height          =   195
      Left            =   3000
      TabIndex        =   13
      Top             =   4560
      Width           =   2775
   End
   Begin VB.CheckBox chkFileAsLink 
      Caption         =   "Treat every file as &Link"
      Height          =   195
      Left            =   240
      TabIndex        =   12
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Frame Frame 
      Caption         =   "Files"
      Height          =   3735
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5655
      Begin VB.CommandButton cmdFav 
         Caption         =   "&Favourites"
         Height          =   375
         Left            =   4320
         TabIndex        =   11
         Top             =   3270
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   2940
         Width           =   5415
      End
      Begin VB.FileListBox File 
         Enabled         =   0   'False
         Height          =   480
         Left            =   4440
         TabIndex        =   7
         Top             =   1080
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.DirListBox Dir 
         Height          =   2115
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   5415
      End
      Begin VB.DriveListBox Drive 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   5415
      End
      Begin VB.Label lblFoundImages 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2760
         TabIndex        =   10
         Top             =   3360
         Width           =   1425
      End
      Begin VB.Label lblImageFound 
         AutoSize        =   -1  'True
         Caption         =   "Files(s) found in current directory:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   3360
         Width           =   2325
      End
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Untitled"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label lblFileName 
      AutoSize        =   -1  'True
      Caption         =   "&File Name:"
      Height          =   195
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   750
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      Caption         =   "Page &Title:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   765
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "mnuPopUp"
      Visible         =   0   'False
      Begin VB.Menu mnuPopUpOpenFile 
         Caption         =   "&Open recently created HTML file..."
      End
      Begin VB.Menu mnuPopUpEditFile 
         Caption         =   "&Edit using default handler..."
      End
      Begin VB.Menu z0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpEditFileNotepad 
         Caption         =   "Edit using &Notepad..."
      End
      Begin VB.Menu z1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopUpBrowseCurrentDir 
         Caption         =   "&Explore current directory"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkThumbDirectory_Click()
    If chkThumbDirectory.Value = 1 Then
        cmbThumb.ListIndex = 0
        cmbThumb.Enabled = False
    Else
        cmbThumb.Enabled = True
    End If
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub cmdFav_Click()
    frmFav.Show
End Sub

Private Sub cmdOptions_Click()
'
End Sub

Private Sub cmdOptions_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        PopupMenu mnuPopUp
    End If
End Sub

Private Sub cmdWrite_Click()
Dim htmlOutputFile As String
Dim htmlFile As String
Dim htmlOpenNewWindow As String
Dim htmlShowImages As String
Dim htmlImageWidth As String
Dim htmlCarriageReturn As String
Dim htmlThumdDirExist As String
Dim htmlTotalImage As String
Dim htmlDocType As String
Dim htmlAppCreate As String

Dim nFileNum As Integer

' Clear txtOut
txtOut.Text = ""
File.Refresh

' Set variable of output file
If (Right$(Dir.Path, 1) = "\") Then
    htmlOutputFile = Dir.Path & txtFileName.Text & ".html"
Else
    htmlOutputFile = Dir.Path & "\" & txtFileName.Text & ".html"
End If

cmdWrite.Enabled = False
txtName.Enabled = False
txtFileName.Enabled = False
Drive.Enabled = False
chkFileAsLink.Enabled = False
chkOpenNewWindow.Enabled = False
chkNoBreakReturn.Enabled = False
chkShowImages.Enabled = False
chkThumbDirectory.Enabled = False
chkTotalImage.Enabled = False

If chkTotalImage.Value = 1 Then
    htmlTotalImage = "      <title>" & txtName.Text & " | " & (File.ListCount) & " file(s) total" & "</title>"
Else
    htmlTotalImage = "      <title>" & txtName.Text & "</title>"
End If

htmlAppCreate = "<!-- Created using HTML Editor (version " & App.Major & "." & App.Minor & " BUILD " & App.Revision & ") on " & _
                Date$ & " " & Time$ & " by " & Environ("UserName") & " -->"
                
htmlDocType = "<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">" & vbCrLf & _
"<html xmlns=""http://www.w3.org/1999/xhtml"" xml:lang=""en"">" & vbCrLf & _
"   " & htmlAppCreate & vbCrLf & _
"   <head>" & vbCrLf & _
"      <style type=""text/css"">" & vbCrLf & _
"         a {text-decoration: none; font-family: Tahoma, sans-serif;}" & vbCrLf & _
"         a:hover {text-decoration: underline; color: red;}" & vbCrLf & _
"      </style>" & vbCrLf & _
htmlTotalImage & vbCrLf & _
"   </head>"

' Writing the actual HTML code in txtOut
txtOut.SelText = htmlDocType & vbCrLf & _
"   <body>" & vbCrLf

For X = 0 To (File.ListCount - 1)
    File.ListIndex = X
    
    If (Right$(Dir.Path, 1) = "\") Then
        htmlFile = File.FileName
    Else
        htmlFile = File.FileName
    End If
    
    If chkOpenNewWindow.Value = 1 Then
        htmlOpenNewWindow = " target=""_blank"""
    End If
    
    If chkNoBreakReturn.Value = 1 Then
        htmlCarriageReturn = ""
    Else
        htmlCarriageReturn = "<br>"
    End If
    
    If chkShowImages.Value = 1 Then
    
        If cmbThumb.ListIndex = 0 Then
            htmlImageWidth = ""
        Else
            htmlImageWidth = " width=" & cmbThumb.Text
        End If
        
        If chkThumbDirectory.Value = 1 Then
            htmlThumdDirExist = "thumb/"
        Else
            htmlThumdDirExist = ""
        End If
                
        Select Case LCase$(Right$(File.FileName, 3))
            Case "jpg"
                htmlShowImages = "<img src=" & """" & htmlThumdDirExist & htmlFile & """" & htmlImageWidth & "/>"
            Case "gif"
                htmlShowImages = "<img src=" & """" & htmlThumdDirExist & htmlFile & """" & htmlImageWidth & "/>"
            Case "png"
                htmlShowImages = "<img src=" & """" & htmlThumdDirExist & htmlFile & """" & htmlImageWidth & "/>"
            Case "bmp"
                htmlShowImages = "<img src=" & """" & htmlThumdDirExist & htmlFile & """" & htmlImageWidth & "/>"
            Case "peg"
                htmlShowImages = "<img src=" & """" & htmlThumdDirExist & htmlFile & """" & htmlImageWidth & "/>"
            Case Else
                htmlShowImages = File.FileName
        End Select
    End If
    
    If chkFileAsLink.Value = 1 Then
        If chkShowImages.Value = 1 Then
            txtOut.SelText = "         <a href=" & """" & htmlFile & """" & htmlOpenNewWindow & ">" & htmlShowImages & "</a>" & htmlCarriageReturn & vbCrLf
        Else
            txtOut.SelText = "         <a href=" & """" & htmlFile & """" & htmlOpenNewWindow & ">" & htmlFile & "</a>" & htmlCarriageReturn & vbCrLf
        End If
        
    ElseIf chkShowImages.Value = 1 Then
        txtOut.SelText = htmlShowImages & htmlCarriageReturn & vbCrLf
    Else
        txtOut.SelText = htmlFile & htmlCarriageReturn & vbCrLf
    End If
    
    'txtOut.SelText = htmlFile & "<br>" & vbCrLf
    DoEvents
Next X

txtOut.SelText = "   </body>" & vbCrLf & "</html>"

' Get a free file number
nFileNum = FreeFile

' Create the Output file
Open htmlOutputFile For Output As nFileNum

' Write the contents of txtOut to Output file
Print #nFileNum, txtOut.Text

' Close the file
Close nFileNum

cmdWrite.Enabled = True
txtName.Enabled = True
txtFileName.Enabled = True
Drive.Enabled = True
chkFileAsLink.Enabled = True
chkOpenNewWindow.Enabled = True
chkNoBreakReturn.Enabled = True
chkShowImages.Enabled = True
chkThumbDirectory.Enabled = True
chkTotalImage.Enabled = True

LastCreatedHTMLFile = htmlOutputFile

End Sub

Private Sub Dir_Change()
    File.Path = Dir.Path
    txtPath.Text = Dir.Path
    lblFoundImages.Caption = File.ListCount
End Sub

Private Sub Drive_Change()
On Error Resume Next
    Dir.Path = Drive.Drive
End Sub

Private Sub Form_Load()
    DesktopPath = Environ("USERPROFILE") & "\Desktop\"
    
cmbThumb.AddItem "(None)"
cmbThumb.AddItem "100"
cmbThumb.AddItem "200"
cmbThumb.AddItem "300"
cmbThumb.AddItem "400"
cmbThumb.AddItem "500"
cmbThumb.ListIndex = 0

txtPath.Text = Dir.Path
lblFoundImages.Caption = File.ListCount
End Sub

Private Sub mnuPopUpBrowseCurrentDir_Click()
Dim ExplorerPath As String

ExplorerPath = "explorer.exe /e," & txtPath.Text
    
    Shell ExplorerPath, vbNormalFocus
End Sub

Private Sub mnuPopUpEditFile_Click()
If LastCreatedHTMLFile = "" Then
    MsgBox "Unable to open last created HTML file.", vbExclamation, "Error"
Else
    ShellEx LastCreatedHTMLFile, essSW_SHOWNORMAL, , "c:\", "edit", Me.hWnd
End If
End Sub

Private Sub mnuPopUpEditFileNotepad_Click()
If LastCreatedHTMLFile = "" Then
    MsgBox "Unable to open last created HTML file.", vbExclamation, "Error"
Else
    Shell "notepad " & LastCreatedHTMLFile, vbNormalFocus
End If
End Sub

Private Sub mnuPopUpOpenFile_Click()
If LastCreatedHTMLFile = "" Then
    MsgBox "Unable to open last created HTML file.", vbExclamation, "Error"
Else
    ShellEx LastCreatedHTMLFile, essSW_SHOWNORMAL, , "c:\", , Me.hWnd
End If
End Sub
