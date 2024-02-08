VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FBFBFB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About..."
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label lblAllRightsReserved 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All rights reserved."
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   960
      Width           =   1290
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2007-2008 Steve Jobs."
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.Label lblAppName 
      BackStyle       =   0  'Transparent
      Caption         =   "HTML Creator"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Image Image 
      Height          =   960
      Left            =   240
      Picture         =   "frmAbout.frx":0000
      Top             =   240
      Width           =   960
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

