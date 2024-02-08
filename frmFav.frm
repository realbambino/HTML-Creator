VERSION 5.00
Begin VB.Form frmFav 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Favourites"
   ClientHeight    =   3450
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
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
   ScaleHeight     =   3450
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lstFav 
      Height          =   2595
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4455
   End
   Begin VB.Frame Frame 
      Height          =   25
      Left            =   -480
      TabIndex        =   1
      Top             =   600
      Width           =   7335
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add current directory to favourites..."
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
