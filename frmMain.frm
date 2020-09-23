VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "App.Path Replacement"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtProg 
      ForeColor       =   &H00008000&
      Height          =   300
      Left            =   1065
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   480
      Width           =   3450
   End
   Begin VB.TextBox txtApp 
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   1065
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3450
   End
   Begin VB.Label Label3 
      Caption         =   "Copy the compiled program to a CD or a Network Drive and run it... You'll see what the matter could be."
      Height          =   480
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   4425
   End
   Begin VB.Label Label2 
      Caption         =   "ProgPath ="
      Height          =   270
      Left            =   180
      TabIndex        =   3
      Top             =   525
      Width           =   825
   End
   Begin VB.Label Label1 
      Caption         =   "App.Path ="
      Height          =   270
      Left            =   180
      TabIndex        =   2
      Top             =   165
      Width           =   825
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
GetProgPath

If Right(App.Path, 1) <> "\" Then
    txtApp.Text = App.Path + "\"
Else
    txtApp.Text = App.Path
End If

txtProg.Text = ProgPath

End Sub
