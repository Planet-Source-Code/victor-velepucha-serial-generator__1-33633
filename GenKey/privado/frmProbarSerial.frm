VERSION 5.00
Begin VB.Form frmProbarSerial 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Try serial"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "try serial"
      Default         =   -1  'True
      Height          =   360
      Left            =   1560
      TabIndex        =   9
      Top             =   1920
      Width           =   1560
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Top             =   600
         Width           =   2595
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   0
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   2595
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   2
         Top             =   960
         Width           =   2595
      End
      Begin VB.TextBox txtData 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   1
         Top             =   1320
         Width           =   2595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User name"
         Height          =   195
         Left            =   840
         TabIndex        =   8
         Top             =   600
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Program code:"
         Height          =   195
         Left            =   600
         TabIndex        =   7
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Email: "
         Height          =   195
         Left            =   1200
         TabIndex        =   6
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serial number"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   1350
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmProbarSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegistrar_Click()
    Dim v As Collection
    Dim s As New Serial
    Set v = s.VerificarSerial(txtData(0), txtData(1), txtData(2), txtData(3))
    If v("ValidKey") Then
        MsgBox "Serial OK"
    Else
        MsgBox "Incorrect serial number", vbExclamation, "Error!!"
    End If
    Set s = Nothing
    Set v = Nothing
    Set crypt = Nothing
End Sub

Private Sub Form_Terminate()
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
