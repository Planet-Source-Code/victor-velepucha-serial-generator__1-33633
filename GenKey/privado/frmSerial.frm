VERSION 5.00
Begin VB.Form frmSerial 
   Caption         =   "Serial Generator"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmSerial.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Serial number: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   4455
      Begin VB.TextBox txtSerial 
         BackColor       =   &H8000000F&
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
         Left            =   600
         TabIndex        =   9
         Top             =   360
         Width           =   3330
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1080
         Width           =   2595
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Show serial"
         Default         =   -1  'True
         Height          =   375
         Left            =   1680
         TabIndex        =   7
         Top             =   1440
         Width           =   2580
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   4
         Top             =   720
         Width           =   2595
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   0
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   2
         Top             =   360
         Width           =   2595
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "&Email: "
         Height          =   195
         Left            =   1200
         TabIndex        =   5
         Top             =   1125
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "User name"
         Height          =   195
         Left            =   840
         TabIndex        =   3
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Program Code"
         Height          =   195
         Left            =   600
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmSerial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGenerar_Click()
    Dim s As New Serial
    txtSerial = s.GenerarSerial(txtData(0), txtData(1), txtData(2))
    frmProbarSerial.txtData(0).Text = txtData(0)
    frmProbarSerial.txtData(1).Text = txtData(1)
    frmProbarSerial.txtData(2).Text = txtData(2)
    frmProbarSerial.txtData(3).Text = txtSerial
End Sub

Private Function validarDatos() As Boolean
    If InStr(txtData(2).Text, "@") = 0 Then
        MsgBox "Incorrect email", vbInformation, "Alerta"
        validarDatos = False
        Exit Function
    End If
    validarDatos = True
End Function

Private Sub Form_Load()
    Left = 500
    Top = 500
    frmProbarSerial.Left = 480 + Me.Width
    frmProbarSerial.Top = Top
    frmProbarSerial.Show
End Sub


Private Sub Form_Terminate()
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub
