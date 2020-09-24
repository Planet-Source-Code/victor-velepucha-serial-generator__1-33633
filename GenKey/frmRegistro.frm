VERSION 5.00
Begin VB.Form frmRegistro 
   Caption         =   "Registering the program"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4725
   Icon            =   "frmRegistro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4455
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
         TabIndex        =   8
         Top             =   1320
         Width           =   2595
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   2
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   6
         Top             =   960
         Width           =   2595
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   2595
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Index           =   1
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   2
         Text            =   "WRITE YOUR NAME"
         Top             =   600
         Width           =   2595
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Serial number"
         Height          =   195
         Left            =   600
         TabIndex        =   9
         Top             =   1350
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "&Email(optional): "
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   1005
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Program Code:"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "User name"
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   600
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   360
      Left            =   1680
      TabIndex        =   0
      Top             =   2760
      Width           =   1560
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      Height          =   735
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "frmRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdRegistrar_Click()
    Dim Filenr As Integer
    Dim crypt As New clsCryptAPI
    Dim v As Collection
    Dim s As New serial
    Dim sEnc As String
    Dim DestFile As String
    Dim clave As String
    Set v = s.VerificarSerial(txtData(0), txtData(1), txtData(2), txtData(3))
    If v("ValidKey") Then
        sEnc = crypt.EncryptString(txtData(3), "QualityLider")
        DestFile = App.Path & "/siprec.lic"
        If (FileExist(DestFile)) Then Kill DestFile
        Filenr = FreeFile
        Open DestFile For Binary As #Filenr
        Put #Filenr, , sEnc
        Close #Filenr
        MsgBox "Serial number aproved" & vbCrLf & vbCrLf & "Thaks for purchase this program", vbInformation, "Register OK!!"
        Unload Me
        frmBdd.Show
    Else
        MsgBox "Incorrect serial number", vbExclamation, "Error!!"
    End If
    Set s = Nothing
    Set v = Nothing
    Set crypt = Nothing

End Sub

Private Sub Form_Load()
    Dim v As Collection
    Dim s As New serial
    Dim snhd As New clsHardDisk
    Dim Filenr As Integer
    Dim crypt As New clsCryptAPI
    Dim sEnc As String, sDec As String
    Dim filLic As String
    Dim ByteArray() As Byte
    Dim tmp As String
    filLic = App.Path & "/siprec.lic"
    tmp = snhd.getSerialNumber
    If (FileExist(filLic)) Then
        Filenr = FreeFile
        Open filLic For Binary As #Filenr
        ReDim ByteArray(0 To LOF(Filenr) - 1)
        Get #Filenr, , ByteArray()
        Close #Filenr
        sEnc = StrConv(ByteArray(), vbUnicode)
        sDec = crypt.DecryptString(sEnc, "QualityLider")
        Set v = s.VerificarLicencia(tmp, sDec)
        If v("ValidKey") Then
            Unload Me
            frmBdd.Show
            GoTo fin
        End If
    End If
    txtData(0).Text = tmp
    Label1.Caption = "You need a licence for use this program" & vbCrLf & _
                     "Get one :seller@email.com" & vbCrLf & _
                     "Telephone: 1800 900 900"
fin:
    Set crypt = Nothing
    Set c = Nothing
    Set snhd = Nothing
End Sub
Public Function FileExist(Filename As String) As Boolean
On Error GoTo NotExist
    Call FileLen(Filename)
    FileExist = True
    Exit Function
NotExist:
    FileExist = False
End Function

Private Sub Form_Terminate()
    End
End Sub
