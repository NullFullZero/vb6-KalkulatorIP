VERSION 5.00
Begin VB.Form frmKalkulatorIP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kalkulator IP"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7725
   Icon            =   "frmKalkulatorIP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Wynik:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   7455
      Begin VB.Frame Frame4 
         Caption         =   "Adresacja:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         TabIndex        =   31
         Top             =   1680
         Width           =   7215
         Begin VB.Label lblTyp 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   375
            Left            =   120
            TabIndex        =   40
            Top             =   2520
            Width           =   6975
         End
         Begin VB.Label lblBroadCast 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   39
            Top             =   720
            Width           =   4935
         End
         Begin VB.Label lblEtykieta 
            BackStyle       =   0  'Transparent
            Caption         =   "Adres rozg³.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   240
            TabIndex        =   38
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblEtykieta 
            BackStyle       =   0  'Transparent
            Caption         =   "Adres sieci:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   240
            TabIndex        =   37
            Top             =   360
            Width           =   1815
         End
         Begin VB.Label lblAdresSieci 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   36
            Top             =   360
            Width           =   4935
         End
         Begin VB.Label lblOstatniAdres 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   35
            Top             =   1440
            Width           =   4935
         End
         Begin VB.Label lblEtykieta 
            BackStyle       =   0  'Transparent
            Caption         =   "Ostatni adres:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   240
            TabIndex        =   34
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Label lblPierwszyAdres 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2160
            TabIndex        =   33
            Top             =   1080
            Width           =   4935
         End
         Begin VB.Label lblEtykieta 
            BackStyle       =   0  'Transparent
            Caption         =   "Pierwszy adres:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   238
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   32
            Top             =   1080
            Width           =   1815
         End
      End
      Begin VB.Label lblLiczbaAdresow 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   30
         Top             =   1320
         Width           =   4935
      End
      Begin VB.Label lblEtykieta 
         BackStyle       =   0  'Transparent
         Caption         =   "Liczba adresów:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   29
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label lblMaskaSieci 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   27
         Top             =   840
         Width           =   4935
      End
      Begin VB.Label lblIP 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblEtykieta 
         BackStyle       =   0  'Transparent
         Caption         =   "Maska(2):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   23
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lblEtykieta 
         BackStyle       =   0  'Transparent
         Caption         =   "IP(2):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   22
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Funckje:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   5520
      TabIndex        =   19
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton cmdKoniec 
         Caption         =   "Zamknij"
         Height          =   375
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrzelicz 
         Caption         =   "Przelicz"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   840
         Picture         =   "frmKalkulatorIP.frx":0442
         Top             =   840
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Wartoœci:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox cmbCIDR 
         Height          =   315
         ItemData        =   "frmKalkulatorIP.frx":0884
         Left            =   1440
         List            =   "frmKalkulatorIP.frx":08EB
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1440
         Width           =   3255
      End
      Begin VB.TextBox MM 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox MM 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   3
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox MM 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   6
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox MM 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox MIP 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   1
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox MIP 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   2
         Left            =   2280
         MaxLength       =   3
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox MIP 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   3
         Left            =   3120
         MaxLength       =   3
         TabIndex        =   3
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox MIP 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   4
         Left            =   3960
         MaxLength       =   3
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Wybór:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Maska:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Adres IP:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lblDot 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   5
         Left            =   3840
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblDot 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   3000
         TabIndex        =   15
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblDot 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   3
         Left            =   2160
         TabIndex        =   14
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lblDot 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   2160
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblDot 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   1
         Left            =   3000
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.Label lblDot 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   2
         Left            =   3840
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox MM 
      Height          =   375
      Index           =   0
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   9
      Top             =   6600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox MIP 
      Alignment       =   1  'Right Justify
      Height          =   375
      Index           =   0
      Left            =   7200
      MaxLength       =   3
      TabIndex        =   0
      Top             =   6240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Kalkulator IP v.1.0 NullFullZero, NullFullZero@gmail.com"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   7440
      Width           =   7335
   End
End
Attribute VB_Name = "frmKalkulatorIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Aliquet 1.0
'Autor: Adam Shaman Czana
'Nazwa: Kalkulator IP
Option Explicit

Private Sub cmbCIDR_Click()
ZCIDRMASKA = MaskaZCIDR(cmbCIDR)

'Call MsgBox(Mid(ZCIDRMASKA, 1, 8), vbOKOnly)
'Call MsgBox(Mid(ZCIDRMASKA, 9, 8), vbOKOnly)
'Call MsgBox(Mid(ZCIDRMASKA, 17, 8), vbOKOnly)
'Call MsgBox(Mid(ZCIDRMASKA, 25, 8), vbOKOnly)

MM(1) = Bin2Dec(Mid(ZCIDRMASKA, 1, 8))
MM(2) = Bin2Dec(Mid(ZCIDRMASKA, 9, 8))
MM(3) = Bin2Dec(Mid(ZCIDRMASKA, 17, 8))
MM(4) = Bin2Dec(Mid(ZCIDRMASKA, 25, 8))

'MaskaSieci(1) = DzNaBin(MM(1), 8)
'MaskaSieci(2) = DzNaBin(MM(2), 8)
'MaskaSieci(3) = DzNaBin(MM(3), 8)
'MaskaSieci(4) = DzNaBin(MM(4), 8)
End Sub


Private Sub cmdKoniec_Click()
End
End Sub

Private Sub cmdPrzelicz_Click()
On Error Resume Next
Dim i As Byte

'zamiana
lblLiczbaAdresow.Caption = "..."
lblTyp.Caption = ""
For i = 1 To 4
    
    IPAdres(i) = DzNaBin(MIP(i), 8)
    MaskaSieci(i) = DzNaBin(MM(i), 8)
    

Next i



lblIP.Caption = IPAdres(1) & " . " & IPAdres(2) & " . " & IPAdres(3) & " . " & IPAdres(4)
lblMaskaSieci.Caption = MaskaSieci(1) & " . " & MaskaSieci(2) & " . " & MaskaSieci(3) & " . " & MaskaSieci(4)

cmbCIDR = LiczJedynki(lblMaskaSieci.Caption)
For i = 1 To 8
    ZanegowanaMaskaSieci(1) = ZanegowanaMaskaSieci(1) & nNeg(Mid(MaskaSieci(1), i, 1))
    ZanegowanaMaskaSieci(2) = ZanegowanaMaskaSieci(2) & nNeg(Mid(MaskaSieci(2), i, 1))
    ZanegowanaMaskaSieci(3) = ZanegowanaMaskaSieci(3) & nNeg(Mid(MaskaSieci(3), i, 1))
    ZanegowanaMaskaSieci(4) = ZanegowanaMaskaSieci(4) & nNeg(Mid(MaskaSieci(4), i, 1))
Next i
'Tester.Text = ZanegowanaMaskaSieci(4)



Dim LAU As Long
LAU = LiczAdrUzytkowe(cmbCIDR)

lblLiczbaAdresow.Caption = LiczAdr(cmbCIDR) & "         u¿ytkowych: " & LAU

If LAU = "0" Or LAU = "-1" Then
    lblLiczbaAdresow.Caption = LiczAdr(cmbCIDR) & "         u¿ytkowych: brak"
End If

'lblLiczbaAdresow.Caption = LiczAdr(LiczJedynki(lblMaskaSieci.Caption))
Dim Adr As String

Adr = Bin2Dec(CalyAND(IPAdres(1), MaskaSieci(1)))
Adr = Adr & " . "
Adr = Adr & Bin2Dec(CalyAND(IPAdres(2), MaskaSieci(2)))
Adr = Adr & " . "
Adr = Adr & Bin2Dec(CalyAND(IPAdres(3), MaskaSieci(3)))
Adr = Adr & " . "
Adr = Adr & Bin2Dec(CalyAND(IPAdres(4), MaskaSieci(4)))

'IPAdresSieci(5)
IPAdresSieci(1) = Bin2Dec(CalyAND(IPAdres(1), MaskaSieci(1)))
IPAdresSieci(2) = Bin2Dec(CalyAND(IPAdres(2), MaskaSieci(2)))
IPAdresSieci(3) = Bin2Dec(CalyAND(IPAdres(3), MaskaSieci(3)))
IPAdresSieci(4) = Bin2Dec(CalyAND(IPAdres(4), MaskaSieci(4)))

lblAdresSieci.Caption = Adr

AB(1) = ""
AB(2) = ""
AB(3) = ""
AB(4) = ""

For i = 1 To 8
    NAS(1) = DzNaBin(IPAdresSieci(1), 8)
    NAS(2) = DzNaBin(IPAdresSieci(2), 8)
    NAS(3) = DzNaBin(IPAdresSieci(3), 8)
    NAS(4) = DzNaBin(IPAdresSieci(4), 8)
    
    AB(1) = AB(1) & oOR(Mid(NAS(1), i, 1), Mid(ZanegowanaMaskaSieci(1), i, 1))
    AB(2) = AB(2) & oOR(Mid(NAS(2), i, 1), Mid(ZanegowanaMaskaSieci(2), i, 1))
    AB(3) = AB(3) & oOR(Mid(NAS(3), i, 1), Mid(ZanegowanaMaskaSieci(3), i, 1))
    AB(4) = AB(4) & oOR(Mid(NAS(4), i, 1), Mid(ZanegowanaMaskaSieci(4), i, 1))

Next i

'Tester = Tester & "AS: " & IPAdresSieci(1) & " . " & IPAdresSieci(2) & " . " & IPAdresSieci(3) & " . " & IPAdresSieci(4) & vbCrLf '& " . "
lblBroadCast.Caption = "..."
lblBroadCast.Caption = Bin2Dec(AB(1)) & " . " & Bin2Dec(AB(2)) & " . " & Bin2Dec(AB(3)) & " . " & Bin2Dec(AB(4))

'lblPierwszyAdres.Caption = IPAdresSieci(1) & " . " & IPAdresSieci(2) & " . " & IPAdresSieci(3) & " . " & IPAdresSieci(3) + 1
lblOstatniAdres.Caption = Bin2Dec(AB(1)) & " . " & Bin2Dec(AB(2)) & " . " & Bin2Dec(AB(3)) & " . " & Bin2Dec(AB(4) - 1)


If MIP(1) & MIP(2) = "192168" Then
    lblTyp = "SIEÆ PRYWATNA 192.168.0.0"
ElseIf MIP(1) = "172" Then
    
    If (MIP(2) >= 16) And (MIP(2) <= 31) Then  'And Val(MIP(2).Text) <= 31
        lblTyp = "SIEÆ PRYWATNA 172.16.0.0 - 172.31.0.0"
    End If
ElseIf MIP(1) = "10" Then
    lblTyp = "SIEÆ PRYWATNA 10.0.0.0"
End If

Adr = ""
AB(1) = ""
AB(2) = ""
AB(3) = ""
AB(4) = ""
IPAdresSieci(1) = ""
IPAdresSieci(2) = ""
IPAdresSieci(3) = ""
IPAdresSieci(4) = ""
NAS(1) = ""
NAS(2) = ""
NAS(3) = ""
NAS(4) = ""
LAU = 0
i = 0
ZanegowanaMaskaSieci(1) = ""
ZanegowanaMaskaSieci(2) = ""
ZanegowanaMaskaSieci(3) = ""
ZanegowanaMaskaSieci(4) = ""

'Call PierwszyAdres(lblIP.Caption, lblMaskaSieci.Caption)
lblPierwszyAdres.Caption = Replace(PierwszyAdres(lblIP.Caption, lblMaskaSieci.Caption, True), ".", " . ")
End Sub

Private Sub Command1_Click()
lblBroadCast.Caption = ""
End Sub

Private Sub Form_Load()
On Error Resume Next
MIP(1).Text = "192"
MIP(2).Text = "168"
MIP(3).Text = "0"
MIP(4).Text = "0"

MM(1).Text = "255"
MM(2).Text = "255"
MM(3).Text = "255"
MM(4).Text = "0"

End Sub

Private Sub MIP_Change(Index As Integer)
On Error Resume Next
If Len(MIP(Index)) = 3 Then
    If Index = 4 Then
        MM(1).SetFocus
    Else
        MIP(Index + 1).SetFocus
    End If
End If
End Sub

Private Sub MM_Change(Index As Integer)
On Error Resume Next
If Len(MM(Index)) = 3 Then
    If Index = 4 Then
        MIP(1).SetFocus
    Else
        MM(Index + 1).SetFocus
    End If
End If
End Sub
