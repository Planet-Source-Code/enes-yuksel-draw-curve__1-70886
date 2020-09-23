VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Serbest Çizim"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6255
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "?"
      Height          =   255
      Left            =   4080
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   3840
      Width           =   255
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Check6"
      Height          =   255
      Left            =   8280
      TabIndex        =   26
      Top             =   3840
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0CCA
      Left            =   3840
      List            =   "Form1.frx":0CE3
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   3420
      Width           =   495
   End
   Begin VB.Frame Frame3 
      Caption         =   " Çizgi Rengi "
      Height          =   855
      Left            =   4560
      TabIndex        =   16
      Top             =   3280
      Width           =   1575
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   7
         Left            =   460
         TabIndex        =   24
         Top             =   240
         Width           =   320
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   800
         TabIndex        =   23
         Top             =   240
         Width           =   320
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   1140
         TabIndex        =   22
         Top             =   240
         Width           =   320
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF00FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   320
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   460
         TabIndex        =   20
         Top             =   480
         Width           =   320
      End
      Begin VB.Label Label4 
         BackColor       =   &H000080FF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   800
         TabIndex        =   19
         Top             =   480
         Width           =   320
      End
      Begin VB.Label Label4 
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   18
         Top             =   480
         Width           =   320
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   320
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Klavuz "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4560
      TabIndex        =   12
      Top             =   2160
      Width           =   1575
      Begin VB.CheckBox Check2 
         Caption         =   "Noktalar Yok"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   240
         Width           =   1300
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Çizgiler Yok"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   1300
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Sayý Yok"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   480
         Width           =   1300
      End
   End
   Begin VB.PictureBox TMPres 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   8280
      ScaleHeight     =   2955
      ScaleWidth      =   3315
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Check5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8280
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Onay"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4560
      TabIndex        =   9
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   " Çizgi Biçimi "
      Height          =   855
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   1575
      Begin VB.OptionButton Option2 
         Caption         =   "Eðri Hatlar"
         Height          =   200
         Left            =   180
         TabIndex        =   6
         Top             =   280
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Keskin Hatlar"
         Height          =   200
         Left            =   180
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Nokta Ekleme"
      Height          =   195
      Left            =   4560
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "&Temizle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   340
      Left            =   4560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.PictureBox RES1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   3195
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Son Çalýþýlan Nokta: 0"
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   4215
   End
   Begin VB.Label Label2 
      Caption         =   "Toplam Nokta: 0"
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   120
      TabIndex        =   7
      Top             =   3960
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Yeni nokta eklemek için görüntüye týklayýn!"
      ForeColor       =   &H00008000&
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   4215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===============================================================================================
'Bu kodu 2002 de yazmýþým (hatýrlamýyorum.)
'Bu kodun içindeki bazý öðeleri planet-source 'den almýþtým.
'Aldýðým kaynaðý hatýrlamadýðým için belirtemiyorum.
'Bu çalýþma, bir çizim programýnda modül olarak kullanýldý.
'Bu çalýþmayý, modülün iþ görüp görmeyeceðini göstermek için hazýrlamýþtým.
'Kaynaklýk ettiði modül çok daha fazla özellik/güzellik içeriyor. Ancak:
'Ticari bir projede yeraldýðýndan kaynaðý veremedim.
'Kolay gelsin...
'===============================================================================================
Dim Renk
Dim nc As Integer
Dim Cont(100, 1) As Integer
Dim NewLocPoint As Integer
Const Smooth = 0.02
Dim Dragging As Boolean
Function B(k, n, u)
B = C(n, k) * (u ^ k) * (1 - u) ^ (n - k)
End Function
Function C(n, r)
C = fact(n) / (fact(r) * fact(n - r))
End Function
Function fact(n)
If n = 1 Or n = 0 Then
   fact = 1
Else
   fact = n * fact(n - 1)
End If
End Function
Private Sub AddCont(X, Y)
Cont(nc, 0) = X: Cont(nc, 1) = Y
nc = nc + 1
Label2.Caption = "Toplam Nokta: " & nc
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Label1.Caption = "Noktalarý taþýmak için klavuzlarý sürükleyin..."
Else
   Label1.Caption = "Yeni bir nokta eklemek için görüntüye týklayýn!"
End If
YenidenCIZ
End Sub
Private Sub Check2_Click()
If Check2.Value = 1 Then Check4.Value = 1
YenidenCIZ
End Sub
Private Sub Check3_Click()
YenidenCIZ
End Sub
Private Sub Check4_Click()
If Check2.Value = 1 Then Check4.Value = 1
YenidenCIZ
End Sub

Private Sub cmdReset_Click()
nc = 0
RES1.Cls
End Sub
Private Sub Combo1_Change()
RES1.DrawWidth = Combo1.Text
If Combo1.Text > 1 Then If Check3.Enabled = True Then Check3.Value = 1
YenidenCIZ
End Sub

Private Sub Combo1_Click()
RES1.DrawWidth = Combo1.Text
If Combo1.Text > 1 Then If Check3.Enabled = True Then Check3.Value = 1
YenidenCIZ
End Sub

Private Sub Command1_Click()
If Option2.Value = True Then
   Check5.Value = 1
Else
   Check6.Value = 1
End If
YenidenCIZ
TMPres.Picture = RES1.Image
nc = 0
Check5.Value = 0
Check6.Value = 0
RES1.Cls
RES1.Picture = TMPres.Image
End Sub

Private Sub Command2_Click()
MsgBox "Enes Deniz YUKSEL, 2002", vbInformation
End Sub

Private Sub Form_Load()
Renk = 0
Combo1.ListIndex = 0
Form1.ScaleMode = vbTwips
RES1.ScaleMode = vbPixels
RES1.FontSize = 7
End Sub
Private Sub Label4_Click(Index As Integer)
Renk = Label4(Index).BackColor
YenidenCIZ
End Sub
Private Sub Option1_Click()
Check3.Value = 0
Check3.Enabled = False
YenidenCIZ
End Sub
Private Sub Option2_Click()
Check3.Enabled = True
YenidenCIZ
End Sub
Private Sub Res1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo Hata
xv = Int(X): yv = Int(Y)
cval = Clicked(xv, yv)
If cval > -1 And Button = 1 Then
   Dragging = True
   NewLocPoint = cval
   Label1.Caption = "Çalýþtýðýnýz Nokta:  " + Trim$(cval + 1)
   Label3.Caption = "Son Çalýþýlan Nokta:  " + Trim$(cval + 1)
   
Else
If Check1.Value = 0 Then
        AddCont xv, yv
        RES1.Circle (xv, yv), 2, 255
        RES1.Print nc
End If

        If nc = 1 Then
            PSet (xv, yv)
        Else
            RES1.DrawStyle = vbDot
            RES1.Line (Cont(nc - 2, 0), Cont(nc - 2, 1))-(Cont(nc - 1, 0), Cont(nc - 1, 1)), 0
            RES1.DrawStyle = vbSolid
        End If
        If nc > 1 Then YenidenCIZ
    End If
Exit Sub
Hata:
MsgBox "(" & Err.Number & ") Nokta eklenemedi! ('Nokta Ekleme' sekmesi iþaretli mi?)", vbExclamation
End Sub
Private Sub Res1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Clicked(X, Y) > -1 Then
        MousePointer = vbCrosshair
    Else
        MousePointer = vbDefault
    End If


    If Dragging = True Then
        xv = Int(X): yv = Int(Y)
        Cont(NewLocPoint, 0) = xv: Cont(NewLocPoint, 1) = yv
        YenidenCIZ
    End If
End Sub
Private Sub Res1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Dragging = True Then
        Dragging = False
        YenidenCIZ
        If Check1.Value = 1 Then
           Label1.Caption = "Noktalarý taþýmak için klavuzlarý sürükleyin..."
        Else
           Label1.Caption = "Yeni nokta eklemek için görüntüye týklayýn!"
        End If
    End If
End Sub
Private Function Clicked(X, Y)
    For i = 0 To nc
        xp = Cont(i, 0): yp = Cont(i, 1)
        If Abs(xp - X) < 3 And Abs(yp - Y) < 3 Then
            Clicked = i
            Exit Function
        End If
    Next i
    Clicked = -1
End Function
Sub YenidenCIZ()
    'YenidenCIZ ekran çizgileri
    RES1.Cls
    For i = 1 To nc
        xv = Cont(i - 1, 0): yv = Cont(i - 1, 1)
        If Check5.Value = 0 And Check6.Value = 0 Then
           If Check2.Value = 0 Then RES1.Circle (xv, yv), 2, 255 'Klavuz noktasý rengi
           If Check4.Value = 0 Then RES1.Print i
        End If
    Next i
    'Klavuz çizgi biçimi
    If Option2.Value = True Then
       RES1.DrawStyle = vbDot
    Else
       RES1.DrawStyle = vbSolid
    End If
    For i = 0 To nc - 2
        'Klavuz Çizgileri
        If Check5.Value = 0 Then
           If Option1.Value = True Then
              If Check3.Value = 0 Then RES1.Line (Cont(i, 0), Cont(i, 1))-(Cont(i + 1, 0), Cont(i + 1, 1)), Renk   'Klavuz çizgisi rengi
           Else
              If Check3.Value = 0 Then RES1.Line (Cont(i, 0), Cont(i, 1))-(Cont(i + 1, 0), Cont(i + 1, 1)), 0   'Klavuz çizgisi rengi
           End If
        End If
    Next i
    RES1.DrawStyle = vbSolid
    DrawBezier Smooth
    Me.Cls
End Sub
Sub DrawBezier(du)
    Rem On Error Resume Next
    n = nc - 1
    If n < 1 Then
       Rem  MsgBox "Konrol edilebilecek aktif nokta yok", vbInformation
        Exit Sub
    End If
    If Option2.Value = True Then
       If Check6.Value = 0 Then RES1.PSet (Cont(0, 0), Cont(0, 1))
    End If
    For u = 0 To 1 Step du
        X = 0: Y = 0
        For k = 0 To n
            bv = B(k, n, u)
            X = X + Cont(k, 0) * bv
            Y = Y + Cont(k, 1) * bv
        Next k
        If Option2.Value = True Then
           If Check6.Value = 0 Then RES1.Line -(X, Y), Renk ' Çizilen Renk (255 kýrmýzý)
        End If
    Next u
    RES1.Line -(Cont(n, 0), Cont(n, 1)), 255
End Sub

 

 


 

