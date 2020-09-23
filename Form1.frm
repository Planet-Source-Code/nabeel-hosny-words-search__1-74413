VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "Form1"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00CA8826&
      BorderStyle     =   0  'None
      Height          =   11040
      Left            =   11130
      ScaleHeight     =   736
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   276
      TabIndex        =   2
      Top             =   240
      Width           =   4140
      Begin VB.CommandButton Hint 
         BackColor       =   &H0000C000&
         Caption         =   "Hint"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   1560
         MouseIcon       =   "Form1.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   1800
         Width           =   735
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   2895
         Left            =   120
         TabIndex        =   21
         Top             =   7440
         Visible         =   0   'False
         Width           =   3135
         Begin VB.ListBox List4 
            Height          =   1230
            Left            =   1320
            MultiSelect     =   1  'Simple
            TabIndex        =   26
            Top             =   1560
            Width           =   1695
         End
         Begin VB.ListBox listall 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1320
            ItemData        =   "Form1.frx":1126
            Left            =   120
            List            =   "Form1.frx":1832
            TabIndex        =   25
            Top             =   240
            Width           =   1455
         End
         Begin VB.ListBox List3 
            Height          =   1230
            Left            =   600
            TabIndex        =   24
            Top             =   1560
            Width           =   735
         End
         Begin VB.ListBox List2 
            Height          =   1230
            ItemData        =   "Form1.frx":27D0
            Left            =   1680
            List            =   "Form1.frx":27EC
            TabIndex        =   23
            Top             =   240
            Width           =   615
         End
         Begin VB.ListBox List1 
            Height          =   1230
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   495
         End
         Begin VB.Image img 
            Height          =   255
            Index           =   1
            Left            =   2760
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Image img 
            Height          =   255
            Index           =   0
            Left            =   2400
            Stretch         =   -1  'True
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2520
            TabIndex        =   27
            ToolTipText     =   "Lenght of word"
            Top             =   960
            Width           =   345
         End
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FF8080&
         DrawWidth       =   10
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         MouseIcon       =   "Form1.frx":2810
         MousePointer    =   99  'Custom
         ScaleHeight     =   15
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   258
         TabIndex        =   15
         ToolTipText     =   "Current Postion"
         Top             =   10650
         Width           =   3900
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   0
            ScaleHeight     =   16
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   1
            TabIndex        =   16
            Top             =   0
            Width           =   15
         End
         Begin VB.Label lblSBar 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00CA8826&
            BackStyle       =   0  'Transparent
            Caption         =   "nabeelhosny@yahoo.com"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   600
            TabIndex        =   17
            Top             =   -30
            Width           =   2730
         End
      End
      Begin VB.PictureBox Picword 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   3840
         Left            =   120
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   12
         Top             =   5670
         Width           =   2655
         Begin VB.ListBox listrnd 
            Appearance      =   0  'Flat
            BackColor       =   &H00808000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3630
            ItemData        =   "Form1.frx":2B1A
            Left            =   120
            List            =   "Form1.frx":2B30
            MultiSelect     =   1  'Simple
            TabIndex        =   13
            Top             =   120
            Width           =   2430
         End
      End
      Begin VB.TextBox txtguss 
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000004&
         Height          =   405
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   7860
         Width           =   1170
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00CA8826&
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   2880
         TabIndex        =   10
         Top             =   6120
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00CA8826&
         Caption         =   "PAUSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   405
         Left            =   2850
         TabIndex        =   9
         Top             =   5640
         Width           =   1200
      End
      Begin VB.PictureBox exitme 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   2400
         MouseIcon       =   "Form1.frx":2B60
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":2CB2
         ScaleHeight     =   780
         ScaleWidth      =   1545
         TabIndex        =   8
         Top             =   9750
         Width           =   1545
      End
      Begin VB.PictureBox startme 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   780
         Left            =   180
         MouseIcon       =   "Form1.frx":3394
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":34E6
         ScaleHeight     =   52
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   103
         TabIndex        =   7
         Top             =   9750
         Width           =   1545
      End
      Begin VB.Label labmove 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   480
         Left            =   2970
         TabIndex        =   20
         Top             =   9000
         Width           =   900
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moves"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   360
         Left            =   2895
         TabIndex        =   19
         Top             =   8520
         Width           =   1035
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Right"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   435
         Left            =   2940
         TabIndex        =   18
         Top             =   6720
         Width           =   930
      End
      Begin VB.Label labcount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   480
         Left            =   2970
         TabIndex        =   14
         Top             =   7200
         Width           =   900
      End
      Begin VB.Image Image1 
         Height          =   3960
         Index           =   0
         Left            =   2100
         Picture         =   "Form1.frx":3C69
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Image Image1 
         Height          =   3960
         Index           =   1
         Left            =   120
         Picture         =   "Form1.frx":14C5D
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   1920
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Score :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000A&
         Height          =   555
         Left            =   180
         TabIndex        =   6
         Top             =   5040
         Width           =   1740
      End
      Begin VB.Label labscore 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "000000"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   540
         Left            =   2160
         TabIndex        =   5
         Top             =   5040
         Width           =   1740
      End
      Begin VB.Label gamelab 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WordMat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   42
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   1020
         Index           =   1
         Left            =   45
         TabIndex        =   3
         Top             =   30
         Width           =   3885
      End
      Begin VB.Label gamelab 
         BackStyle       =   0  'Transparent
         Caption         =   "WordMat"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   42
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   1020
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   4320
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   0
         Left            =   2880
         Shape           =   4  'Rounded Rectangle
         Top             =   6600
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000080FF&
         BorderWidth     =   6
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   1095
         Index           =   1
         Left            =   2880
         Shape           =   4  'Rounded Rectangle
         Top             =   8400
         Width           =   1095
      End
   End
   Begin VB.PictureBox Back 
      BorderStyle     =   0  'None
      Height          =   11040
      Left            =   0
      ScaleHeight     =   736
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   734
      TabIndex        =   0
      Top             =   240
      Width           =   11010
      Begin VB.PictureBox pic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         DrawWidth       =   8
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1080
         Index           =   0
         Left            =   240
         MouseIcon       =   "Form1.frx":1625E
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":163B0
         ScaleHeight     =   72
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   1
         Tag             =   "A"
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   840
      Top             =   1080
   End
   Begin VB.Image yellowletter 
      Height          =   1080
      Left            =   9720
      Picture         =   "Form1.frx":16966
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Image blueletter 
      Height          =   1080
      Left            =   9360
      Picture         =   "Form1.frx":16F84
      Top             =   840
      Width           =   1080
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hrgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim i As Integer, dummy As Integer
 Const MainPicHeight = 720
 Const MainPicWidth = 720
Dim nArr(99)

Private Sub Check1_Click()
If Check1.Value = 1 Then
For i = 0 To 99
  pic(i).Picture = blueletter.Picture
   nArr(i) = pic(i).Tag
   pic(i).Cls
   pic(i).CurrentX = (72 - pic(i).TextWidth(nArr(i))) \ 2
   pic(i).CurrentY = (72 - pic(i).TextHeight(nArr(i))) \ 2
   pic(i).Print nArr(i)
   DoEvents
   Next
sndPlaySound App.Path & "\pic\wrong.wav", 1: DoEvents
   txtguss.Text = ""
   Check1.Value = 0
 End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
Timer1.Enabled = False
Else
Timer1.Enabled = True
End If

End Sub

Private Sub exitme_Click()
End
End Sub

Private Sub Form_Load()
' 65 90   97 122
 Me.Caption = ""
 Me.WindowState = 2
 Me.Move 0, 0, Screen.Width, Screen.Height
SetWindowRgn Check2.hwnd, CreateRoundRectRgn(0, 0, Check2.Width, Check2.Height, 10, 10), True
SetWindowRgn Check1.hwnd, CreateRoundRectRgn(0, 0, Check1.Width, Check1.Height, 10, 10), True
SetWindowRgn exitme.hwnd, CreateRoundRectRgn(0, 0, exitme.Width, exitme.Height, 30, 30), True
SetWindowRgn startme.hwnd, CreateRoundRectRgn(0, 0, startme.Width, startme.Height, 30, 30), True
SetWindowRgn txtguss.hwnd, CreateRoundRectRgn(0, 0, txtguss.Width, txtguss.Height, 10, 10), True
SetWindowRgn Hint.hwnd, CreateRoundRectRgn(0, 0, Hint.Width, Hint.Height, 50, 50), True
MaxPicControl = 100
        
    dummy = Sqr(MaxPicControl)
    
    For i = 0 To MaxPicControl - 1
        If i <> 0 Then Load pic(i)
        pic(i).Width = MainPicWidth / dummy
        pic(i).Height = MainPicHeight / dummy
        pic(i).Left = 8 + (i Mod dummy) * pic(i).Width
        pic(i).Top = 9 + (i \ dummy) * MainPicHeight / dummy
        pic(i).Visible = True
        pic(i).BorderStyle = 1
        nArr(i) = ""
    Next i
ColBox Picword, 0, 0, Picword.ScaleWidth, Picword.ScaleHeight, 7, 64, 128, 0, 228, 255, 0
ColBox Back, 0, 0, Back.ScaleWidth, Back.ScaleHeight, 7, 64, 128, 0, 128, 255, 0
ColBox Picture1, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, 7, 64, 128, 0, 128, 255, 0

clearme

swapme

End Sub
Sub swapme()
Dim idd As Integer

Randomize Timer
Randomize
listrnd.Clear
 For idd = 1 To 1000
acard:
 card1 = Int(Rnd(1) * listall.ListCount + 0)
 card2 = Int(Rnd(1) * listall.ListCount + 0)
 If card1 = card2 Then
 GoTo acard
 End If
 Number1 = listall.List(card1)
 listall.List(card1) = listall.List(card2)
 listall.List(card2) = Number1
 Next idd
 For idd = 1 To 1000
 card1 = Int(Rnd(1) * List2.ListCount + 0)
 card2 = Int(Rnd(1) * List2.ListCount + 0)
 Number1 = List2.List(card1)
 List2.List(card1) = List2.List(card2)
 List2.List(card2) = Number1
 Next idd
 For i = 0 To 5
 listrnd.AddItem listall.List(i)
 Next
cmdSort2_Click

End Sub
Private Sub cmdSort2_Click()
Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim temp As String
 iCount = listrnd.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With listrnd
        If Len(listrnd.List(i)) < Len(listrnd.List(i + 1)) Then
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
        End If
     End With
    Next i
Next j

End Sub

Private Sub Hint_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim vv As Integer
If List4.ListCount = 0 Then Exit Sub
For vv = 0 To List4.ListCount - 1
If List4.Selected(vv) = False Then
img(0).Picture = pic(Val(Mid(List4.List(vv), 4, 2))).Image
img(1).Picture = pic(Val(Right(List4.List(vv), 2))).Image
pic(Val(Mid(List4.List(vv), 4, 2))).Circle (10, 10), 7, vbRed
pic(Val(Right(List4.List(vv), 2))).Circle (10, 10), 7, vbRed
sndPlaySound App.Path & "\pic\alarm_beep.wav", 1: DoEvents
Exit For
End If
Next
End Sub

Private Sub Hint_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If List4.ListCount = 0 Then Exit Sub
For vv = 0 To List4.ListCount - 1
If List4.Selected(vv) = False Then
pic(Val(Mid(List4.List(vv), 4, 2))).Picture = img(0).Picture
pic(Val(Right(List4.List(vv), 2))).Picture = img(1).Picture
sndPlaySound App.Path & "\pic\alarm_beep.wav", 1: DoEvents
Exit For
End If
Next
End Sub

Private Sub labcount_Change()
If labcount.Caption = "0000" Then Exit Sub
If Val(labcount.Caption) Mod 6 = 0 Then
sndPlaySound App.Path & "\pic\won.wav", 1: DoEvents
tx = MsgBox("  Congratulations! ? " & vbCrLf & "Do u Want to play Again : ", vbYesNo + vbQuestion, " Start a New Game")
     If tx = vbNo Then exitme_Click
     If tx = vbYes Then clearme: startme_Click
End If
End Sub

Private Sub List2_Click()
If List2.List(List2.ListIndex) = "UU" Then dome: UU: addme
If List2.List(List2.ListIndex) = "DD" Then dome: DD: addme
If List2.List(List2.ListIndex) = "RR" Then dome: RR: addme
If List2.List(List2.ListIndex) = "LL" Then dome: LL: addme
If List2.List(List2.ListIndex) = "RU" Then dome: RU: addme
If List2.List(List2.ListIndex) = "RD" Then dome: RD: addme
If List2.List(List2.ListIndex) = "LU" Then dome: LU: addme
If List2.List(List2.ListIndex) = "LD" Then dome: LD: addme

End Sub
Sub dome()
Dim Word As String
'Dim Wrong As Single
'Dim Template As String
List1.Clear: List3.Clear
listrnd.ListIndex = List2.ListIndex
Word = listrnd.List(List2.ListIndex)
Label1.Caption = Val(Label1.Caption) + Len(listrnd.List(List2.ListIndex))
For i = 1 To Len(Word)
   C = Mid(Word, i, 1)
   List1.AddItem C
   'If C = MyGuess Then
   '  Mid(Template, i, 1) = C
   'End If
   DoEvents
Next i

End Sub
Sub addme()
List4.List(listrnd.ListIndex) = Format(listrnd.ListIndex, "00")
For i = 0 To List3.ListCount - 1
 pic(List3.List(i)).Tag = List1.List(i)
 pic(List3.List(i)).Visible = False
 List4.List(listrnd.ListIndex) = List4.List(listrnd.ListIndex) & " " & List3.List(i)
 DoEvents
 Next
List1.Clear: List3.Clear
End Sub
Sub RR()
Randomize
baba1:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card + Len(listrnd.List(listrnd.ListIndex)) - 1) > ((10 * Int(card \ 10)) + 9) Then GoTo baba1
 If (card + Len(listrnd.List(listrnd.ListIndex)) - 1) <= ((10 * Int(card \ 10)) + 9) Then
   If pic(card + i).Visible = False Then GoTo baba1
   If pic(card + i).Visible = True Then List3.AddItem Format(card + i, "00")
 End If
DoEvents
Next

End Sub
Sub LL()
Randomize
baba2:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card - Len(listrnd.List(listrnd.ListIndex)) - 1) < ((10 * Int(card \ 10))) Then GoTo baba2
 If (card - Len(listrnd.List(listrnd.ListIndex)) - 1) >= ((10 * Int(card \ 10))) Then
   If pic(card - i).Visible = False Then GoTo baba2
   If pic(card - i).Visible = True Then List3.AddItem Format(card - i, "00")
 End If
DoEvents
Next
End Sub
Sub UU()
Randomize
baba3:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card - 10 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) < 9 Then GoTo baba3
If (card - 10 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) >= 9 Then
 If pic(card - i * 10).Visible = False Then GoTo baba3
 If pic(card - i * 10).Visible = True Then List3.AddItem Format(card - i * 10, "00")
 End If
DoEvents
Next
End Sub
Sub DD()
Randomize
baba3:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card + 10 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) > 99 Then GoTo baba3
If (card + 10 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) <= 99 Then
 If pic(card + i * 10).Visible = False Then GoTo baba3
 If pic(card + i * 10).Visible = True Then List3.AddItem Format(card + i * 10, "00")
End If
DoEvents
Next
End Sub
Sub RU()
Randomize
baba5:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card - 9 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) < 9 Or _
   Int(card \ 10) - Int((card - 9 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) \ 10) <> Len(listrnd.List(listrnd.ListIndex)) - 1 _
   Then
   GoTo baba5
  Else
  If pic(card - i * 9).Visible = False Then GoTo baba5
  If pic(card - i * 9).Visible = True Then List3.AddItem Format(card - i * 9, "00")
End If
DoEvents
Next
End Sub
Sub RD()
Randomize
baba6:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card + 11 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) > 99 Or _
   Int(card \ 10) - Int((card + 11 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) \ 10) <> -1 * (Len(listrnd.List(listrnd.ListIndex)) - 1) _
   Then
   GoTo baba6
  Else
 If pic(card + i * 11).Visible = False Then GoTo baba6
 If pic(card + i * 11).Visible = True Then List3.AddItem Format(card + i * 11, "00")
End If
DoEvents
Next
End Sub
Sub LU()
Randomize
baba7:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card - 11 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) < 9 Or _
   Int(card \ 10) - Int((card - 11 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) \ 10) <> Len(listrnd.List(listrnd.ListIndex)) - 1 _
   Then
   GoTo baba7
  Else
 If pic(card - i * 11).Visible = False Then GoTo baba7
 If pic(card - i * 11).Visible = True Then List3.AddItem Format(card - i * 11, "00")
End If
DoEvents
Next
End Sub
Sub LD()
Randomize
baba8:
List3.Clear
card = Int(Rnd(1) * 100)
For i = 0 To Len(listrnd.List(listrnd.ListIndex)) - 1
If (card + 9 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) > 99 Or _
   Int(card \ 10) - Int((card + 9 * (Len(listrnd.List(listrnd.ListIndex)) - 1)) \ 10) <> -1 * (Len(listrnd.List(listrnd.ListIndex)) - 1) _
   Then
   GoTo baba8
  Else
 If pic(card + i * 9).Visible = False Then GoTo baba8
 If pic(card + i * 9).Visible = True Then List3.AddItem Format(card + i * 9, "00")
 DoEvents
End If
Next
End Sub

Public Sub ColBox(Obj As Object, BX%, BY%, EX%, EY%, h%, R%, G%, B%, RE%, GE%, BE%)
Dim H2%, H3%, IvR%, IvG%, IvB%
Obj.AutoRedraw = True
Obj.ScaleMode = 3 'pixel
H3 = Int(h / 2)
IvR = Int(RE - R) / H3
IvG = Int(GE - G) / H3
IvB = Int(BE - B) / H3
Do While h >= H3
Obj.Line (BX + H2, BY + H2)-(EX - H2, EY - H2), RGB(R, G, B), B
Obj.Line (BX + h, BY + h)-(EX - h, EY - h), RGB(R, G, B), B
h = h - 1
H2 = H2 + 1
R = R + IvR
G = G + IvG
B = B + IvB
Loop
End Sub



Private Sub pic_Click(Index As Integer)
If Check2.Value = 1 Then Exit Sub
labmove.Caption = Format(Val(labmove.Caption) + 1, "0000")
Timer1.Enabled = False
   sndPlaySound App.Path & "\pic\swish.wav", 1: DoEvents
   pic(Index).Picture = yellowletter.Picture
   nArr(Index) = pic(Index).Tag
   pic(Index).Cls
   pic(Index).CurrentX = (72 - pic(Index).TextWidth(nArr(Index))) \ 2
   pic(Index).CurrentY = (72 - pic(Index).TextHeight(nArr(Index))) \ 2
   pic(Index).Print nArr(Index)
   DoEvents
   txtguss.Text = txtguss.Text & pic(Index).Tag
   txtguss_Change
   Timer1.Enabled = True
End Sub

Private Sub startme_Click()
Dim idd As Integer
txtguss.Text = "": List4.Clear
Picture3.Width = 1
swapme

' reset the progress bar

 Label1.Caption = "": Timer1.Enabled = True
 For i = 0 To 99: pic(i).Visible = True: pic(i).Picture = blueletter.Picture: pic(i).Cls: pic(i).Tag = "": nArr(i) = "": DoEvents: Next
 
 For idd = 0 To 5
 List2.Selected(idd) = True
 DoEvents
 Next
 
 Randomize
 For i = 0 To 99
 If pic(i).Tag = "" Then pic(i).Tag = Chr(Int(Rnd(1) * 26 + 65))
 DoEvents
 Next
 For i = 0 To 99
   pic(i).Visible = True
   nArr(i) = pic(i).Tag
   pic(i).Cls
   pic(i).CurrentX = (72 - pic(i).TextWidth(nArr(i))) \ 2
   pic(i).CurrentY = (72 - pic(i).TextHeight(nArr(i))) \ 2
   pic(i).Print nArr(i)
   DoEvents
 Next
For i = 0 To 5
List2.Selected(i) = False
listrnd.Selected(i) = False
DoEvents
Next
End Sub

Private Sub txtguss_Change()
If Len(txtguss.Text) > 6 Then Check1.Value = 1: Check1_Click: DoEvents: Exit Sub
For i = 0 To listrnd.ListCount - 1
 If txtguss.Text = listrnd.List(i) Then
 List4.Selected(i) = True
 Timer1.Enabled = False
    listrnd.Selected(i) = True
    labscore.Caption = Format(Val(labscore.Caption) + Len(txtguss.Text) * 50, "000000")
    labcount.Caption = Format(Val(labcount.Caption) + 1, "0000")
    sndPlaySound App.Path & "\pic\yes.wav", 1: DoEvents
    txtguss.Text = ""
    Timer1.Enabled = True
    Exit Sub
End If
Next
End Sub
Private Sub Timer1_Timer()
'If Picture3.Width = 10 Then sndPlaySound App.Path & "\pic\alarm_beep.wav", 1: DoEvents
If Picture3.Width >= Picture2.Width Then
sndPlaySound App.Path & "\pic\gameover.wav", 1: DoEvents
    Timer1.Enabled = False
    tx = MsgBox("  Game Over ? " & vbCrLf & "Do u Want to play Again : ", vbYesNo + vbQuestion, " Start a New Game")
     If tx = vbNo Then exitme_Click
     If tx = vbYes Then clearme: startme_Click
   Else
    Picture3.Width = Picture3.Width + 1 ' 0.5
    sndPlaySound App.Path & "\pic\Bella.wav", 1: DoEvents
End If
End Sub
Sub clearme()
Picture2.ScaleWidth = 260
Picture3.Width = 1
labscore.Caption = "000000": labcount.Caption = "0000"
txtguss.Text = "": labmove.Caption = "0000"
End Sub


