VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selphie is so cuteeeeeeeeeee !"
   ClientHeight    =   8040
   ClientLeft      =   330
   ClientTop       =   330
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   Begin VB.CommandButton Command5 
      Caption         =   "Begin !"
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   7440
   End
   Begin VB.CommandButton Command4 
      Appearance      =   0  'Flat
      Caption         =   "very@gobytown.com"
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   7440
      Width           =   6015
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   6960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   495
      Left            =   7560
      TabIndex        =   3
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   495
      Left            =   7560
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Render"
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   6690
      Left            =   120
      ScaleHeight     =   446
      ScaleMode       =   0  'User
      ScaleWidth      =   551.71
      TabIndex        =   0
      Top             =   120
      Width           =   7260
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

'this program will create a magnifier animation
'you can change the command1_click(bla bla bla) sub to form1_paint(bla bla bla)
'to hide the rendering progress and autoexecute the timer after the rendering is done
'any question, better algorithm, comment please email me : very@gobytow.com
'thanx !
'i'm sorry i did'nt put any comment in the code
'forgive me for my laziness
'God Bless You !

Private maxframes As Integer
Private r As Integer
Private himgdump As Long
Private renderingdone As Boolean
Private frames As Integer, adder As Integer
Private prog As Integer

Private Sub Command1_Click()
Dim i As Integer
Dim cap As String

Me.MousePointer = vbHourglass
cap = Command4.Caption
Command1.Enabled = False
Command2.Enabled = False
Command4.Caption = "Rendering frames, Please Wait ...."


 For i = 0 To maxframes
  
  himgdump = CreateCompatibleBitmap(Picture1.hdc, Picture1.width, Picture1.height)
  hframe(i) = CreateCompatibleDC(Picture1.hdc)
  'MsgBox CStr(himgdump)
  err = SelectObject(hframe(i), himgdump)
  'MsgBox CStr(himgdump)
  err = DeleteObject(himgdump)
  
  Magnifier r - (i * shrink), hframe(i)
  prog = ProgressBar1.Value + (100 / (maxframes + 1))
  If prog > 100 Then prog = 100
  ProgressBar1.Value = prog
 
 Next i
  
MsgBox "Please push Begin ! to start the animation", vbOKOnly, "Rendering Complete"
Command5.Enabled = True
Command4.Caption = cap
renderingdone = True
Me.MousePointer = 0

End Sub

Private Sub Command2_Click()
 
  Timer1.Enabled = False
  Command1.Enabled = True
  Command2.Enabled = False
  Command5.Enabled = True
 
End Sub

Private Sub Command3_Click()
Dim i As Integer

 err = DeleteDC(himg)
 err = DeleteObject(hpic)
 
 For i = 0 To maxframes
  
  err = DeleteDC(hframe(i))
 
 Next i
  
 End
 
End Sub


Private Sub Command4_Click()
 
 MsgBox "Sorry, you can use your e-mail applications to send me an e-mail, Thanx ! ", vbOKOnly, "very@gobytown.com"

End Sub

Private Sub Command5_Click()
 
 Timer1.Enabled = True
 Command5.Enabled = False
 Command1.Enabled = False
 Command2.Enabled = True
  
End Sub

Private Sub Form_Load()
Dim i As Integer

 hpic = LoadImageA(0, App.Path & "/selphie.bmp", IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
 himg = CreateCompatibleDC(Picture1.hdc)
 err = SelectObject(himg, hpic)

 err = GetObjectA(hpic, Len(picinfo), picinfo)
 
 Picture1.width = picinfo.bwidth
 Picture1.height = picinfo.bheight
 ProgressBar1.Enabled = True
 ProgressBar1.Value = 0
 
 renderingdone = False
 frames = 0
 adder = 1
 
 Command2.Enabled = False
 Command5.Enabled = False
  
 If picinfo.bwidth > picinfo.bheight Then r = Int(picinfo.bheight / 2)
 If picinfo.bheight > picinfo.bwidth Then r = Int(picinfo.bwidth / 2)
   
 maxframes = Round((r / shrink)) - 1
 'MsgBox CStr(maxframes)
 ReDim hframe(maxframes)
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer

 err = DeleteDC(himg)
 err = DeleteObject(hpic)
 'MsgBox CStr(err)
 
 For i = 0 To maxframes
  
  DeleteDC (hframe(i))
 
 Next i
  
End Sub

Private Sub Timer1_Timer()
 
 If frames >= maxframes Then adder = -1
 If frames <= 0 Then adder = 1
 frames = frames + adder
 
 err = BitBlt(Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, hframe(maxframes - frames), 0, 0, vbSrcCopy)

End Sub
