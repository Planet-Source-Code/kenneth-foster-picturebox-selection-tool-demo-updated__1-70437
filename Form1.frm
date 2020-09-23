VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Demo of Selection Tool "
   ClientHeight    =   6615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   ScaleHeight     =   441
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   703
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save picture"
      Height          =   420
      Left            =   3630
      TabIndex        =   3
      Top             =   6090
      Width           =   1365
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   5445
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   42
      TabIndex        =   2
      Top             =   210
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox Picture1 
      Height          =   5880
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   5820
      ScaleWidth      =   4785
      TabIndex        =   1
      Top             =   150
      Width           =   4845
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   3  'Dot
         Height          =   270
         Left            =   150
         Top             =   120
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Label Label1 
      Caption         =   "BY DIGITAL VISIT MY SITE HTTP://WWW.DIGITALFX2001.COM"
      Height          =   495
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   2805
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'*******************************************************************
'**                     Picturebox Selection Tool Demo
'**                               Version 1.2.3
'**                               By Ken Foster
'**                                 April 2008
'**                     Freeware--- no copyrights claimed
'*******************************************************************

Option Explicit
'

Private Sub Form_Load()
   'set up our properties
   Form1.ScaleMode = 3
   Picture1.AutoRedraw = True
   Picture1.ScaleMode = 3
   Picture2.AutoRedraw = True
   Picture2.ScaleMode = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
   ReleaseSR   'release cursor...just in case form is closed while cursor is still confined
   Unload Me
End Sub

Private Sub cmdSave_Click()
   SavePicture Picture2.Picture, App.Path & "\Test.bmp"
   MsgBox "Picture Saved in " & App.Path & "\Test.bmp"
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   DrawSR Picture1
   
   'reset to zero's for next selection
   Shape1.Height = 0
   Shape1.Width = 0
   Picture2.Height = 0
   Picture2.Width = 0
   
   'I'm using button 1 for this demo, you may want button 2 in your programs
   'if so, change all the Button = 1 to Button = 2
   If Button = 1 Then    'set starting points
      Picture1.MousePointer = 2    'change pointer to cross
      Shape1.Left = X
      Shape1.Top = Y
   End If

   Shape1.Visible = True    'show rectangle
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   On Error GoTo Skip
   If Button = 1 Then     'draw the selection rectangle
      Picture2.Visible = False
      Shape1.Height = Y - Shape1.Top
      Shape1.Width = X - Shape1.Left
   
      'make picture2 the same dimensions as the selection rectangle
       Picture2.Width = Shape1.Width
       Picture2.Height = Shape1.Height
   End If
Skip:
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      Shape1.Visible = False    'now we can hide the rectangle
      
      'plaster the selection into picture2
      If Picture2.Width > 3 Or Picture2.Height > 3 Then
         Picture2.Visible = True   'make image visible
         Picture2.Picture = Picture2.Image   'render picture so it can be used
         BitBlt Picture2.hDC, 0, 0, Picture2.Width, Picture2.Height, Picture1.hDC, Shape1.Left, Shape1.Top, vbSrcCopy
      End If
   End If
   Picture1.MousePointer = 0          'set pointer back to default
   ReleaseSR
End Sub
