VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Shutdown Schedule"
   ClientHeight    =   3405
   ClientLeft      =   10005
   ClientTop       =   7320
   ClientWidth     =   4260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   480
      Top             =   1560
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu curtime 
         Caption         =   ""
      End
      Begin VB.Menu seperator 
         Caption         =   "-"
      End
      Begin VB.Menu start 
         Caption         =   "Start Timer"
      End
      Begin VB.Menu stopit 
         Caption         =   "Stop Timer"
      End
      Begin VB.Menu reset 
         Caption         =   "Reset Timer"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim howlong
Dim newday
Dim curday
Dim shutTime
Dim moo
Dim newsec
Dim cursec
Dim newmin
Dim curmin
Dim newhour
Dim curhour
Public Counter As Integer
Public IconObject As Object



Private Sub Exit_Click()
    Unload Form1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    delIcon IconObject.Handle
    delIcon Form1.Icon.Handle
End Sub


Private Sub Form_Load()
    Set IconObject = Form1.Icon
    AddIcon Form1, IconObject.Handle, IconObject, "Animated TrayIcon"


Timer1.Enabled = False

howlong = 0
newday = 0
curday = 0
newsec = 0
cursec = 0
newmin = 0
 curmin = 0
newhour = 0
 curhour = 0
stopit.Enabled = False
 

    
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static Message As Long
    Message = X / Screen.TwipsPerPixelX
    Select Case Message
    Case WM_RBUTTONUP:
        Me.PopupMenu Popup
    End Select
End Sub



Private Sub reset_Click()
howlong = 0
newday = 0
curday = 0
newsec = 0
cursec = 0
newmin = 0
 curmin = 0
newhour = 0
 curhour = 0
If moo = 1 Then
moo = 0
curtime.Caption = "Timer Not Active"
End If
End Sub

Private Sub shed_Click()
Form1.Visible = True

End Sub

Private Sub shutclr_Click()
shutStatus = 0
shutday.Text = ""
shuthour.Text = ""
shutmin.Text = ""
shutsec.Text = ""

End Sub

Private Sub shutSet_Click()
shutStatus = 1
enable.Enabled = True

Form1.Visible = False

End Sub

Private Sub start_Click()
Timer1.Enabled = True
start.Enabled = False
stopit.Enabled = True


End Sub

Private Sub stopit_Click()
Timer1.Enabled = False
stopit.Enabled = False
start.Enabled = True
moo = 1



End Sub

Private Sub Text1_Change()
cursec = cursec + 1
Call calculate
End Sub


Private Sub Timer1_Timer()
Text1.Text = Second(Time)




End Sub
Private Function calculate()

If cursec >= 60 Then
    curmin = curmin + 1
    cursec = 0
End If



If cursec < 10 Then
newsec = "0" & cursec
Else
newsec = cursec

End If

If curmin >= 60 Then
    curhour = curhour + 1
curmin = 0
End If



If curmin < 10 Then
newmin = "0" & curmin
Else
newmin = curmin
End If

If curhour >= 24 Then
    curday = curday + 1
curhour = 0
End If


If curday < 10 Then
newday = "0" & curday
Else
newday = curday
End If

howlong = newday & " Days " & newhour & ":" & newmin & ":" & newsec

curtime.Caption = howlong



End Function

