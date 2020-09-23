VERSION 5.00
Begin VB.Form frmpop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Popup menus"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnucolour 
      Caption         =   "Changecolour"
      Visible         =   0   'False
      Begin VB.Menu mnublack 
         Caption         =   "Black"
      End
      Begin VB.Menu mnuwhite 
         Caption         =   "White"
      End
      Begin VB.Menu mnuyellow 
         Caption         =   "Yellow"
      End
      Begin VB.Menu mnugreen 
         Caption         =   "Green"
      End
      Begin VB.Menu mnublue 
         Caption         =   "Blue"
      End
      Begin VB.Menu mnured 
         Caption         =   "Red"
      End
   End
End
Attribute VB_Name = "frmpop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' simple example of using popupmenus in your program, this could be incorprated
' into textboxs to change fonts, text size, or what ever you wish, i hope you
' found this useful
' Make the menu in "menu editor, located next to addform, make a menu structure
' you may need to learn how to do this, once made make sure the desired menu which
' in this instance would be "mnucolour" is set to invisable.

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then            ' if the mousebutton 2 is pressed (right)
      PopupMenu mnucolour      ' then it calls the menu which is hidden to be shown
   End If                      ' at the cordinates of the mousepointer (x,y)
End Sub

Private Sub mnublack_Click()
frmpop.BackColor = vbBlack     ' This changes the form colour to black
End Sub

Private Sub mnublue_Click()
frmpop.BackColor = vbBlue      ' This changes the form colour to blue
End Sub

Private Sub mnugreen_Click()
frmpop.BackColor = vbGreen     ' This changes the form colour to green
End Sub

Private Sub mnured_Click()
frmpop.BackColor = vbRed       ' This changes the form colour to red
End Sub

Private Sub mnuwhite_Click()
frmpop.BackColor = vbWhite     ' This changes the form colour to white
End Sub

Private Sub mnuyellow_Click()
frmpop.BackColor = vbYellow    ' This changes the form colour to yellow
End Sub
