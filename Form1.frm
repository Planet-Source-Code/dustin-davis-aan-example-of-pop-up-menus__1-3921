VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pop-up Menu"
   ClientHeight    =   4350
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuPopUp 
      Caption         =   "Pop-Up"
      Begin VB.Menu mnuPopUp_WebPage 
         Caption         =   "Our Web Page"
      End
      Begin VB.Menu mnuPopUp_Somethingelse 
         Caption         =   "Something else"
      End
      Begin VB.Menu mnuPopUp_Ugly 
         Caption         =   "Pictures of my mom"
      End
      Begin VB.Menu mnuPopUp_About 
         Caption         =   "About"
      End
      Begin VB.Menu mnuPopUp_Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Pop-Up menu example
'Dustin Davis
'Bootleg Software Inc.
'Http://www.warpnet.org/bsi
'
'I hope this helps you! If not, well, I dunno
'Please feel free to e-mail me for questions. I get lonley sometimes!!
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Sub Form_Load()
mnuPopUp.Visible = False 'you can have this on if you want
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'this is the place to control the buttons
If Button = 2 Then 'if they right click, 1=left, 2=right
    Form1.PopupMenu mnuPopUp 'show popup menu
Else 'else if they clicked the left button
    DoEvents
End If
End Sub

Private Sub mnuPopUp_About_Click()
'and still, they keep coming!
MsgBox "Dustin Davis" & vbCrLf & "Bootleg Software Inc." & vbCrLf & "Planet Source Code Rules!"
End Sub

Private Sub mnuPopUp_Exit_Click()
'this is one of the menu functions. You can add more
Unload Me
End Sub

Private Sub mnuPopUp_Somethingelse_Click()
'yet another menu function
MsgBox "Alot more junk!"
End Sub

Private Sub mnuPopUp_Ugly_Click()
'still, another function of the menu!
MsgBox "HEY! Are you some kind of sick wierd`o? Why would you wanna see my mom naked?!"
End Sub

Private Sub mnuPopUp_WebPage_Click()
'another menu function
MsgBox "This is where you add stuff!"
End Sub
