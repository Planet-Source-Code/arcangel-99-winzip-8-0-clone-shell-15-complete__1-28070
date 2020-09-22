Attribute VB_Name = "modSubMain"
Sub Main()

'\\ Show the Main form Window.

frmMain.Show

End Sub



Function Exit_App()

'\\ Close all Open Windows to Save Memory then End the Application.


Unload frmAbout
Unload frmMain


End Function
