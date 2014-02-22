Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
   ' MailSave Macro
   ' by Patrick Stockton (patrick@codejnki.com)
   ' Purpose: Saves all sent messages in a folder as .msg files for use in other applications
   
   ' Revision Log
   ' 0.1 -- Initial release 2/21/2014
   
   ' Check to see if there is a MailSave folder in my documents
   MailSavePath = Environ$("USERPROFILE") & "\My Documents\MailSave\"
   If Dir(MailSavePath, vbDirectory) = "" Then
       MkDir MailSavePath
   End If
       
   ' Create a folder for today, this is mostly to help keep things organized
   TodayFolder = MailSavePath & Format(Date, "yyyy-mm-d")
   If Dir(TodayFolder, vbDirectory) = "" Then
       MkDir TodayFolder
   End If
   
   ' The file name is going to be the first recipient address plus a date & time stamp
   MessageFileName = "\" & Item.Recipients(1) & "_" & Format(Now(), "yyyy-mm-dd_hh_mm_ss") & ".msg"
   
   ' Finally save the file
   Item.SaveAs TodayFolder & MessageFileName, olMSG
End Sub
