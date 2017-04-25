Option Explicit
Private objNS As Outlook.NameSpace
Private WithEvents objNewMailItems As Outlook.Items

Private Sub Application_Startup()
 
    Dim objMyInbox As Outlook.MAPIFolder
    Set objNS = Application.GetNamespace("MAPI")
    Set objMyInbox = objNS.GetDefaultFolder(olFolderInbox)
    Set objNewMailItems = objMyInbox.Items
    Set objMyInbox = Nothing
    
End Sub

Private Sub objNewMailItems_ItemAdd(ByVal Item As Object)
    Call editTitle(Item)
End Sub

Function editTitle(Item As Object)
    Dim Msg As Outlook.MailItem
    Dim re As Object
    If TypeName(Item) = "MailItem" Then

        Set Msg = Item
        Set re = CreateObject("vbscript.regexp")
        re.Pattern = "((Reply)|(Antwort)|(AW)|(Re)):( )*"
        re.Global = True
            
        If Msg.Subject Like "*Antwort:*" Or Msg.Subject Like "*AW:*" Or Msg.Subject Like "*Re:*" Then
            Msg.Subject = re.Replace(Msg.Subject, "")
            Msg.Subject = "Re: " + Msg.Subject
            Msg.Save
        End If
        
  End If
End Function


Sub iterateEmails()
    Dim ns As Outlook.NameSpace
    Set ns = CreateObject("Outlook.Application").GetNamespace("MAPI")
    Dim DefaultInboxFldr As MAPIFolder
    Set DefaultInboxFldr = ns.GetDefaultFolder(olFolderInbox)
    Dim Item As Object
    
    For Each Item In DefaultInboxFldr.Items
        Call editTitle(Item)
    Next Item

End Sub
