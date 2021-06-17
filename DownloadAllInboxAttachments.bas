Attribute VB_Name = "Module1"

Sub SaveOutlookAttachments()
    
    Dim ol As Outlook.Application
    Dim ns As Outlook.NameSpace
    Dim fol As Outlook.Folder
    Dim i As Object
    Dim mi As Outlook.MailItem
    Dim at As Outlook.Attachment
    Dim fso As Scripting.FileSystemObject
    Dim dir As Scripting.Folder
    Dim dirName As String
    
    Set fso = New Scripting.FileSystemObject
    
    Set ol = New Outlook.Application
    Set ns = ol.GetNamespace("MAPI")
    Set fol = ns.GetDefaultFolder(olFolderInbox)
    
    For Each i In fol.Items
    
        If i.Class = olMail Then
        
            Set mi = i
            
            If mi.Attachments.Count > 0 Then
                'Debug.Print mi.SenderName, mi.ReceivedTime, mi.Attachments.Count
                
                dirName = _
                    "V:\Student\Kashish - Summer 2021\EmailAttachmentTesting\" & _
                    Format(mi.ReceivedTime, "yyyy-mm-dd hh-nn-ss ") & _
                    Left(Replace(mi.Subject, ":", ""), 10)
                
                If fso.FolderExists(dirName) Then
                    Set dir = fso.GetFolder(dirName)
                Else
                    Set dir = fso.CreateFolder(dirName)
                End If
                
                For Each at In mi.Attachments
                
                    'Debug.Print vbTab, at.DisplayName, at.Size
                    at.SaveAsFile dir.Path & "\" & at.FileName
                    
                Next at
                
            End If
            
        End If
    
    Next i
    
End Sub

