' Permet d'attacher une ou plusieurs pièces jointes à un courrier (même fichiers pour l'ensemble des destinataires)
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)


    Dim objFolder As Object
    Dim objFile As Object
    Dim DEST As Recipient
    Dim SMTPTO As String

    If Item.Class = olMail Then
        Dim objCurrentMessage As MailItem
        Set objCurrentMessage = Item
        ' Permet de vérifier que le sujet du courrier contient le terme "PUBLIIDEM"
        If UCase(objCurrentMessage.Subject) Like "*PUBLIIDEM*" Then
            On Error Resume Next
            Dim i As Long
            i = 0
            If publipostagePJ <> "" Then
                ' Permet d'ajouter la même pièce jointe à chacun de ses correspondants
                While publipostagePJ(i) <> "fin"
                    objCurrentMessage.Attachments.Add Source:=publipostagePJ(i)
                    i = i + 1
                Wend
            End If

            'On supprime le terme PUBLIIDEM du sujet et le courrier s'envoie avec l'ensemble des pièces jointes
            objCurrentMessage.Subject = Replace(objCurrentMessage.Subject, "PUBLIIDEM ", "")
        End If
        Set objCurrentMessage = Nothing
    End If
End Sub