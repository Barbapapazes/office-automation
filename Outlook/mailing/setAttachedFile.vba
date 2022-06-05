Public publipostagePJ As Variant

Sub setPublipostage()
  On Error Resume Next

  If publipostagePJ(0) = "" Then publipostagePJ = Array("fin", "fin", "fin", "fin", "fin", "fin", "fin", "fin", "fin", "fin")

  While publipostagePJ(i) <> "fin"
    contenu = contenu & vbCr & publipostagePJ(i)
    i = i + 1
  Wend

  If contenu = "" Then contenu = "vide"

  modifier = MsgBox(contenu & vbCr & "Voulez vous modifier les fichiers ?", vbYesNo, "Fichiers paramétrés")
    
  If modifier = vbYes Then
    For i = 0 To 9
      If i > 0 Then encore = MsgBox("un autre ?", vbYesNo)
        quest:
      If encore <> vbNo Then
        PJ = InputBox("Emplacement du fichier joint au PUBLIPOSTAGE?", _
          "Paramétrage du PUBLIPOSTAGE pour la session", publipostagePJ(i))
      If "" = Dir(PJ, vbNormal) Then GoTo quest
        publipostagePJ(i) = PJ
      Else: Exit For
      End If
      Next i
  End If
  MsgBox "Votre publipostage doit comporter le terme :" & vbCr & "PUBLIIDEM" & vbCr & "dans le sujet." & vbCr & "Celui-ci sera retiré lors de l'envoi"

End Sub