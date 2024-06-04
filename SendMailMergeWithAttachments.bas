Sub SendMailMergeWithAttachments()

    Dim Source As Document
    Dim MailList As Workbook
    Dim DataSheet As Worksheet
    Dim i As Long
    Dim OutApp As Object
    Dim OutMail As Object
    Dim LastRow As Long

    ' Initialiser l'objet Outlook
    Set OutApp = CreateObject("Outlook.Application")

    ' Ouvrir la source de données Excel
    Set MailList = Workbooks.Open("C:\Path\To\Your\ExcelFile.xlsx")
    Set DataSheet = MailList.Sheets("Sheet1")

    ' Définir le document Word actuel comme source
    Set Source = ActiveDocument

    ' Obtenir la dernière ligne avec des données
    LastRow = DataSheet.Cells(DataSheet.Rows.Count, "A").End(xlUp).Row

    ' Boucle à travers chaque ligne de la feuille Excel
    For i = 2 To LastRow ' Supposant que la première ligne contient les en-têtes

        ' Créer un nouvel email
        Set OutMail = OutApp.CreateItem(0)

        With OutMail
            .To = DataSheet.Cells(i, 1).Value
            .Subject = "Sujet de votre email"
            .Body = Source.Content.Text

            ' Ajouter la pièce jointe si disponible
            If DataSheet.Cells(i, 2).Value <> "" Then
                .Attachments.Add DataSheet.Cells(i, 2).Value
            End If

            ' Envoyer l'email
            .Send
        End With

        ' Libérer l'objet email
        Set OutMail = Nothing

    Next i

    ' Fermer la source de données Excel
    MailList.Close SaveChanges:=False

    ' Libérer l'objet Outlook
    Set OutApp = Nothing

    MsgBox "Les emails ont été envoyés avec succès !"

End Sub
