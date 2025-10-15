Sub SalvarAtas()

ActiveDocument.MailMerge.DataSource.ActiveRecord = wdFirstRecord

    For i = 1 To ActiveDocument.MailMerge.DataSource.RecordCount

    With ActiveDocument.MailMerge
        .Destination = wdSendToNewDocument
        .SuppressBlankLines = True
        With .DataSource
            .FirstRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
            .LastRecord = ActiveDocument.MailMerge.DataSource.ActiveRecord
            ata = .DataFields("ata").Value
            ata = Format(ata, "000")
            ano = .DataFields("ano_ata").Value
            fornecedor = .DataFields("nome_arquivo").Value
        End With
        pasta = Application.ActiveDocument.Path
        .Execute Pause:=False
    End With
    
    ActiveDocument.SaveAs2 FileName:=pasta & "\Ata " & ata & "-" & ano & " - " & fornecedor & ".docx", _
        FileFormat:=wdFormatXMLDocument, LockComments:=False, Password:="", _
        AddToRecentFiles:=True, WritePassword:="", ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=14
    ActiveWindow.Close
    
    ActiveDocument.MailMerge.DataSource.ActiveRecord = wdNextRecord
    
    Next i

End Sub