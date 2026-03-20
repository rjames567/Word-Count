Sub ExcludeTableCaptionAndHeadingWordsFromWordCount()
    Dim objTable As Table
    Dim objParagraph As Paragraph
    Dim objDoc As Document
    
    Dim nTableWords As Long
    Dim nDocWords As Long
    Dim nWordCount As Long
    Dim nCaptionWords As Long
    Dim nHeadingWords As Long
    
    Dim styleName As String
    
    Set objDoc = ActiveDocument
    
    nTableWords = 0
    nCaptionWords = 0
    nHeadingWords = 0
    
    nDocWords = objDoc.ComputeStatistics(wdStatisticWords)
    
    With objDoc
        For Each objTable In .Tables
            nTableWords = nTableWords + objTable.Range.ComputeStatistics(wdStatisticWords)
        Next objTable
    End With
    
    With objDoc
        For Each objParagraph In .Paragraphs
            styleName = CStr(objParagraph.Range.Style)
            
            If styleName = "Caption" Then
                nCaptionWords = nCaptionWords + objParagraph.Range.ComputeStatistics(wdStatisticWords)
            
            ElseIf styleName = "Title" _
                Or styleName = "Subtitle" _
                Or Left$(styleName, 7) = "Heading" Then
                
                nHeadingWords = nHeadingWords + objParagraph.Range.ComputeStatistics(wdStatisticWords)
            End If
        Next objParagraph
    End With
    
    nWordCount = nDocWords - nTableWords - nCaptionWords - nHeadingWords
    
    MsgBox "There are " & nWordCount & " main text words in this document." & vbCrLf & vbCrLf & _
           "The following items are excluded from the word count:" & vbCrLf & _
           "• Total words in tables: " & nTableWords & vbCrLf & _
           "• Total caption words: " & nCaptionWords & vbCrLf & _
           "• Total heading/title/subtitle words: " & nHeadingWords
End Sub
