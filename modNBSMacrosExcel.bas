Attribute VB_Name = "modNBSMacrosExcel"
' ========================================================================
'
' Macro code written to help work with MS Excel
' single function currently around generating cost spreadsheets
'
' Contributors:
' 1. Stephen Hamil
'
' Use at own risk
' Always save and keep copies of Word documents before running the macros
'
' Support Community - https://support.thenbs.com/support/home
'
' modNBSMacrosExcel     3rd Dec 2020
'
' modNBSMacros
' ============
' This is the main module with the primary function cals
'
' 1. ExtractPriceableHeadingsFromWord() - loops through an exported DOCX file and generates costing spreadsheet for either CAWS or Uniclass 2015
' ========================================================================

' Note for improvements - would work better with decent progress bar
' https://www.excel-easy.com/vba/examples/progress-indicator.html
' May also want consider openfiledialog and saveas dialog
' https://www.automateexcel.com/vba/open-file-dialog/

' Set of enums to use once classification system has been determined
Enum enumClassificationType
    enClassCAWS = 1
    enClassMasterFormat = 2
    enClassUniclass = 3
    enClassUnknown = 4
End Enum



' Primary function - ExtractPriceableHeadingsFromWord()
' ==============================
'
' Designed to work with any DOCX file exported from Chorus
'
' Will open the file location specified in cell E2
' Then loop through the section titles
' and add each of these as priceable items
' finally, the cells will be formatted as currency and an auto-add for the total cell at the end
' ===============================
Public Sub ExtractPriceableHeadingsFromWord()
    ' Better error handling *could* be written - coded below for things to work well
    
    ' Basic error check - have they put a file path in cell E2
    ' If not we'll kick them straight off
    Dim strFilePath As String
    strFilePath = Cells(2, 5).Value
    
    Dim bProblem As Boolean
    bProblem = False
    
    If strFilePath = "" Then bProblem = True
    If InStr(1, strFilePath, ".DOCX", vbTextCompare) = 0 Then bProblem = True
        
      
    If bProblem = True Then
        MsgBox "Please enter a valid file path to a DOCX file in cell E2"
        Exit Sub
    End If
    
    
    ' Open the Word document
    Set wordapp = New Word.Application
    Dim wordDoc As Word.Document
    Set wordDoc = wordapp.documents.Open(strFilePath, ReadOnly:=True)
    
    
    
    ' Loop through every paragraph until we can work out the classification type - then exit loop
    ' It would be nice if we were working with a proper object model
    ' but we're not - so we'll have to try and find patterns
    Dim eClass As enumClassificationType
    
    Dim sLine As Paragraph
    Dim strLineText As String
    Dim intPos As Integer
    Dim strSectionCode As String
    
        
    Dim strGroupCode As String
    Dim strSectionTitle As String
    
    Dim strClauseCode As String
    Dim strClauseTitle As String

    Dim strFullSectionLine As String
    
    Dim strGroupLetter As String
    Dim strPreviousGroupLetter As String
    Dim strGroupTitle As String
    
    Dim nCount As Integer
    nCount = 0
    
    Dim nRowCount As Integer
    nRowCount = 3 'the cost items will be added starting at row three of spreadsheet
    
    eClass = ReturnClassification(wordDoc)
    
    ' Following code will be executed if spec is Uniclass 2015 or CAWS
    If eClass = enClassCAWS Or eClass = enClassUniclass Then
        System.Cursor = wdCursorWait
        
        DoEvents
        
    
        ' Loop through every paragraph in the word document
        ' looking out for the word styles we need
        For Each sLine In wordDoc.Paragraphs
                    
            Select Case sLine.Style
                ' First time into a new group - we need to grab the group code...
                
                Case "chorus-section-header"
                    strLineText = sLine.Range.Text ' get the text of the paragraph
                
                    ' We have text of the format E10Concrete
                    ' Splitting on the  character
                    intPos = InStr(1, strLineText, "", vbTextCompare)
                    strSectionCode = Mid(strLineText, 1, intPos - 1)
                    strSectionTitle = Mid(strLineText, intPos + 1, Len(strLineText) - intPos - 1)
                    strSectionCode = Trim(strSectionCode)
                    strSectionTitle = Trim(strSectionTitle)
                    
                    ' We want to add a number of cells to get the row looking like F10 -> Brick and Blockwork -> £00.00
                    Cells(nRowCount, 1).Value = strSectionCode
                    Cells(nRowCount, 2).Value = strSectionTitle
                    Cells(nRowCount, 3).NumberFormat = "$#,##0.00" ' standard currency code format
                    
                    ' increment row count for next loop
                    nRowCount = nRowCount + 1
                    
                    ' print out the debug info
                    strFullSectionLine = strSectionCode & vbTab & strSectionTitle
                    Debug.Print "strFullSectionLine: " & strFullSectionLine
                    
                    
                Case "chorus-clause-title"
                    ' could be developed further to do all clauses as well
                    ' see the MS Word macro for keynotes to copy logic
                    
            End Select
            
        Next sLine
        
        Dim nItemsTotal As Integer
        nItemsTotal = nRowCount - 2 ' work out what the formula below needs to do the adding up
        
        ' Finalise things by adding a row at the bottom to add the costs up as they are typed
        Cells(nRowCount + 1, 2).Value = "Total cost"
        Cells(nRowCount + 1, 2).Font.Bold = True
        Cells(nRowCount + 1, 3).FormulaR1C1 = "=SUM(R[-" & nItemsTotal & "]C:R[-2]C)" ' again, standard MS Excel formula
              
        System.Cursor = wdCursorNormal
        MsgBox "Complete"
    
    Else
        MsgBox "The Word document must be CAWS or Uniclass 2015 structure"
    End If
    
    ' try and close down the word objects so files don't get locked and memory hogged up etc...
    wordDoc.Close
    Set wordDoc = Nothing
    Set wordapp = Nothing
End Sub

' Method to quickly check the content looks right before we start
Private Function ReturnClassification(objDoc As Word.Document) As enumClassificationType

    Dim eClass As enumClassificationType
    Dim sLine As Paragraph
    Dim strLineText As String
    Dim intPos As Integer
    Dim strSectionCode As String
    

    ' Loop through the lines until we can determine what classification it is
    For Each sLine In objDoc.Paragraphs
        strLineText = sLine.Range.Text ' get the text of the paragraph
        
        Select Case sLine.Style
            ' First time into a new group - we need to grab the group code...
            Case "chorus-section-header"
                intPos = InStr(1, strLineText, "", vbTextCompare)
                strSectionCode = Mid(strLineText, 1, intPos - 1)
                
                ' For CAWS the section code is three digits - C10 for example
                If Len(strSectionCode) = 3 Then
                    eClass = enClassCAWS
                    Exit For
                End If
                
                ' For Uniclass the section code contains an underbar _ - Ss_25_30_25 for example
                If InStr(1, strSectionCode, "_", vbTextCompare) <> 0 Then
                    eClass = enClassUniclass
                    Exit For
                End If
                
                ' All others, we'll assume Masterformat for now
                eClass = enClassMasterFormat
                Exit For ' no point in sending the code through every line
        End Select
        
    Next sLine

    ReturnClassification = eClass
End Function
