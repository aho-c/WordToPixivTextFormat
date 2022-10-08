Attribute VB_Name = "WordToPixivTextFormat"
Option Explicit

'/**************************
' Convert to Pixiv Text Format.
'   (c) 2022 ahoc
' License: MIT License
'
' Note: Check Microsoft "VBScript Regular Expressions 5.5" in references.
'**************************/

Private Const PIXIV_CE_UTF16LE = 1200 'UTF-16LE

Public Sub ToPixivTextFormat()
    If MsgBox("Start converting to pixiv text format?", vbQuestion + vbYesNo, "Start confirmation") = vbNo Then Exit Sub

    If Not validatePixivTextFormat Then Exit Sub

    Application.ScreenUpdating = False
    
    Dim tempFileName As String: tempFileName = Format(Date, "yyyymmdd") & Format(Time, "hhmmss") & "_ptf_temp_file.ptffile.temp"
    Dim tempFilePath As String: tempFilePath = ActiveDocument.Path & "\" & tempFileName
    
    copyDocumentToPixivText tempFilePath
    
    Application.ScreenUpdating = True

    MsgBox "Saved to clipboard.", vbInformation, "Final report"
End Sub

Private Function validatePixivTextFormat() As Boolean
    validatePixivTextFormat = False
    
    If ActiveDocument.Fields.Count > 0 Then
        If ActiveDocument.Fields(1).ShowCodes Then ActiveDocument.Fields.ToggleShowCodes
    End If
    
    If ActiveDocument.Tables.Count > 0 Then _
        MsgBox "Tables cannot be converted.", vbExclamation, "Please confirm.": Exit Function
    
    Dim fieldsObject As Fields: Set fieldsObject = ActiveDocument.Fields
    Dim fieldObject As Field
    Dim fieldText As String
    Dim patternText As String: patternText = "EQ \\\* jc[0-4] \\\* "".+?"" \\\* hps[0-9]{1,} \\o.*?\(\\s\\up [0-9]{1,}\((.+?)\),(.+?)\)"
    For Each fieldObject In fieldsObject
        fieldText = fieldObject.Code.text
        If Not regTestForPixivText(fieldText, patternText) Then _
            MsgBox "There are fields that are not ruby.", vbExclamation, "Please confirm.": Exit Function
    Next fieldObject

    If regTestForPixivText(ActiveDocument.Range.text, Chr(19) & patternText & Chr(21)) Then _
        MsgBox "There are ruby field code characters.", vbExclamation, "Please confirm.": Exit Function

    validatePixivTextFormat = True
End Function

Private Sub copyDocumentToPixivText(temp_file_path As String)
    initUtf16PixivTextFile temp_file_path
    
    Dim markup: markup = Array()
    Dim paras As Paragraphs: Set paras = ActiveDocument.Paragraphs
    Dim para As Paragraph
    Dim paraText As String: paraText = ""
    Dim paraCount As Long: paraCount = 1
    
    For Each para In paras
        markup = convParagraphToPixivTextFormat(para)
        paraText = convDocumentToPixivText(para)
        paraText = para.Range.ListFormat.ListString & Mid(paraText, 1, Len(paraText) - 1)
        
        appendUtf16PixivTextFile temp_file_path, markup(0) & paraText & markup(1) & vbCrLf

        Application.StatusBar = CStr(paraCount) & "/" & CStr(paras.Count)
        paraCount = paraCount + 1
    Next para
    Application.StatusBar = False

    copyUtf16PixivText temp_file_path
End Sub

Private Function convDocumentToPixivText(paragraph_object As Paragraph) As String
    If ActiveDocument.Fields.Count > 0 Then
        If Not ActiveDocument.Fields(1).ShowCodes Then ActiveDocument.Fields.ToggleShowCodes
    End If
    
    convDocumentToPixivText = convNewPageToPixivNewPage(paragraph_object.Range.text)
    convDocumentToPixivText = convRubyToPixvRuby(convDocumentToPixivText)

    If ActiveDocument.Fields.Count > 0 Then
        If ActiveDocument.Fields(1).ShowCodes Then ActiveDocument.Fields.ToggleShowCodes
    End If
End Function

Private Function convNewPageToPixivNewPage(paragraph_text As String) As String
    convNewPageToPixivNewPage = regReplaceForPixivText(paragraph_text, Chr(12) & vbCr & "$", "[newpage]" & vbCr)
End Function

Private Function convRubyToPixvRuby(paragraph_text As String)
    Dim regMatchs As MatchCollection
    Dim patternText As String: patternText = Chr(19) & "EQ \\\* jc[0-4] \\\* "".+?"" \\\* hps[0-9]{1,} \\o.*?\(\\s\\up [0-9]{1,}\((.+?)\),(.+?)\)" & Chr(21)
    convRubyToPixvRuby = regReplaceForPixivText(paragraph_text, patternText, "[[rb:$1 > $2]]")
End Function

Private Function convParagraphToPixivTextFormat(paragraph_object As Paragraph)
    convParagraphToPixivTextFormat = Array("", "")
    
    If Len(paragraph_object.Range.text) = 1 Then Exit Function
    
    Dim paraFirstChar As String: paraFirstChar = paragraph_object.Range.Characters(1).text
    Dim paraText As String: paraText = paragraph_object.Range.Characters(Len(paragraph_object.Range.text) - 1).text
    Dim existsStartSerif As Boolean: existsStartSerif = regTestForPixivText(paraFirstChar, "^[ÅuÅwÅyÅmÅkÅiÅoÅqÅsÅÉÅe ]$")
    Dim existsStopSerif As Boolean: existsStopSerif = regTestForPixivText(paraText, "^[ÅvÅxÅzÅnÅlÅjÅpÅrÅtÅÑÅf]$")
    
    If _
        (LenB(paragraph_object.Range.text) > 2) And (paragraph_object.Range.text <> vbCr) And (paragraph_object.Range.text <> Chr(12) & vbCr) _
    Then
        If _
            (Not ((existsStartSerif) And (existsStopSerif))) _
            And _
            (InStr(1, paragraph_object.Style, "å©èoÇµ ") = 1) _
        Then
            convParagraphToPixivTextFormat = Array("[chapter:", "]")
        ElseIf _
            (Not ((existsStartSerif) And (existsStopSerif))) _
            And _
            ((paragraph_object.Style <> "ï\ëË") And (paragraph_object.Style <> "ïõëË") And (InStr(1, paragraph_object.Style, "å©èoÇµ ") = 0)) _
        Then
            convParagraphToPixivTextFormat = Array("Å@", "")
        End If
    End If
End Function

Private Sub initUtf16PixivTextFile(file_path As String)
    Dim fileNumber As Integer: fileNumber = FreeFile
    
    Open file_path For Output As #fileNumber: Close #fileNumber
End Sub

Private Function appendUtf16PixivTextFile(file_path As String, utf16_text As String) As Boolean
    Dim bytes() As Byte: bytes = utf16_text
    appendUtf16PixivTextFile = writeUtf16PixivBinaryFile(file_path, bytes, -1)
End Function

Private Function saveUtf16PixivTextFile(file_path As String, utf16_text As String, Optional ByVal start_pointer As Long = 1) As Boolean
    Dim bytes() As Byte: bytes = utf16_text
    saveUtf16PixivTextFile = writeUtf16PixivBinaryFile(file_path, bytes, start_pointer)
End Function

Private Function writeUtf16PixivBinaryFile(file_path As String, utf16_bytes() As Byte, Optional ByVal start_pointer As Long = 1) As Boolean
    writeUtf16PixivBinaryFile = False
    
    Dim fileNumber As Integer: fileNumber = FreeFile
    Open file_path For Binary As #fileNumber
        If start_pointer <= 0 Then start_pointer = LOF(fileNumber) + CLng(1)
        If start_pointer > LOF(fileNumber) + CLng(1) Then Close #fileNumber: Exit Function
        Seek #fileNumber, start_pointer
        Put #fileNumber, , utf16_bytes
    Close #fileNumber

    writeUtf16PixivBinaryFile = True
End Function

Private Function readUtf16PixivBinaryFile(file_path As String, Optional ByVal start_pointer As Long = 1, Optional ByVal read_bytes As Long = 1) As Byte()
    Dim fileNumber As Integer: fileNumber = FreeFile
    
    Open file_path For Binary As #fileNumber
        If read_bytes <= 0 Then read_bytes = LOF(fileNumber)
        If read_bytes > LOF(fileNumber) + CLng(1) Then Close #fileNumber: Exit Function
        
        Seek #fileNumber, start_pointer
        readUtf16PixivBinaryFile = InputB(LOF(fileNumber), fileNumber)
    Close #fileNumber
End Function

Private Sub copyUtf16PixivText(file_path As String)
    Documents.Open fileName:=file_path, _
        ConfirmConversions:=False, _
        Format:=wdOpenFormatAuto, _
        Encoding:=PIXIV_CE_UTF16LE
    
    ActiveDocument.Range.Copy
    ActiveDocument.Close SaveChanges:=False
    
    Kill file_path
End Sub

Private Function regTestForPixivText( _
    target_text As String, _
    pattern_text As String, _
    Optional global_mode As Boolean = True, _
    Optional ignore_case As Boolean = True, _
    Optional multi_line As Boolean = False _
) As Boolean
    Dim reg As RegExp: Set reg = New RegExp
    With reg
        .Global = global_mode
        .IgnoreCase = ignore_case
        .MultiLine = multi_line
        .Pattern = pattern_text
        regTestForPixivText = .Test(target_text)
    End With
End Function

Private Function regReplaceForPixivText( _
    target_text As String, _
    pattern_text As String, _
    replace_text As String, _
    Optional global_mode As Boolean = True, _
    Optional ignore_case As Boolean = True, _
    Optional multi_line As Boolean = False _
) As String
    Dim reg As RegExp: Set reg = New RegExp
    With reg
        .Global = global_mode
        .IgnoreCase = ignore_case
        .MultiLine = multi_line
        .Pattern = pattern_text
        regReplaceForPixivText = .Replace(target_text, replace_text)
    End With
End Function

Private Function getRegMatchsForPixivText( _
    target_text As String, _
    pattern_text As String, _
    Optional global_mode As Boolean = True, _
    Optional ignore_case As Boolean = True, _
    Optional multi_line As Boolean = False _
) As MatchCollection
    Dim reg As RegExp: Set reg = New RegExp
    With reg
        .Global = global_mode
        .IgnoreCase = ignore_case
        .MultiLine = multi_line
        .Pattern = pattern_text
        Set getRegMatchsForPixivText = .Execute(target_text)
    End With
End Function

