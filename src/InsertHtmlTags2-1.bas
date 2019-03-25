Attribute VB_Name = "NewMacros"
Sub InsertHTMLTags2()
Attribute InsertHTMLTags2.VB_Description = "InsertHTMLTags v 2.1b"
Attribute InsertHTMLTags2.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.InsertHTMLTags2"
'
' Вставляет теги HTML
'
'
    
    'Переход в начало документа
    Selection.HomeKey Unit:=wdStory
    'Замена < на &lt;
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = "<"
        .Replacement.Text = "&lt;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    'Переход в начало документа
    Selection.HomeKey Unit:=wdStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .Text = ">"
        .Replacement.Text = "&gt;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    'Переход в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Поиск того, что по центру
    Selection.Find.ClearFormatting
    With Selection.Find
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Вставка соответствующих тегов
    InsertTags "<center>", "</center>"
    'Возврат в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Поиск того, что выделено "Курьером"
    Selection.Find.ClearFormatting
    With Selection.Find
        .Font.NameAscii = "Courier New"
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Вставка соответствующих тегов
    InsertTags "<code>", "</code>"
    'Возврат в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Поиск того, что выделено жирным
    Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Bold = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Вставка соответствующих тегов
    InsertTags "<b>", "</b>"
    'Возврат в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Поиск того, что выделено курсивом
    Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Italic = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Вставка соответствующих тегов
    InsertTags "<i>", "</i>"
    'Возврат в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Поиск того, что выделено зачеркнутым
    Selection.Find.ClearFormatting
    With Selection.Find
        .Font.StrikeThrough = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Вставка соответствующих тегов
    InsertTags "<s>", "</s>"
    'Возврат в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Поиск того, что выделено надстрочным индексом
    Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Superscript = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Вставка соответствующих тегов
    InsertTags "<sup>", "</sup>"
    'Возврат в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Поиск того, что выделено подстрочным индексом
    Selection.Find.ClearFormatting
    With Selection.Find
        .Font.Subscript = True
        .Text = ""
        .Replacement.Text = ""
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    'Вставка соответствующих тегов
    InsertTags "<sub>", "</sub>"
    'Возврат в начало документа
    Selection.HomeKey Unit:=wdStory
    
    'Очистка условий поиска
    Selection.Find.ClearFormatting
    
End Sub
Private Sub ShiftArr(ByRef arr As Variant, pos As Long, ctr As Long)
    For I = 1 To ctr
        For J = pos To UBound(arr) - 1
            arr(J) = arr(J + 1)
        Next J
    Next I
End Sub
Private Sub InsertTags(OpenTag As String, CloseTag As String)
    Dim Start_s() As Long
    Dim End_s() As Long
    Dim I As Long
    
    ctr = 0
    shiftctr = 0
    Selection.Find.Execute
    If Not Selection.Find.Found Then Exit Sub
    'Поиск необходимых интервалов
    Do While Selection.Find.Found
        ctr = ctr + 1
        ReDim Preserve Start_s(ctr)
        ReDim Preserve End_s(ctr)
        Start_s(ctr) = Selection.Start
        End_s(ctr) = Selection.End
        Selection.Find.Execute
    Loop
    
    
    ' Удаление лишних интервалов (в которых конец текущего интервала совпадает с началом следующего)
    For I = 1 To ctr - 1
        Do While End_s(I) = Start_s(I + 1)
            ShiftArr Start_s, I + 1, 1
            ShiftArr End_s, I, 1
            shiftctr = shiftctr + 1
        Loop
    Next I
    
    'Изменение размерности - удаление ненужных элементов
    ReDim Preserve Start_s(UBound(Start_s) - shiftctr)
    ReDim Preserve End_s(UBound(End_s) - shiftctr)
    
    'Вставка тегов
    TagLen = Len(OpenTag) + Len(CloseTag)
    AllTagLen = TagLen
    For I = 1 To UBound(Start_s)
        Selection.Start = Start_s(I)
        Selection.End = End_s(I)
        Selection.Text = OpenTag + Selection.Text + CloseTag
        If I <> UBound(Start_s) Then
            Start_s(I + 1) = Start_s(I + 1) + AllTagLen
            End_s(I + 1) = End_s(I + 1) + AllTagLen
            AllTagLen = AllTagLen + TagLen
        End If
    Next I
End Sub
Sub InsertLink()
Attribute InsertLink.VB_Description = "Insert A HREF TAG"
Attribute InsertLink.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.InsertLink"

'data format:
'Link text [http://example.com/example/example.html]
'Result:
'<b><a href="http://example.com/example/example.html" target="_blank">Link text</a></b>
'Select text & run macros

    Dim LinkTemplate As String
    Dim LinkAddr As String
    Dim LinkText As String
    Dim LinkOut As String
    Dim LinkStart As Integer
    Dim LinkEnd As Integer
    
    LinkTemplate = "<b><a href=""%addr%"" target=""_blank"">%text%</a></b>"
    LinkStart = InStr(1, Selection.Text, "[")
    LinkEnd = InStr(1, Selection.Text, "]")
    If (LinkStart = 0) Or (LinkEnd = 0) Then
        MsgBox "No Link :("
        Exit Sub
    End If
    
    LinkAddr = Trim$(Mid$(Selection.Text, LinkStart + 1, LinkEnd - LinkStart - 1))
    LinkText = Trim$(Mid$(Selection.Text, 1, LinkStart - 1))
    
    LinkOut = Replace(LinkTemplate, "%addr%", LinkAddr)
    LinkOut = Replace(LinkOut, "%text%", LinkText)
    
    Selection.Text = LinkOut
    Selection.MoveRight Unit:=wdCharacter, Count:=1

End Sub
Sub InsertIMG()
Attribute InsertIMG.VB_Description = "Insert IMG Tag"
Attribute InsertIMG.VB_ProcData.VB_Invoke_Func = "Project.NewMacros.InsertIMG"

    'replace image address to IMG HTML tag
    'data format:
    '[http://example.com/images/example.jpg]
    'Result:
    '<img src="http://example.com/images/example.jpg">
    'Select text & run macros
    
    Dim IMGTemplate As String
    Dim IMGAddr As String
    Dim IMGOut As String
    Dim IMGStart As Integer
    Dim IMGEnd As Integer
    
    IMGTemplate = "<img src=""%addr%"">"
    IMGStart = InStr(1, Selection.Text, "[")
    IMGEnd = InStr(1, Selection.Text, "]")
    If (IMGStart = 0) Or (IMGEnd = 0) Then
        MsgBox "No Link :("
        Exit Sub
    End If
    
    IMGAddr = Trim$(Mid$(Selection.Text, IMGStart + 1, IMGEnd - IMGStart - 1))
    IMGOut = Replace(IMGTemplate, "%addr%", IMGAddr)
    
    Selection.Text = IMGOut
    Selection.MoveRight Unit:=wdCharacter, Count:=1

End Sub
