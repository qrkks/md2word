' 处理标题的格式化宏
Sub FormatTitleByHeadingStyle()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "标题" Then
            With para.Range
                .Font.NameFarEast = "黑体"
                .Font.Name = "黑体"
                .Font.Size = 18 ' 小二
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
        End If
    Next para
    MsgBox "题目格式化完成！"
End Sub

' 一级标题格式化宏
Sub FormatLevel1Heading()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 1" Or para.Style = "标题 1" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 16 ' 小三
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphCenter
            End With
        End If
    Next para
    MsgBox "一级标题格式化完成！"
End Sub

' 二级标题格式化宏
Sub FormatLevel2Heading()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 2" Or para.Style = "标题 2" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 14 ' 四号
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
            End With
        End If
    Next para
    MsgBox "二级标题格式化完成！"
End Sub

' 三级标题格式化宏
Sub FormatLevel3Heading()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "Heading 3" Or para.Style = "标题 3" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 12 ' 小四
                .Font.Bold = True
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphLeft
            End With
        End If
    Next para
    MsgBox "三级标题格式化完成！"
End Sub

' 正文格式化宏
Sub FormatBodyText()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        ' 假设正文样式为“正文文本”、“Normal”、“First Paragraph”、“正文”
        If para.Style = "正文文本" Or para.Style = "Normal" Or para.Style = "First Paragraph" Or para.Style = "正文" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 12 ' 小四
                .Font.Bold = False
                .Font.Color = wdColorBlack
                .ParagraphFormat.Alignment = wdAlignParagraphLeft ' 左对齐
                .ParagraphFormat.FirstLineIndent = 24 ' 首行缩进两字符
            End With
        End If
    Next para
    MsgBox "正文格式化完成！"
End Sub

Sub SetPageAndBodyFormat()
    ' 设置页面
    With ActiveDocument.PageSetup
        .PaperSize = wdPaperA4
        .TopMargin = CentimetersToPoints(3)
        .BottomMargin = CentimetersToPoints(2.5)
        .LeftMargin = CentimetersToPoints(3)
        .RightMargin = CentimetersToPoints(2.5)
    End With

    ' 设置正文行距为1.5倍
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "正文文本" Or para.Style = "Normal" Or para.Style = "First Paragraph" Then
            para.Range.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
        End If
    Next para

    MsgBox "页面和正文行距设置完成！"
End Sub

' 摘要格式化宏
Sub MergeAndFormatAbstract()
    Dim i As Integer
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim txt As String
    Dim contentTxt As String
    Dim rngEnd As Range
    Dim rng As Range
    
    For i = ActiveDocument.Paragraphs.Count To 1 Step -1
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        If txt = "摘要" Or Left(txt, 3) = "摘要：" Or _
           txt = "关键词" Or Left(txt, 4) = "关键词：" Or _
           txt = "Abstract" Or Left(txt, 9) = "Abstract:" Or _
           txt = "Keywords" Or Left(txt, 9) = "Keywords:" Then
            ' 处理摘要、关键词段落
            Dim needMerge As Boolean
            needMerge = (txt = "摘要" Or txt = "关键词" Or txt = "Abstract" Or txt = "Keywords")
            
            If needMerge Then
                ' 需要合并的情况
                MsgBox "找到段落内容: [" & para.Range.Text & "]"
                Set nextPara = para.Next
                If Not nextPara Is Nothing Then
                    MsgBox "找到内容段: [" & nextPara.Range.Text & "]"
                    contentTxt = nextPara.Range.Text
                    contentTxt = Replace(contentTxt, vbCr, "")
                    contentTxt = Replace(contentTxt, vbLf, "")
                    contentTxt = Trim(contentTxt)
                    ' 如果有下一个标题，只取前面部分
                    Dim nextTitlePos As Integer
                    If txt = "摘要" Or txt = "关键词" Then
                        nextTitlePos = InStr(contentTxt, "Abstract")
                        If nextTitlePos = 0 Then nextTitlePos = InStr(contentTxt, "Keywords")
                    ElseIf txt = "Abstract" Then
                        nextTitlePos = InStr(contentTxt, "Keywords")
                    End If
                    If nextTitlePos > 0 Then
                        contentTxt = Left(contentTxt, nextTitlePos - 1)
                    End If
                    ' 获取 para 段落的最后一个字符（段落符号）前的位置
                    Dim rngInsert As Range
                    Set rngInsert = para.Range.Duplicate
                    rngInsert.End = rngInsert.End - 1  ' 不包括段落符号
                    rngInsert.Collapse wdCollapseEnd
                    rngInsert.InsertAfter contentTxt
                    MsgBox "合并后段落内容: [" & para.Range.Text & "]"
                    nextPara.Range.Delete
                Else
                    MsgBox "未找到内容段"
                End If
            End If
            
            ' 统一格式化（合并后或直接格式化）
            ' 判断标题后是否有冒号，没有则补冒号
            Dim paraText As String
            paraText = para.Range.Text
            Dim titleLen As Integer
            If Left(txt, 3) = "摘要：" Or Left(txt, 4) = "关键词：" Then
                titleLen = 3
            ElseIf Left(txt, 9) = "Abstract:" Or Left(txt, 9) = "Keywords:" Then
                titleLen = 9
            Else
                titleLen = Len(txt)
            End If
            
            If Len(paraText) < titleLen + 1 Or _
               (Mid(paraText, titleLen + 1, 1) <> "：" And Mid(paraText, titleLen + 1, 1) <> ":") Then
                If txt = "摘要" Or Left(txt, 3) = "摘要：" Then
                    If Len(paraText) < 3 Or Mid(paraText, 3, 1) <> "：" Then
                        para.Range.Characters(2).InsertAfter "："
                    End If
                ElseIf txt = "关键词" Or Left(txt, 4) = "关键词：" Then
                    If Len(paraText) < 4 Or Mid(paraText, 4, 1) <> "：" Then
                        para.Range.Characters(3).InsertAfter "："
                    End If
                ElseIf txt = "Abstract" Or Left(txt, 9) = "Abstract:" Then
                    If Len(paraText) < 9 Or Mid(paraText, 9, 1) <> ":" Then
                        para.Range.Characters(8).InsertAfter ":"
                    End If
                ElseIf txt = "Keywords" Or Left(txt, 9) = "Keywords:" Then
                    If Len(paraText) < 9 Or Mid(paraText, 9, 1) <> ":" Then
                        para.Range.Characters(8).InsertAfter ":"
                    End If
                End If
            End If
            
            ' 设置段落样式和字体
            para.Style = ActiveDocument.Styles("正文文本")
            
            ' 根据语言设置字体
            If txt = "摘要" Or Left(txt, 3) = "摘要：" Or txt = "关键词" Or Left(txt, 4) = "关键词：" Then
                ' 中文：宋体
                With para.Range.Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = False
                    .Color = wdColorBlack
                End With
            Else
                ' 英文：Times New Roman
                With para.Range.Font
                    .NameFarEast = "Times New Roman"
                    .Name = "Times New Roman"
                    .Size = 12
                    .Bold = False
                    .Color = wdColorBlack
                End With
            End If
            
            ' 设置标题加粗
            If txt = "摘要" Or Left(txt, 3) = "摘要：" Then
                para.Range.Characters(1).Font.Bold = True
                para.Range.Characters(2).Font.Bold = True
                If para.Range.Characters.Count >= 3 Then
                    para.Range.Characters(3).Font.Bold = True
                End If
            ElseIf txt = "关键词" Or Left(txt, 4) = "关键词：" Then
                para.Range.Characters(1).Font.Bold = True
                para.Range.Characters(2).Font.Bold = True
                para.Range.Characters(3).Font.Bold = True
                If para.Range.Characters.Count >= 4 Then
                    para.Range.Characters(4).Font.Bold = True
                End If
            ElseIf txt = "Abstract" Or Left(txt, 9) = "Abstract:" Then
                For j = 1 To 8
                    para.Range.Characters(j).Font.Bold = True
                Next j
                If para.Range.Characters.Count >= 9 Then
                    para.Range.Characters(9).Font.Bold = True
                End If
            ElseIf txt = "Keywords" Or Left(txt, 9) = "Keywords:" Then
                For k = 1 To 8
                    para.Range.Characters(k).Font.Bold = True
                Next k
                If para.Range.Characters.Count >= 9 Then
                    para.Range.Characters(9).Font.Bold = True
                End If
            End If
            
            ' 最后设置段落缩进
            With para.Range.ParagraphFormat
                .Alignment = wdAlignParagraphLeft
                .FirstLineIndent = 24 ' 首行缩进两字符
            End With
        End If
    Next i
    MsgBox "摘要格式化完成！"


End Sub

' 目录处理相关宏

' 查找目录位置
Sub FindTableOfContentsPosition()
    Dim para As Paragraph
    For Each para In ActiveDocument.Paragraphs
        If Trim(Replace(para.Range.Text, vbCr, "")) = "目录" Then
            MsgBox "找到目录位置：" & para.Range.Start
            Exit Sub
        End If
    Next para
    MsgBox "未找到目录标记"
End Sub

' 插入目录
Sub InsertTableOfContents()
    Dim para As Paragraph
    Dim tocRange As Range
    Dim found As Boolean
    
    found = False
    For Each para In ActiveDocument.Paragraphs
        If Trim(Replace(para.Range.Text, vbCr, "")) = "目录" Then
            ' 找到目录位置，删除"目录"文字
            para.Range.Delete
            ' 在该位置插入分页符
            Set tocRange = para.Range
            tocRange.Collapse wdCollapseStart
            tocRange.InsertBreak Type:=wdPageBreak
            ' 插入目录标题
            tocRange.InsertAfter "目录" & vbCr
            tocRange.Collapse wdCollapseEnd
            ' 先设置目录标题为TOC 标题样式，再格式化
            Dim tocTitlePara As Paragraph
            Set tocTitlePara = tocRange.Paragraphs(1)
            tocTitlePara.Style = ActiveDocument.Styles("TOC 标题")
            With tocTitlePara.Range.Font
                .NameFarEast = "宋体"
                .Name = "Times New Roman"
                .Size = 18 ' 小二
                .Bold = True
                .Color = wdColorBlack
            End With
            With tocTitlePara.Range.ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .FirstLineIndent = 0
            End With
            ' 插入目录域代码
            tocRange.Fields.Add Range:=tocRange, Type:=wdFieldTOC, Text:="", PreserveFormatting:=True
            ' 更新目录
            tocRange.Fields.Update
            found = True
            MsgBox "目录插入完成！"
            Exit For
        End If
    Next para
    
    If Not found Then
        MsgBox "未找到目录标记，请在文档中插入'目录'段落"
    End If
End Sub

' 更新目录
Sub UpdateTableOfContents()
    Dim fld As Field
    Dim updated As Boolean
    
    updated = False
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldTOC Then
            fld.Update
            updated = True
        End If
    Next fld
    
    If updated Then
        MsgBox "目录更新完成！"
    Else
        MsgBox "未找到目录域，请先插入目录"
    End If
End Sub

' 设置目录格式
Sub FormatTableOfContents()
    Dim para As Paragraph
    Dim tocTitle As String
    
    ' 查找目录标题段落
    For Each para In ActiveDocument.Paragraphs
        If Trim(Replace(para.Range.Text, vbCr, "")) = "目录" Then
            ' 先设置目录标题为TOC 标题样式，再格式化
            para.Style = ActiveDocument.Styles("TOC 标题")
            With para.Range.Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 18 ' 小二
                .Bold = True
                .Color = wdColorBlack
            End With
            With para.Range.ParagraphFormat
                .Alignment = wdAlignParagraphCenter
                .FirstLineIndent = 0
            End With
            MsgBox "目录标题格式设置完成！"
            Exit Sub
        End If
    Next para
    MsgBox "未找到目录标题"
End Sub

' 完整的目录处理（查找位置、插入目录、设置格式）
Sub ProcessTableOfContents()
    ' 先查找目录位置
    FindTableOfContentsPosition
    ' 插入目录
    InsertTableOfContents
    ' 设置目录格式
    FormatTableOfContents
    ' 更新目录
    UpdateTableOfContents
    MsgBox "目录处理完成！"
End Sub

' 正文中数字和英文字体格式化宏
Sub FormatNumbersAndEnglishInBody()
    Dim para As Paragraph
    Dim char As Range
    Dim charText As String
    Dim i As Integer
    
    For Each para In ActiveDocument.Paragraphs
        ' 只处理正文段落
        If para.Style = "正文文本" Or para.Style = "Normal" Or para.Style = "First Paragraph" Or para.Style = "正文" Then
            ' 遍历段落中的每个字符
            For i = 1 To para.Range.Characters.Count
                Set char = para.Range.Characters(i)
                charText = char.Text
                
                ' 检查是否为数字或英文字符
                If IsNumeric(charText) Or _
                   (Asc(charText) >= 65 And Asc(charText) <= 90) Or _  ' A-Z
                   (Asc(charText) >= 97 And Asc(charText) <= 122) Then ' a-z
                    ' 设置为Times New Roman
                    With char.Font
                        .Name = "Times New Roman"
                        .NameFarEast = "Times New Roman"
                    End With
                End If
            Next i
        End If
    Next para
    
    MsgBox "正文中数字和英文字体格式化完成！"
End Sub

' 更高效的正文中数字和英文字体格式化宏（使用正则表达式）
Sub FormatNumbersAndEnglishInBodyAdvanced()
    Dim para As Paragraph
    Dim paraText As String
    Dim i As Integer
    Dim j As Integer
    Dim charCode As Integer
    Dim isEnglishOrNumber As Boolean
    
    For Each para In ActiveDocument.Paragraphs
        ' 只处理正文段落
        If para.Style = "正文文本" Or para.Style = "Normal" Or para.Style = "First Paragraph" Or para.Style = "正文" Then
            ' 遍历段落中的每个字符
            For i = 1 To para.Range.Characters.Count
                charCode = Asc(para.Range.Characters(i).Text)
                
                ' 检查是否为数字或英文字符
                isEnglishOrNumber = False
                If charCode >= 48 And charCode <= 57 Then  ' 0-9
                    isEnglishOrNumber = True
                ElseIf charCode >= 65 And charCode <= 90 Then  ' A-Z
                    isEnglishOrNumber = True
                ElseIf charCode >= 97 And charCode <= 122 Then ' a-z
                    isEnglishOrNumber = True
                End If
                
                If isEnglishOrNumber Then
                    ' 设置为Times New Roman
                    With para.Range.Characters(i).Font
                        .Name = "Times New Roman"
                        .NameFarEast = "Times New Roman"
                    End With
                End If
            Next i
        End If
    Next para
    
    MsgBox "正文中数字和英文字体格式化完成！"
End Sub

