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
Sub FormatAbstractTitle()
    Dim para As Paragraph
    Dim txt As String
    For Each para In ActiveDocument.Paragraphs
        txt = Trim(para.Range.Text)
        ' 只处理以“摘要”开头的段落
        If Left(txt, 2) = "摘要" Then
            ' 检查第三个字符是否为冒号（全角或半角）
            Dim thirdChar As String
            If para.Range.Characters.Count >= 3 Then
                thirdChar = para.Range.Characters(3)
            Else
                thirdChar = ""
            End If
            If thirdChar <> "：" And thirdChar <> ":" Then
                para.Range.Characters(2).InsertAfter "："
            End If
            ' 格式化“摘要”二字
            With para.Range.Characters(1).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            With para.Range.Characters(2).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            ' 格式化冒号
            If para.Range.Characters.Count >= 3 Then
                With para.Range.Characters(3).Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = True
                    .Color = wdColorBlack
                End With
            End If
            ' 段落格式
            para.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            para.Range.ParagraphFormat.FirstLineIndent = 24
        End If
    Next para
    MsgBox "摘要标题格式化完成！"
End Sub

Sub FormatAbstractInlineAndMerge()
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim txt As String
    Dim absLen As Integer
    Dim rngContent As Range
    For Each para In ActiveDocument.Paragraphs
        txt = Trim(para.Range.Text)
        ' 只处理以“摘要”开头的段落
        If Left(txt, 2) = "摘要" Then
            absLen = 2
            ' 检查第三个字符是否为冒号（全角或半角）
            Dim thirdChar As String
            If para.Range.Characters.Count >= 3 Then
                thirdChar = para.Range.Characters(3)
            Else
                thirdChar = ""
            End If
            If thirdChar <> "：" And thirdChar <> ":" Then
                ' 在“摘要”后插入全角冒号
                para.Range.Characters(2).InsertAfter "："
            End If
            ' 合并下一段到本段冒号后
            Set nextPara = para.Next
            If Not nextPara Is Nothing Then
                Dim nextTxt As String
                nextTxt = Trim(nextPara.Range.Text)
                If nextTxt <> "" Then
                    ' 在当前段落末尾插入下一个段落内容
                    para.Range.Characters(para.Range.Characters.Count).InsertAfter nextTxt
                    ' 删除下一段
                    nextPara.Range.Delete
                End If
            End If
            ' 格式化“摘要”二字
            With para.Range.Characters(1).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            With para.Range.Characters(2).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            ' 格式化冒号
            If para.Range.Characters.Count >= 3 Then
                With para.Range.Characters(3).Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = True
                    .Color = wdColorBlack
                End With
            End If
            ' 格式化内容部分
            If para.Range.Characters.Count > 3 Then
                Set rngContent = para.Range.Duplicate
                rngContent.Start = rngContent.Start + 3 * 2 ' 3个中文字符
                With rngContent.Font
                    .NameFarEast = "宋体"
                    .Name = "Times New Roman"
                    .Size = 12
                    .Bold = False
                    .Color = wdColorBlack
                End With
            End If
            para.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            para.Range.ParagraphFormat.FirstLineIndent = 24
        End If
    Next para
    MsgBox "摘要合并与格式化完成！"
End Sub

Sub FormatAbstractTitleAndMerge()
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim txt As String
    For Each para In ActiveDocument.Paragraphs
        txt = Trim(para.Range.Text)
        ' 只处理以“摘要”开头的段落
        If Left(txt, 2) = "摘要" Then
            ' 检查第三个字符是否为冒号（全角或半角）
            Dim thirdChar As String
            If para.Range.Characters.Count >= 3 Then
                thirdChar = para.Range.Characters(3)
            Else
                thirdChar = ""
            End If
            If thirdChar <> "：" And thirdChar <> ":" Then
                para.Range.Characters(2).InsertAfter "："
            End If
            ' 合并下一段到本段冒号后
            Set nextPara = para.Next
            If Not nextPara Is Nothing Then
                Dim nextTxt As String
                nextTxt = Replace(Replace(nextPara.Range.Text, vbCr, ""), vbLf, "") ' 去掉换行
                nextTxt = Trim(nextTxt)
                If nextTxt <> "" Then
                    ' 在当前段落末尾插入下一个段落内容
                    Dim rngEnd As Range
                    Set rngEnd = para.Range.Duplicate
                    rngEnd.Collapse wdCollapseEnd
                    rngEnd.InsertAfter nextTxt
                    ' 删除下一段
                    nextPara.Range.Delete
                End If
            End If
            ' 格式化“摘要”二字
            With para.Range.Characters(1).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            With para.Range.Characters(2).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            ' 格式化冒号
            If para.Range.Characters.Count >= 3 Then
                With para.Range.Characters(3).Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = True
                    .Color = wdColorBlack
                End With
            End If
            ' 段落格式
            para.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            para.Range.ParagraphFormat.FirstLineIndent = 24
        End If
    Next para
    MsgBox "摘要标题及内容合并完成！"
End Sub

Sub MergeAbstractParagraph()
    Dim para As Paragraph
    Dim nextPara As Paragraph
    Dim txt As String
    For Each para In ActiveDocument.Paragraphs
        txt = Trim(para.Range.Text)
        If Left(txt, 2) = "摘要" Then
            ' 检查第三个字符是否为冒号（全角或半角），没有则插入
            Dim thirdChar As String
            If para.Range.Characters.Count >= 3 Then
                thirdChar = para.Range.Characters(3)
            Else
                thirdChar = ""
            End If
            If thirdChar <> "：" And thirdChar <> ":" Then
                para.Range.Characters(2).InsertAfter "："
            End If
            ' 如果当前段落后还有内容段落，则合并
            Set nextPara = para.Next
            If Not nextPara Is Nothing Then
                Dim nextTxt As String
                nextTxt = Trim(nextPara.Range.Text)
                If nextTxt <> "" Then
                    ' 在当前段落末尾插入下一个段落内容
                    para.Range.Collapse wdCollapseEnd
                    para.Range.InsertAfter nextTxt
                    ' 删除下一段
                    nextPara.Range.Delete
                End If
            End If
            ' 格式化“摘要”二字
            With para.Range.Characters(1).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            With para.Range.Characters(2).Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            ' 格式化冒号
            If para.Range.Characters.Count >= 3 Then
                With para.Range.Characters(3).Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = True
                    .Color = wdColorBlack
                End With
            End If
            ' 段落格式
            para.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
            para.Range.ParagraphFormat.FirstLineIndent = 24
        End If
    Next para
    MsgBox "摘要标题及内容合并完成！"
End Sub

