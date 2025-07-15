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
        If txt = "摘要"  Then
            ' 只处理这类段落
            MsgBox "找到摘要段内容: [" & para.Range.Text & "]"
            Set nextPara = para.Next
            If Not nextPara Is Nothing Then
                MsgBox "找到内容段: [" & nextPara.Range.Text & "]"
                contentTxt = nextPara.Range.Text
                contentTxt = Replace(contentTxt, vbCr, "")
                contentTxt = Replace(contentTxt, vbLf, "")
                contentTxt = Trim(contentTxt)
                ' 如果有“关键词”，只取前面部分
                Dim kwPos As Integer
                kwPos = InStr(contentTxt, "关键词")
                If kwPos > 0 Then
                    contentTxt = Left(contentTxt, kwPos - 1)
                End If
                ' 获取 para 段落的最后一个字符（段落符号）前的位置
                Dim rngInsert As Range
                Set rngInsert = para.Range.Duplicate
                rngInsert.End = rngInsert.End - 1  ' 不包括段落符号
                rngInsert.Collapse wdCollapseEnd
                rngInsert.InsertAfter contentTxt
                MsgBox "合并后摘要段内容: [" & para.Range.Text & "]"
                nextPara.Range.Delete
            Else
                MsgBox "未找到内容段"
            End If
            ' 格式化“摘要”二字
            Set rng = para.Range.Characters(1)
            With rng.Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12 ' 小四
                .Bold = True
                .Color = wdColorBlack
            End With
            Set rng = para.Range.Characters(2)
            With rng.Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = True
                .Color = wdColorBlack
            End With
            ' 格式化冒号
            If para.Range.Characters.Count >= 3 Then
                Set rng = para.Range.Characters(3)
                With rng.Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = True
                    .Color = wdColorBlack
                End With
            End If
            ' 格式化内容部分
            Dim rngContent As Range
            If para.Range.Characters.Count > 3 Then
                Set rngContent = para.Range.Duplicate
                rngContent.Start = rngContent.Start + 3 * 2 ' 跳过3个中文字符
                With rngContent.Font
                    .NameFarEast = "宋体"
                    .Name = "宋体"
                    .Size = 12
                    .Bold = False
                    .Color = wdColorBlack
                End With
            End If
            ' 判断“摘要”后是否有冒号，没有则补全角冒号
            Dim paraText As String
            paraText = para.Range.Text
            If Len(paraText) < 3 Or (Mid(paraText, 3, 1) <> "：" And Mid(paraText, 3, 1) <> ":") Then
                para.Range.Characters(2).InsertAfter "："
            End If
            ' 合并后，先将段落样式改为正文文本
            para.Style = ActiveDocument.Styles("正文文本")

            ' 整段统一设置为宋体、小四、黑色、非加粗
            With para.Range.Font
                .NameFarEast = "宋体"
                .Name = "宋体"
                .Size = 12
                .Bold = False
                .Color = wdColorBlack
            End With

            ' 只对“摘要”二字和冒号加粗
            para.Range.Characters(1).Font.Bold = True
            para.Range.Characters(2).Font.Bold = True
            If para.Range.Characters.Count >= 3 Then
                para.Range.Characters(3).Font.Bold = True
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

Sub ShowFirstHeading1Trim()
    Dim para As Paragraph
    Dim txt As String
    For Each para In ActiveDocument.Paragraphs
        If para.Style = "标题1" Or para.Style = "Heading 1" Or para.Style = "标题 1" Then
            txt = para.Range.Text
            MsgBox "原始内容: [" & txt & "]"
            MsgBox "Trim后: [" & Trim(txt) & "]"
            MsgBox "Replace+Trim后: [" & Trim(Replace(txt, vbCr, "")) & "]"
            Exit Sub
        End If
    Next para
    MsgBox "文档中没有找到一级标题（标题1/Heading 1/标题 1）"
End Sub
