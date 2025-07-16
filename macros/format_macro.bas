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
                    para.Range.Characters(2).InsertAfter "："
                ElseIf txt = "关键词" Or Left(txt, 4) = "关键词：" Then
                    para.Range.Characters(3).InsertAfter "："
                ElseIf txt = "Abstract" Or Left(txt, 9) = "Abstract:" Then
                    para.Range.Characters(8).InsertAfter ":"
                ElseIf txt = "Keywords" Or Left(txt, 9) = "Keywords:" Then
                    para.Range.Characters(8).InsertAfter ":"
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

