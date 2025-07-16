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
        ' Compact样式：字体与正文一样，但无缩进
        ElseIf para.Style = "Compact" Then
            With para.Range
                .Font.NameFarEast = "宋体"
                .Font.Name = "Times New Roman"
                .Font.Size = 12 ' 小四
                .Font.Bold = False
                .Font.Color = wdColorBlack
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



' 参考文献格式化宏
Sub FormatReferences()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim pageBreakAdded As Boolean
    
    pageBreakAdded = False
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 查找参考文献标题
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            ' 只在第一次找到时添加分页符
            If Not pageBreakAdded Then
                ' 在参考文献标题前添加分页符
                Dim refRange As Range
                Set refRange = para.Range.Duplicate
                refRange.Collapse wdCollapseStart
                refRange.InsertBreak Type:=wdPageBreak
                pageBreakAdded = True
            End If
            
            ' 先设置样式，再格式化
            On Error Resume Next
            para.Style = ActiveDocument.Styles("标题 1")
            If Err.Number <> 0 Then
                ' 如果标题1样式不存在，尝试使用默认标题样式
                para.Style = ActiveDocument.Styles("Heading 1")
            End If
            On Error GoTo 0
            
            ' 强制应用格式（覆盖样式）
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
                .LeftIndent = 0
                .RightIndent = 0
            End With
            
        End If
    Next i
    
    MsgBox "参考文献标题格式化完成！"
End Sub

' 格式化参考文献条目
Sub FormatReferenceEntries()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim foundReferences As Boolean
    Dim referenceCount As Integer
    
    foundReferences = False
    referenceCount = 0
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 检查是否到达参考文献部分
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            foundReferences = True
            GoTo NextPara
        End If
        
        ' 检查是否到达其他章节（结束参考文献部分）
        If foundReferences And (txt = "附录" Or txt = "Appendix" Or _
           Left(txt, 3) = "图 " Or Left(txt, 3) = "表 " Or _
           Left(txt, 4) = "Figure" Or Left(txt, 4) = "Table" Or _
           Left(txt, 5) = "致谢" Or Left(txt, 5) = "Acknowledgments" Or _
           Left(txt, 6) = "作者简介" Or Left(txt, 6) = "Author Bio") Then
            foundReferences = False
            GoTo NextPara
        End If
        
        ' 通用判断：检查是否遇到下一个标题样式（结束参考文献部分）
        If foundReferences And (para.Style = "标题 1" Or para.Style = "标题 2" Or para.Style = "标题 3" Or _
           para.Style = "Heading 1" Or para.Style = "Heading 2" Or para.Style = "Heading 3") Then
            ' 在参考文献部分结束后添加分页符
            Dim endRange As Range
            Set endRange = para.Range.Duplicate
            endRange.Collapse wdCollapseStart
            endRange.InsertBreak Type:=wdPageBreak
            foundReferences = False
            GoTo NextPara
        End If
        
        ' 如果在参考文献部分，格式化条目
        If foundReferences Then
            ' 跳过空行
            If Len(txt) = 0 Then
                GoTo NextPara
            End If
            
            ' 检查是否为参考文献条目（不是空行且不是标题，且不是其他章节标题）
            If Len(txt) > 0 And txt <> "参考文献" And txt <> "References" And _
               Left(txt, 5) <> "参考文献：" And Left(txt, 11) <> "References:" And _
               Left(txt, 3) <> "图 " And Left(txt, 3) <> "表 " And _
               Left(txt, 4) <> "Figure" And Left(txt, 4) <> "Table" And _
               Left(txt, 5) <> "致谢" And Left(txt, 5) <> "Acknowledgments" And _
               Left(txt, 6) <> "作者简介" And Left(txt, 6) <> "Author Bio" And _
               txt <> "附录" And txt <> "Appendix" And _
               para.Style <> "标题 1" And para.Style <> "标题 2" And para.Style <> "标题 3" And _
               para.Style <> "Heading 1" And para.Style <> "Heading 2" And para.Style <> "Heading 3" Then
                referenceCount = referenceCount + 1
                
                ' 先设置段落格式（悬挂缩进）
                With para.Range.ParagraphFormat
                    .Alignment = wdAlignParagraphLeft
                    .FirstLineIndent = -36 ' 首行缩进为负值，实现悬挂缩进（APA标准：0.5英寸）
                    .LeftIndent = 36 ' 左缩进0.5英寸（APA标准）
                    .LineSpacingRule = wdLineSpace1pt5
                End With
                
                ' 再设置字体格式
                With para.Range.Font
                    .NameFarEast = "宋体"
                    .Name = "Times New Roman"
                    .Size = 12 ' 小四
                    .Bold = False
                    .Color = wdColorBlack
                End With
            End If
        End If
        
NextPara:
    Next i
    
    MsgBox "参考文献条目格式化完成！共处理 " & referenceCount & " 个条目。"
End Sub

' 自动编号参考文献
Sub AutoNumberReferences()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim foundReferences As Boolean
    Dim referenceCount As Integer
    Dim newText As String
    
    foundReferences = False
    referenceCount = 0
    
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 检查是否到达参考文献部分
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            foundReferences = True
            GoTo NextPara2
        End If
        
        ' 如果在参考文献部分，处理条目
        If foundReferences Then
            ' 跳过空行
            If Len(txt) = 0 Then
                GoTo NextPara2
            End If
            
            ' 检查是否为参考文献条目（不以数字开头，且不是标题）
            If Not IsNumeric(Left(txt, 1)) And Left(txt, 1) <> "[" And Left(txt, 1) <> "(" And _
               txt <> "参考文献" And txt <> "References" And _
               Left(txt, 5) <> "参考文献：" And Left(txt, 11) <> "References:" Then
                
                referenceCount = referenceCount + 1
                newText = "[" & referenceCount & "] " & txt
                
                ' 替换段落内容
                para.Range.Text = newText & vbCr
            End If
        End If
        
NextPara2:
    Next i
    
    MsgBox "参考文献自动编号完成！共编号 " & referenceCount & " 个条目。"
End Sub

' 完整的参考文献处理宏（APA格式）
Sub ProcessReferences()
    ' 1. 格式化参考文献标题
    FormatReferences
    ' 2. 格式化参考文献条目
    FormatReferenceEntries
    
    MsgBox "参考文献处理完成！"
End Sub

' 参考文献按字母排序宏
Sub SortReferences()
    Dim para As Paragraph
    Dim txt As String
    Dim i As Integer
    Dim foundReferences As Boolean
    Dim referenceCount As Integer
    Dim references() As String
    Dim referenceRanges() As Range
    Dim tempText As String
    Dim tempRange As Range
    Dim j As Integer, k As Integer
    
    foundReferences = False
    referenceCount = 0
    ReDim references(0)
    ReDim referenceRanges(0)
    
    ' 第一步：收集参考文献条目
    For i = 1 To ActiveDocument.Paragraphs.Count
        Set para = ActiveDocument.Paragraphs(i)
        txt = Trim(Replace(para.Range.Text, vbCr, ""))
        
        ' 检查是否到达参考文献部分
        If txt = "参考文献" Or txt = "References" Or _
           Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
            foundReferences = True
            GoTo NextParaSort
        End If
        
        ' 通用判断：检查是否遇到下一个标题样式（结束参考文献部分）
        If foundReferences And (para.Style = "标题 1" Or para.Style = "标题 2" Or para.Style = "标题 3" Or _
           para.Style = "Heading 1" Or para.Style = "Heading 2" Or para.Style = "Heading 3") Then
            foundReferences = False
            GoTo NextParaSort
        End If
        
        ' 如果在参考文献部分，收集条目
        If foundReferences Then
            ' 跳过空行
            If Len(txt) = 0 Then
                GoTo NextParaSort
            End If
            
            ' 检查是否为参考文献条目
            If Len(txt) > 0 And txt <> "参考文献" And txt <> "References" And _
               Left(txt, 5) <> "参考文献：" And Left(txt, 11) <> "References:" And _
               Left(txt, 3) <> "图 " And Left(txt, 3) <> "表 " And _
               Left(txt, 4) <> "Figure" And Left(txt, 4) <> "Table" And _
               Left(txt, 5) <> "致谢" And Left(txt, 5) <> "Acknowledgments" And _
               Left(txt, 6) <> "作者简介" And Left(txt, 6) <> "Author Bio" And _
               txt <> "附录" And txt <> "Appendix" And _
               para.Style <> "标题 1" And para.Style <> "标题 2" And para.Style <> "标题 3" And _
               para.Style <> "Heading 1" And para.Style <> "Heading 2" And para.Style <> "Heading 3" Then
                
                referenceCount = referenceCount + 1
                ReDim Preserve references(referenceCount - 1)
                ReDim Preserve referenceRanges(referenceCount - 1)
                
                references(referenceCount - 1) = txt
                Set referenceRanges(referenceCount - 1) = para.Range.Duplicate
            End If
        End If
        
NextParaSort:
    Next i
    
    ' 第二步：按字母排序（不区分大小写，符合APA格式）
    For j = 0 To referenceCount - 2
        For k = j + 1 To referenceCount - 1
            If LCase(references(j)) > LCase(references(k)) Then
                ' 交换文本
                tempText = references(j)
                references(j) = references(k)
                references(k) = tempText
                
                ' 交换范围
                Set tempRange = referenceRanges(j)
                Set referenceRanges(j) = referenceRanges(k)
                Set referenceRanges(k) = tempRange
            End If
        Next k
    Next j
    
    ' 第三步：重新排列段落
    If referenceCount > 0 Then
        ' 删除所有参考文献条目
        For j = 0 To referenceCount - 1
            referenceRanges(j).Delete
        Next j
        
        ' 找到参考文献标题位置
        Dim insertRange As Range
        For i = 1 To ActiveDocument.Paragraphs.Count
            Set para = ActiveDocument.Paragraphs(i)
            txt = Trim(Replace(para.Range.Text, vbCr, ""))
            
            If txt = "参考文献" Or txt = "References" Or _
               Left(txt, 5) = "参考文献：" Or Left(txt, 11) = "References:" Then
                Set insertRange = para.Range.Duplicate
                insertRange.Collapse wdCollapseEnd
                Exit For
            End If
        Next i
        
        ' 按排序后的顺序插入
        For j = 0 To referenceCount - 1
            insertRange.InsertAfter references(j) & vbCr
            ' 确保插入的段落使用正文样式
            Dim newPara As Paragraph
            Set newPara = insertRange.Paragraphs(insertRange.Paragraphs.Count)
            If Not newPara Is Nothing Then
                On Error Resume Next
                newPara.Style = ActiveDocument.Styles("正文文本")
                If Err.Number <> 0 Then
                    ' 如果正文文本样式不存在，尝试使用默认样式
                    newPara.Style = ActiveDocument.Styles("Normal")
                End If
                On Error GoTo 0
            End If
        Next j
        
        ' 在参考文献部分结束后添加分页符
        Dim lastRefPara As Paragraph
        Set lastRefPara = insertRange.Paragraphs(insertRange.Paragraphs.Count)
        If Not lastRefPara Is Nothing Then
            Dim endPageRange As Range
            Set endPageRange = lastRefPara.Range.Duplicate
            endPageRange.Collapse wdCollapseEnd
            endPageRange.InsertBreak Type:=wdPageBreak
        End If
    End If
    
    MsgBox "参考文献排序完成！共排序 " & referenceCount & " 个条目。"
End Sub

' 完整的参考文献处理宏（包含排序）
Sub ProcessReferencesWithSort()
    ' 1. 格式化参考文献标题
    FormatReferences
    ' 2. 排序参考文献条目
    SortReferences
    ' 3. 格式化参考文献条目
    FormatReferenceEntries
    
    MsgBox "参考文献处理完成（包含排序）！"
End Sub

' 总的一键格式化宏 - 按顺序执行所有格式化步骤
Sub FormatAllDocument()
    Dim response As Integer
    
    ' 询问用户是否继续
    response = MsgBox("即将执行完整的文档格式化，包括：" & vbCrLf & _
                     "1. 页面设置和正文行距" & vbCrLf & _
                     "2. 标题格式化（题目、一级、二级、三级标题）" & vbCrLf & _
                     "3. 正文格式化" & vbCrLf & _
                     "4. 正文中数字和英文字体格式化（已移除）" & vbCrLf & _
                     "5. 摘要和关键词格式化" & vbCrLf & _
                     "6. 目录处理" & vbCrLf & _
                     "7. 参考文献格式化（包含排序）" & vbCrLf & vbCrLf & _
                     "是否继续？", vbYesNo + vbQuestion, "文档格式化")
    
    If response = vbNo Then
        MsgBox "操作已取消。"
        Exit Sub
    End If
    
    ' 开始执行格式化
    Application.ScreenUpdating = False ' 关闭屏幕更新，提高性能
    
    On Error GoTo ErrorHandler
    
    ' 检查文档是否为空
    If ActiveDocument.Paragraphs.Count = 0 Then
        MsgBox "文档为空，无法执行格式化。", vbExclamation, "警告"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' 1. 页面设置和正文行距
    MsgBox "正在设置页面和正文行距..."
    SetPageAndBodyFormat
    
    ' 2. 标题格式化
    MsgBox "正在格式化标题..."
    FormatTitleByHeadingStyle
    FormatLevel1Heading
    FormatLevel2Heading
    FormatLevel3Heading
    
    ' 3. 正文格式化
    MsgBox "正在格式化正文..."
    FormatBodyText
    
    ' 4. 正文中数字和英文字体格式化（已移除，由其他格式化覆盖）
    ' MsgBox "正在格式化正文中的数字和英文字体..."
    ' FormatNumbersAndEnglishInBody
    
    ' 5. 摘要和关键词格式化
    MsgBox "正在格式化摘要和关键词..."
    MergeAndFormatAbstract
    
    ' 6. 目录处理
    MsgBox "正在处理目录..."
    ProcessTableOfContents
    
    ' 7. 参考文献格式化（包含排序）
    MsgBox "正在格式化参考文献..."
    ProcessReferencesWithSort
    
    Application.ScreenUpdating = True ' 恢复屏幕更新
    
    MsgBox "文档格式化完成！" & vbCrLf & vbCrLf & _
           "所有格式化步骤已执行完毕，包括：" & vbCrLf & _
           "✓ 页面设置和行距" & vbCrLf & _
           "✓ 标题格式化" & vbCrLf & _
           "✓ 正文格式化" & vbCrLf & _
           "✓ 数字和英文字体（已移除）" & vbCrLf & _
           "✓ 摘要和关键词" & vbCrLf & _
           "✓ 目录处理" & vbCrLf & _
           "✓ 参考文献格式化（含排序）", vbInformation, "格式化完成"
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True ' 确保恢复屏幕更新
    MsgBox "格式化过程中出现错误：" & vbCrLf & Err.Description, vbCritical, "错误"
End Sub

' 快速格式化宏（不包含排序，适用于已有排序的文档）
Sub FormatAllDocumentQuick()
    Dim response As Integer
    
    ' 询问用户是否继续
    response = MsgBox("即将执行快速文档格式化（不包含排序），包括：" & vbCrLf & _
                     "1. 页面设置和正文行距" & vbCrLf & _
                     "2. 标题格式化" & vbCrLf & _
                     "3. 正文格式化" & vbCrLf & _
                     "4. 数字和英文字体格式化（已移除）" & vbCrLf & _
                     "5. 摘要和关键词格式化" & vbCrLf & _
                     "6. 目录处理" & vbCrLf & _
                     "7. 参考文献格式化（不排序）" & vbCrLf & vbCrLf & _
                     "是否继续？", vbYesNo + vbQuestion, "快速格式化")
    
    If response = vbNo Then
        MsgBox "操作已取消。"
        Exit Sub
    End If
    
    ' 开始执行格式化
    Application.ScreenUpdating = False
    
    On Error GoTo ErrorHandlerQuick
    
    ' 1. 页面设置和正文行距
    SetPageAndBodyFormat
    
    ' 2. 标题格式化
    FormatTitleByHeadingStyle
    FormatLevel1Heading
    FormatLevel2Heading
    FormatLevel3Heading
    
    ' 3. 正文格式化
    FormatBodyText
    
    ' 4. 正文中数字和英文字体格式化（已移除，由其他格式化覆盖）
    ' FormatNumbersAndEnglishInBody
    
    ' 5. 摘要和关键词格式化
    MergeAndFormatAbstract
    
    ' 6. 目录处理
    ProcessTableOfContents
    
    ' 7. 参考文献格式化（不排序）
    ProcessReferences
    
    Application.ScreenUpdating = True
    
    MsgBox "快速格式化完成！", vbInformation, "格式化完成"
    
    Exit Sub

ErrorHandlerQuick:
    Application.ScreenUpdating = True
    MsgBox "格式化过程中出现错误：" & vbCrLf & Err.Description, vbCritical, "错误"
End Sub

