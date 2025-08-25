Attribute VB_Name = "CUITMacro"
Global Const ERR_CANCEL = vbObjectError + 1
Global Const ERR_USRMSG = vbObjectError + 2
Global Const C_TITLE = "毕业论文"

Public titleCN As String
Public titleEN As String
Public studentName As String
Public studentNo As String
Public firstTeacherName As String
Public firstTeacherTitle As String
Public otherTeacherName As String
Public otherTeacherTitle As String
Public mathTypeFound As Boolean
Public axMathFound As Boolean

Const Version = "v1.3.15"

Const TEXT_GithubUrl = "https://github.com/sk8boy/cuit_dissertation_template"
Const TEXT_GiteeUrl = "https://gitee.com/tiejunwang/cuit_dissertation_template"

Const BookmarkPrefix = "_"
Const RefBrokenCommentTitle = "$_REFERENCE_BROKEN_COMMENT$"

Const TEXT_AppName = "成都信息工程大学硕士学位论文模板"
Const TEXT_Author = "王铁军 @ 成都信息工程大学 计算机学院"
Const TEXT_Description = "为使用 Word 撰写硕士学位论文的同学提供一个快速上手的模板。"
Const TEXT_VersionPrompt = "版本："
Const TEXT_NonCommecialPrompt = "仅限非商业用途"


Public Sub UpdatePages_RibbonFun(ByVal control As IRibbonControl)
    Dim rng As Range
    Dim startPage As Integer, endPage As Integer
    Dim bodyPageCount As Integer
    Dim keyword As String
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    
    ' 设置搜索条件
    keyword = InputBox(prompt:="正文起始章节标题", title:="请输入正文起始章节的标题", Default:="绪论")
    
    ' 初始化搜索范围
    Set rng = ActiveDocument.Content
    rng.Find.ClearFormatting
    rng.Find.Style = ActiveDocument.Styles("标题 1")
    
    ' 执行搜索（结合关键字和样式）
    With rng.Find
        .text = keyword
        .Forward = True
        .Wrap = wdFindStop
        .Execute
        If .found Then
            ur.StartCustomRecord "更新正文页数"
            
            ' 找到匹配关键字且样式正确的段落
            startPage = rng.Information(wdActiveEndPageNumber)
            
            ' 获取文档总页数
            endPage = ActiveDocument.Content.Information(wdNumberOfPagesInDocument)
            bodyPageCount = endPage - startPage + 1
            
            ' 存储为文档变量
            ActiveDocument.Variables("BodyPageCount").Value = bodyPageCount
            MsgBox "正文页数: 共" & bodyPageCount & "页", vbInformation, C_TITLE
            'ActiveDocument.Fields.Add Range:=Selection.Range, Type:=wdFieldDocVariable, Text:="BodyPageCount"
            UpdateFooterFields
            UpdatePagesInToc
            Application.ScreenRefresh
            ur.EndCustomRecord
        Else
            MsgBox "未找到符合关键字 '" & keyword & "' 且样式为 '" & ActiveDocument.Styles("标题 1") & "' 的段落！", vbExclamation, C_TITLE
        End If
    End With
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "更新论文正文页数时出错: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Private Sub UpdateFooterFields()
    Dim sec As Section
    Dim ftr As HeaderFooter
    
    On Error Resume Next
    ' 遍历所有节
    For Each sec In ActiveDocument.Sections
        ' 更新主页脚
        Set ftr = sec.Footers(wdHeaderFooterPrimary)
        If ftr.LinkToPrevious = False Then
            ftr.Range.Fields.Update
        End If
        
        ' 更新首页页脚（如果不同）
        If sec.PageSetup.DifferentFirstPageHeaderFooter Then
            Set ftr = sec.Footers(wdHeaderFooterFirstPage)
            If ftr.LinkToPrevious = False Then
                ftr.Range.Fields.Update
            End If
        End If
        
        ' 更新奇偶页页脚（如果不同）
        If sec.PageSetup.OddAndEvenPagesHeaderFooter Then
            Set ftr = sec.Footers(wdHeaderFooterEvenPages)
            If ftr.LinkToPrevious = False Then
                ftr.Range.Fields.Update
            End If
        End If
    Next sec
    
    'MsgBox "页脚中的域已更新完成!", vbInformation, C_TITLE
End Sub

Private Sub UpdatePagesInToc()
    Dim fld As Field
    
    On Error Resume Next
    For Each fld In ActiveDocument.Fields
        If fld.Type = wdFieldDocVariable Then
            ' 检查特定的变量名
            If InStr(1, fld.Code.text, "BodyPageCount", vbTextCompare) > 0 Then
                fld.Update
                ' 可以在此处更新变量的值
                ' ActiveDocument.Variables("MyVariable").Value = "新值"
            End If
        End If
    Next fld
    
    'MsgBox "文档变量域更新完成!", vbInformation, C_TITLE
End Sub

Sub GetSEQFields()
    Dim doc As Document
    Dim fld As Field
    Dim i As Integer
    
    Set doc = ActiveDocument
    i = 1
    
    For Each fld In doc.Fields
        If fld.Type = wdFieldSequence Then
            Debug.Print "SEQ字段 #" & i & ": " & fld.Code
            Debug.Print "当前值: " & fld.result
            i = i + 1
        End If
    Next fld
End Sub

' 插入章编号
Public Sub InsertChapterSep(ByVal control As IRibbonControl)
    Dim rng As Range
    Set rng = Selection.Range
    
    If axMathFound Then
        MsgBox "请使用AxMath菜单下的章节分割标记功能，在每章开始处插入章分隔符！", vbExclamation, C_TITLE
    ElseIf mathTypeFound Then
        MsgBox "请使用AxMath菜单下的章&节功能，在每章开始处插入章分隔符！", vbExclamation, C_TITLE
    Else
        rng.Fields.Add rng, wdFieldSequence, "CUITChap \* MERGEFORMAT \h", False
        rng.Collapse wdCollapseEnd
        ActiveDocument.Fields.Update
    End If
    
End Sub


' 辅助函数：获取SEQ字段的当前值
Private Function GetSEQValue(seqIdentifier As String) As Integer
    Dim doc As Document
    Dim fld As Field
    Dim seqValue As Integer
    
    Set doc = ActiveDocument
    seqValue = 0
    
    For Each fld In doc.Fields
        If fld.Type = wdFieldSequence Then
            If InStr(fld.Code, "SEQ " & seqIdentifier) > 0 Then
                seqValue = val(fld.result)
                Exit For
            End If
        End If
    Next fld
    
    GetSEQValue = seqValue
End Function

Private Function CheckSEQExist(seqIdentifier As String) As Boolean
    Dim doc As Document
    Dim fld As Field
    
    Set doc = ActiveDocument
    seqValue = 0
    
    For Each fld In doc.Fields
        If fld.Type = wdFieldSequence Then
            If InStr(fld.Code, "SEQ " & seqIdentifier) > 0 Then
                CheckSEQExist = True
                Exit Function
            End If
        End If
    Next fld
    
    CheckSEQExist = False
End Function

Public Sub InsertPicNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入图编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "图"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 图 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText " "
    If Not ApplyParaStyle("论文图题", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertTblNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入表编号"
    Selection.TypeText "表"
    Set currentRange = Selection.Range
    With ActiveDocument
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 表 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText " "
    If Not ApplyParaStyle("论文表题", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertDefNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入定义编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "定义"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 定义 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText "："
    currentPos = Selection.Range.Start
    If Not ApplyParaStyle("论文定义", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    paraStart = Selection.Paragraphs(1).Range.Start
    Set aRange = ActiveDocument.Range(Start:=paraStart, End:=currentPos)
    aRange.Font.Bold = True
    aRange.Font.NameFarEast = "黑体"
    aRange.Font.NameAscii = "Times New Roman"
    
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertTheoremNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    Dim paraStart As Long
    Dim currentPos As Long
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入定理编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "定理"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 定理 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText "："
    currentPos = Selection.Range.Start
    If Not ApplyParaStyle("论文定义", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    paraStart = Selection.Paragraphs(1).Range.Start
    Set aRange = ActiveDocument.Range(Start:=paraStart, End:=currentPos)
    aRange.Font.Bold = True
    aRange.Font.NameFarEast = "黑体"
    aRange.Font.NameAscii = "Times New Roman"
    
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertCorollaryNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    Dim paraStart As Long
    Dim currentPos As Long
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入推论编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "推论"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 推论 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText "："
    currentPos = Selection.Range.Start
    If Not ApplyParaStyle("论文定义", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    paraStart = Selection.Paragraphs(1).Range.Start
    Set aRange = ActiveDocument.Range(Start:=paraStart, End:=currentPos)
    aRange.Font.Bold = True
    aRange.Font.NameFarEast = "黑体"
    aRange.Font.NameAscii = "Times New Roman"
    
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertLemmaNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    Dim paraStart As Long
    Dim currentPos As Long
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入引理编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "引理"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 引理 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText "："
    currentPos = Selection.Range.Start
    If Not ApplyParaStyle("论文定义", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    paraStart = Selection.Paragraphs(1).Range.Start
    Set aRange = ActiveDocument.Range(Start:=paraStart, End:=currentPos)
    aRange.Font.Bold = True
    aRange.Font.NameFarEast = "黑体"
    aRange.Font.NameAscii = "Times New Roman"
    
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertProblemNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    Dim paraStart As Long
    Dim currentPos As Long
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入问题编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "问题"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 问题 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText "："
    currentPos = Selection.Range.Start
    If Not ApplyParaStyle("论文定义", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    paraStart = Selection.Paragraphs(1).Range.Start
    Set aRange = ActiveDocument.Range(Start:=paraStart, End:=currentPos)
    aRange.Font.Bold = True
    aRange.Font.NameFarEast = "黑体"
    aRange.Font.NameAscii = "Times New Roman"
    
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertConclusionNo_RibbonFun(ByVal control As IRibbonControl)
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    Dim paraStart As Long
    Dim currentPos As Long
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入结论编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "结论"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 结论 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText "："
    currentPos = Selection.Range.Start
    If Not ApplyParaStyle("论文定义", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    paraStart = Selection.Paragraphs(1).Range.Start
    Set aRange = ActiveDocument.Range(Start:=paraStart, End:=currentPos)
    aRange.Font.Bold = True
    aRange.Font.NameFarEast = "黑体"
    aRange.Font.NameAscii = "Times New Roman"
    
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertAlgorithmTbl_RibbonFun(ByVal control As IRibbonControl)
    Dim tbl As Table
    Dim rng As Range
    Dim i As Integer
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入算法"
    
    Set rng = Selection.Range
    rng.Collapse Direction:=wdCollapseEnd
    
    InsertAlgorithmNo
    
    Set rng = Selection.Paragraphs(1).Range
    rng.Collapse Direction:=wdCollapseEnd
    rng.InsertParagraphAfter
    
    ' 插入表格（3行 x 2列）
    Set tbl = ActiveDocument.Tables.Add(rng, 3, 2)
    
    ' 设置表格样式和列宽
    With tbl
        ' 表格宽度为100%页面
        .PreferredWidthType = wdPreferredWidthPercent
        .PreferredWidth = 100
        .Columns(1).PreferredWidth = 10
        .Columns(2).PreferredWidth = 90
        ' 边框样式
        .Borders.Enable = True
        .Borders.InsideLineStyle = wdLineStyleSingle
        .Borders.OutsideLineStyle = wdLineStyleSingle
    End With
    
    ' 填充表格内容
    With tbl
        ' 第一行：输入
        .Cell(1, 1).Range.Style = "论文表格文字"
        .Cell(1, 1).Range.text = "输入"
        .Cell(1, 1).Range.Bold = True
        .Cell(1, 1).VerticalAlignment = wdCellAlignVerticalCenter
        .Cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(1, 2).Range.Style = "论文表格文字"
        
        ' 第二行：输出
        .Cell(2, 1).Range.Style = "论文表格文字"
        .Cell(2, 1).Range.text = "输出"
        .Cell(2, 1).Range.Bold = True
        .Cell(2, 1).VerticalAlignment = wdCellAlignVerticalCenter
        .Cell(2, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(2, 2).Range.Style = "论文表格文字"
        
        ' 第三行：伪代码
        .Cell(3, 1).Range.Style = "论文表格文字"
        .Cell(3, 1).Range.text = "伪代码"
        .Cell(3, 1).Range.Bold = True
        .Cell(3, 1).VerticalAlignment = wdCellAlignVerticalCenter
        .Cell(3, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Cell(3, 2).Range.Style = "论文表格文字"
    End With
    
    ' 设置字体和缩进（伪代码部分）
    With tbl.Cell(3, 2).Range
        .ParagraphFormat.SpaceAfter = 0
        .ParagraphFormat.FirstLineIndent = 0
        .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
        .ParagraphFormat.LineSpacing = 12
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Font.NameFarEast = "宋体"
        .Font.NameAscii = "Courier New"
        .Font.Bold = False
        .Font.Size = 9
    End With
    ur.EndCustomRecord
    Exit Sub
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub InsertAlgorithmNo_RibbonFun(ByVal control As IRibbonControl)
    InsertAlgorithmNo
End Sub

Private Sub InsertAlgorithmNo()
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    Dim chapNum As Integer
    Dim ur As UndoRecord
    Dim paraStart As Long
    Dim currentPos As Long
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入算法编号"
    With ActiveDocument
        ' 获取当前章编号
        Selection.TypeText "算法"
        Set currentRange = Selection.Range
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "SEQ 算法 \* ARABIC \s 1", False)
        Set aRange = .Range(currentRange.End, currentRange.End)
        aRange.text = "-"
        Set bField = aRange.Fields.Add(currentRange, wdFieldEmpty, "QUOTE ""一九一一年一月日"" \@""D""", False)
        Set aRange = .Range(bField.Code.End - 9, bField.Code.End - 9)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "STYLEREF ""标题 1"" \s", False)
    End With
    Selection.TypeText " "
    If Not ApplyParaStyle("论文算法标题", 0, False) Then Err.Raise ERR_CANCEL
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub ShowInfoDialog_RibbonFun(ByVal control As IRibbonControl)
    BaseInfoForm.Show
End Sub


Public Sub H1_RibbonFun(control As IRibbonControl)
    ' Applies the "heading1" style To max 1 paragraph
    ' Remove SpaceBefore from a directly following "heading2"
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用标题1样式"
    If Not ApplyParaStyle("标题 1", 0, False) Then Err.Raise ERR_CANCEL
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用标题1样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub H2_RibbonFun(control As IRibbonControl)
    ' Applies the "heading2" style To max 1 paragraph
    ' Remove space before, If H2 directly follows H1
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用标题2样式"
    If Not ApplyParaStyle("标题 2", 0, False) Then Err.Raise ERR_CANCEL
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用标题2样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub H3_RibbonFun(control As IRibbonControl)
    ' Applies the built-in Heading 3 style To max 1 paragraph
    ' Remove vertical space before paragraph If H3 directly follows H2 Or H1
    Dim ur As UndoRecord
    Dim SaveRange As Range
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用标题3样式"
    Set SaveRange = Selection.Range
    'Apply the built-in Heading 3 style (paragraph style)
    If Not ApplyParaStyle("标题 3", 0, False) Then Err.Raise ERR_CANCEL
    SaveRange.Select
    With Selection
        'Remove space before, If H3 directly follows H2 Or H1
        If Not .Paragraphs(1).Previous Is Nothing Then
            If (.Paragraphs(1).Previous.Style = "标题 2") Or (.Paragraphs(1).Previous.Style = "标题 1") Then
                .Paragraphs(1).SpaceBefore = 0
            End If
        End If
    End With
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用标题3样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub H4_RibbonFun(control As IRibbonControl)
    ' Applies the built-in Heading 4 style To max 1 paragraph
    ' Remove vertical space before paragraph If H4 directly follows H3, H2, Or H1
    Dim ur As UndoRecord
    Dim SaveRange As Range
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用标题4样式"
    Set SaveRange = Selection.Range
    'Apply the built-in Heading 4 style (paragraph style)
    If Not ApplyParaStyle("标题 4", 0, False) Then Err.Raise ERR_CANCEL
    SaveRange.Select
    With Selection
        'Remove space before, If H4 directly follows H3, H2 Or H1
        If Not .Paragraphs(1).Previous Is Nothing Then
            If (.Paragraphs(1).Previous.Style = "标题 3") Or (.Paragraphs(1).Previous.Style = "标题 2") Or (.Paragraphs(1).Previous.Style = "标题 1") Then
                .Paragraphs(1).SpaceBefore = 0
            End If
        End If
    End With
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用标题4样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub H5_RibbonFun(control As IRibbonControl)
    ' Applies the built-in Heading 5 style To max 1 paragraph
    ' Remove vertical space before paragraph If H4 directly follows H3, H2, Or H1
    Dim ur As UndoRecord
    Dim SaveRange As Range
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用标题5样式"
    Set SaveRange = Selection.Range
    'Apply the built-in Heading 5 style (paragraph style)
    If Not ApplyParaStyle("标题 5", 0, False) Then Err.Raise ERR_CANCEL
    SaveRange.Select
    With Selection
        'Remove space before, If H5 directly follows H4, H3, H2 Or H1
        If Not .Paragraphs(1).Previous Is Nothing Then
            If (.Paragraphs(1).Previous.Style = "标题 4") Or (.Paragraphs(1).Previous.Style = "标题 3") Or (.Paragraphs(1).Previous.Style = "标题 2") Or (.Paragraphs(1).Previous.Style = "标题 1") Then
                .Paragraphs(1).SpaceBefore = 0
            End If
        End If
    End With
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用标题5样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub H6_RibbonFun(control As IRibbonControl)
    ' Applies the built-in Heading 6 style To max 1 paragraph
    ' Remove vertical space before paragraph If H4 directly follows H3, H2, Or H1
    Dim ur As UndoRecord
    Dim SaveRange As Range
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用标题6样式"
    Set SaveRange = Selection.Range
    'Apply the built-in Heading 6 style (paragraph style)
    If Not ApplyParaStyle("标题 6", 0, False) Then Err.Raise ERR_CANCEL
    SaveRange.Select
    With Selection
        'Remove space before, If H5 directly follows H5, H4, H3, H2 Or H1
        If Not .Paragraphs(1).Previous Is Nothing Then
            If (.Paragraphs(1).Previous.Style = "标题 5") Or (.Paragraphs(1).Previous.Style = "标题 4") Or (.Paragraphs(1).Previous.Style = "标题 3") Or (.Paragraphs(1).Previous.Style = "标题 2") Or (.Paragraphs(1).Previous.Style = "标题 1") Then
                .Paragraphs(1).SpaceBefore = 0
            End If
        End If
    End With
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用标题1样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub MakeBulletItem_RibbonFun(control As IRibbonControl)
    'Applies the "论文无序列表" style
    Dim ur As UndoRecord
    Dim objLastSelPara As Paragraph
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用无序列表样式"
    Set objLastSelPara = Selection.Paragraphs(Selection.Paragraphs.Count)
    'Apply the "bulletitem" style
    If Not ApplyParaStyle("论文无序列表", 0, True) Then Err.Raise ERR_CANCEL
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用论文无序列表时出错: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub MakeNumNoIndentItem_RibbonFun(control As IRibbonControl)
    'Applies the "论文无缩序号" style
    Dim ur As UndoRecord
    Dim objLastSelPara As Paragraph
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用无缩序号样式"
    Set objLastSelPara = Selection.Paragraphs(Selection.Paragraphs.Count)
    'Apply the "dashitem" style
    If Not ApplyParaStyle("论文无缩序号", 0, True) Then Err.Raise ERR_CANCEL
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用论文无缩序号时出错: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub MakeNumItem_RibbonFun(control As IRibbonControl)
    'Applies the "论文有序列表" style
    'Adjust the indent To the number of "numitem" paragraph in the current group
    Dim ur As UndoRecord
    Dim objLastNumPara As Paragraph
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用有序列表样式"
    'Save the first And last paragraph
    Set objLastNumPara = Selection.Paragraphs(Selection.Paragraphs.Count)
    'Apply the "numitem" style
    If Not ApplyParaStyle("论文有序列表", 0, True) Then Err.Raise ERR_CANCEL
    '    CalcNumIndent Selection.Paragraphs(1), Selection.Paragraphs(Selection.Paragraphs.Count)
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用论文有序列表时出错: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Private Sub CalcNumIndent(ByRef objFirstNumPara As Paragraph, ByRef objLastNumPara As Paragraph)
    Dim objRange As Range
    
    'Check For valid input
    If objFirstNumPara Is Nothing Then Exit Sub
    If objLastNumPara Is Nothing Then Exit Sub
    'Only calculate indents For list level 1
    If ActiveDocument.Range(objFirstNumPara.Range.Start, objLastNumPara.Range.End).ListFormat.ListLevelNumber <> 1 Then Exit Sub
    'Check For "论文有序列表" paragraphs before the current selection
    Do
        If objFirstNumPara.Previous Is Nothing Then Exit Do
        If objFirstNumPara.Previous.Style <> "论文有序列表" Then Exit Do
        If objFirstNumPara.Previous.Range.ListFormat.ListLevelNumber <> 1 Then Exit Do
        Set objFirstNumPara = objFirstNumPara.Previous
    Loop
    'Check For "论文有序列表" paragraphs after the current selection
    Do
        If objLastNumPara.Next Is Nothing Then Exit Do
        If objLastNumPara.Next.Range.ListFormat.ListLevelNumber <> 1 Then Exit Do
        If objLastNumPara.Next.Style <> "论文有序列表" Then Exit Do
        Set objLastNumPara = objLastNumPara.Next
    Loop
    With ActiveDocument.Range(objFirstNumPara.Range.Start, objLastNumPara.Range.End)
        Select Case .Paragraphs.Count
            Case Is < 10
                .ParagraphFormat.LeftIndent = CentimetersToPoints(0.4)
                .ParagraphFormat.FirstLineIndent = -CentimetersToPoints(0.1)
            Case Is < 100
                .ParagraphFormat.LeftIndent = CentimetersToPoints(0.55)
                .ParagraphFormat.FirstLineIndent = -CentimetersToPoints(0.1)
            Case Else
                .ParagraphFormat.LeftIndent = CentimetersToPoints(0.7)
                .ParagraphFormat.FirstLineIndent = -CentimetersToPoints(0.1)
        End Select
    End With
End Sub

Public Sub ListLevelUp_RibbonFun(control As IRibbonControl)
    'Increases the current list level, i.e. the indentation
    'Only available For lists
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "提升列表级别"
    If Selection.Style Is Nothing Then
        Err.Raise ERR_USRMSG, , "只能选择相同样式的段落!"
    End If
    Select Case Selection.ParagraphFormat.Style
        Case "论文有序列表", "论文无序列表", "论文无缩序号"
            If Selection.Range.ListFormat.ListLevelNumber > 9 Then
                Err.Raise ERR_USRMSG, , "只能选择相同样式的段落!"
            ElseIf Selection.Range.ListFormat.ListLevelNumber > 6 Then
                Err.Raise ERR_USRMSG, , "已经达到了列表的最大级别!"
            End If
            Selection.Range.ListFormat.ListLevelNumber = Selection.Range.ListFormat.ListLevelNumber + 1
        Case Else
            Err.Raise ERR_USRMSG, , "该功能仅对无序列表、有序列表、无所列表有效!"
    End Select
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "尝试提升列表级别时出错: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub ListLevelDown_RibbonFun(control As IRibbonControl)
    'Decreases the current list level, i.e. the indentation
    'Only available For lists
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "降低列表级别"
    If Selection.Style Is Nothing Then
        Err.Raise ERR_USRMSG, , "只能选择相同样式的段落!"
    End If
    Select Case Selection.ParagraphFormat.Style
        Case "论文无序列表", "论文无缩序号"
            If Selection.Range.ListFormat.ListLevelNumber < 2 Then
                Err.Raise ERR_CANCEL
            End If
            Selection.Range.ListFormat.ListLevelNumber = Selection.Range.ListFormat.ListLevelNumber - 1
        Case "论文有序列表"
            If Selection.Range.ListFormat.ListLevelNumber < 2 Then
                Err.Raise ERR_CANCEL
            End If
            Selection.Range.ListFormat.ListLevelNumber = Selection.Range.ListFormat.ListLevelNumber - 1
            '          CalcNumIndent Selection.Paragraphs(1), Selection.Paragraphs(Selection.Paragraphs.Count)
        Case Else
            Err.Raise ERR_USRMSG, , "该功能仅对无序列表、有序列表、无所列表有效!"
    End Select
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "尝试降低列表级别时出错: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub RestartNumbering_RibbonFun(control As IRibbonControl)
    'Restarts the numbering (If in a numbered list) from the first selected paragraphs
    Dim ur As UndoRecord
    Dim objLF As ListFormat
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "切换序号"
    
    Selection.Collapse wdCollapseStart
    If Selection.Paragraphs.Count < 1 Then Err.Raise ERR_CANCEL
    Set objLF = Selection.Paragraphs(1).Range.ListFormat
    If objLF Is Nothing Then
        Err.Raise ERR_USRMSG, , "该功能仅对自动编号列表有效!"
    ElseIf objLF.ListTemplate Is Nothing Then
        Err.Raise ERR_USRMSG, , "该功能仅对自动编号列表有效!"
    End If
    If objLF.ListValue > 1 Then
        objLF.ApplyListTemplate objLF.ListTemplate, False, wdListApplyToWholeList
    Else
        objLF.ApplyListTemplate objLF.ListTemplate, True, wdListApplyToWholeList
    End If
    '    CalcNumIndent Selection.Paragraphs(1), Selection.Paragraphs(Selection.Paragraphs.Count)
    ur.EndCustomRecord
    Exit Sub
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "切换序号时出错: " & Err.Description & vbCrLf & "(当使用自定义序号时，该功能可能会失效)", vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub RestoreSettings_RibbonFun(control As IRibbonControl)
    RestorePageSetup
    CheckEnsureStyles
End Sub

Public Sub RestorePageSetup()
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    If StdPageSetup Then Exit Sub
    If MsgBox("当前论文的页面尺寸设置不满足模板要求!" & vbCrLf & _
            "是否应用标准模板页面尺寸？", vbExclamation + vbYesNo, C_TITLE) = vbNo Then
        Exit Sub
    End If
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "恢复页面设置"
    With ActiveDocument.PageSetup
        .PageHeight = MillimetersToPoints(297)
        .PageWidth = MillimetersToPoints(210)
        .TopMargin = MillimetersToPoints(25.4)
        .BottomMargin = MillimetersToPoints(25.4)
        .LeftMargin = MillimetersToPoints(31.7)
        .RightMargin = MillimetersToPoints(31.7)
        .HeaderDistance = MillimetersToPoints(15)
        .FooterDistance = MillimetersToPoints(17.5)
        .Orientation = wdOrientPortrait
        .Gutter = 0
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .VerticalAlignment = wdAlignVerticalTop
        .LineNumbering.Active = False
        .SuppressEndnotes = False
        .MirrorMargins = False
        .TwoPagesOnOne = False
    End With
    'Switch on hyphenation
    With ActiveDocument
        .AutoHyphenation = False
        .HyphenateCaps = True
        'Set the hyphenation zone To 20pt, approx. 7mm
        .HyphenationZone = 20
        .ConsecutiveHyphensLimit = 0
    End With
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "检查页面设置时发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Private Function StdPageSetup() As Boolean
    On Error Resume Next
    With ActiveDocument.PageSetup
        If Abs(.PageHeight - MillimetersToPoints(297)) > 1 Then
            Exit Function
        End If
        If Abs(.PageWidth - MillimetersToPoints(210)) > 1 Then
            Exit Function
        End If
        If Abs(.TopMargin - MillimetersToPoints(25.4)) > 1 Then
            Exit Function
        End If
        If Abs(.BottomMargin - MillimetersToPoints(25.4)) > 1 Then
            Exit Function
        End If
        If Abs(.LeftMargin - MillimetersToPoints(31.7)) > 1 Then
            Exit Function
        End If
        If Abs(.RightMargin - MillimetersToPoints(31.7)) > 1 Then
            Exit Function
        End If
        If Abs(.HeaderDistance - MillimetersToPoints(15)) > 1 Then
            Exit Function
        End If
        If Abs(.FooterDistance - MillimetersToPoints(17.5)) > 1 Then
            Exit Function
        End If
        If .Orientation <> wdOrientPortrait Then
            Exit Function
        End If
        If .Gutter <> 0 Then
            Exit Function
        End If
        If .OddAndEvenPagesHeaderFooter Then
            Exit Function
        End If
        If .DifferentFirstPageHeaderFooter Then
            Exit Function
        End If
        If .VerticalAlignment <> wdAlignVerticalTop Then
            Exit Function
        End If
        If .LineNumbering.Active Then
            Exit Function
        End If
        If .SuppressEndnotes Then
            Exit Function
        End If
        If .MirrorMargins Then
            Exit Function
        End If
        If .TwoPagesOnOne Then
            Exit Function
        End If
    End With
    With ActiveDocument
        If .AutoHyphenation Then
            Exit Function
        End If
        If Not .HyphenateCaps Then
            Exit Function
        End If
        'Skip the other hyphenation options, i.e. retain personal settings
    End With
    StdPageSetup = True
End Function

Private Sub CheckEnsureStyles()
    ' Make sure that all styles that are available through the custom ribbon are also present in
    ' this document
    Dim ur As UndoRecord
    Dim objStyle As Style
    Dim i As Integer
    Dim myListTemplate As ListTemplate
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "检查并恢复缺失的样式"
    
    If AddMissingStyle("论文正文", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.CharacterUnitFirstLineIndent = 2
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 12
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文摘要正文", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文摘要正文"
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.CharacterUnitFirstLineIndent = 2
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 12
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文摘要标题", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文摘要正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel1
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 16
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文程序代码", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文程序代码"
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 12
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .ParagraphFormat.TabStops.Add 11.35, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 22.7, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 34, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 45.35, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 56.7, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 68.5, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 79.4, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 90.7, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 102.05, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 113.4, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 124.75, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 136.1, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 147.4, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 158.75, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 170.1, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 181.45, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 192.8, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 204.1, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 215.45, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 226.8, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 238.15, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 249.5, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 260.8, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 272.15, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 283.5, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 294.85, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 306.2, wdAlignTabLeft
            .ParagraphFormat.TabStops.Add 317.5, wdAlignTabLeft
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Courier New"
            .Font.Bold = False
            .Font.Size = 9
            .QuickStyle = True
            With .Shading
                .Texture = wdTextureNone
                .ForegroundPatternColor = wdColorAutomatic
                .BackgroundPatternColor = -603917569
            End With
            With .Borders(wdBorderLeft)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderRight)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderTop)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            With .Borders(wdBorderBottom)
                .LineStyle = wdLineStyleSingle
                .LineWidth = wdLineWidth050pt
                .Color = wdColorAutomatic
            End With
            .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
            With .Borders
                .DistanceFromTop = 1
                .DistanceFromLeft = 4
                .DistanceFromBottom = 1
                .DistanceFromRight = 4
                .Shadow = False
            End With
        End With
        With Options
            .DefaultBorderLineStyle = wdLineStyleSingle
            .DefaultBorderLineWidth = wdLineWidth050pt
            .DefaultBorderColor = wdColorAutomatic
        End With
    End If
    If AddMissingStyle("论文题目", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文题目"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.Hyphenation = False
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 16
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文表格文字", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文表格文字"
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.WidowControl = True
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 10.5
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文定义", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .ParagraphFormat.CharacterUnitFirstLineIndent = 2
            .ParagraphFormat.WidowControl = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceMultiple
            .ParagraphFormat.LineSpacing = Application.LinesToPoints(1.25)
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Font.NameFarEast = "仿宋_GB2312"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 12
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文表题", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文表题"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.Hyphenation = False
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 10.5
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文算法标题", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.Hyphenation = False
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 10.5
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文参考文献标题", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "书目"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel1
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 16
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文封面表格文字", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文封面表格文字"
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "仿宋_GB2312"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 14
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文关键词", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文关键词"
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.WidowControl = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 12
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文图题", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.Hyphenation = False
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 10.5
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文图", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文图题"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 12
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文致谢标题", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel1
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 16
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("论文作者成果", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文作者成果"
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.Alignment = wdAlignParagraphJustify
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 12
            .QuickStyle = True
            ActiveDocument.ListTemplates(1).ListLevels(1).LinkedStyle = "论文作者成果"
        End With
    End If
    If AddMissingStyle("论文作者成果标题", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文作者成果"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel1
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 16
            .QuickStyle = True
        End With
    End If
    If AddMissingStyle("标题 1", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel1
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 16
            .QuickStyle = True
            ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1).LinkedStyle = "标题 1"
        End With
    End If
    If AddMissingStyle("标题 2", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel2
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 14
            .QuickStyle = True
            ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2).LinkedStyle = "标题 2"
        End With
    End If
    If AddMissingStyle("标题 3", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitAfter = 0.5
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel3
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 12
            .QuickStyle = True
            ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3).LinkedStyle = "标题 3"
        End With
    End If
    If AddMissingStyle("标题 4", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .ParagraphFormat.LineUnitBefore = 0.5
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel4
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 12
            .QuickStyle = True
            ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4).LinkedStyle = "标题 4"
        End With
    End If
    If AddMissingStyle("标题 5", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel5
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = True
            .Font.Size = 12
            .QuickStyle = True
            ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5).LinkedStyle = "标题 5"
        End With
    End If
    If AddMissingStyle("标题 6", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .NextParagraphStyle = "论文正文"
            .ParagraphFormat.KeepWithNext = True
            .ParagraphFormat.KeepTogether = True
            .NoSpaceBetweenParagraphsOfSameStyle = True
            .ParagraphFormat.OutlineLevel = wdOutlineLevel6
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.FirstLineIndent = 0
            .ParagraphFormat.Alignment = wdAlignParagraphLeft
            .Font.NameFarEast = "宋体"
            .Font.NameAscii = "Times New Roman"
            .Font.Bold = False
            .Font.Size = 12
            .QuickStyle = True
            ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6).LinkedStyle = "标题 6"
        End With
    End If
    If AddMissingStyle("论文无序列表", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .ParagraphFormat.CharacterUnitFirstLineIndent = 0
            .ParagraphFormat.CharacterUnitLeftIndent = 0
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .Font.Bold = False
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            With ListGalleries(wdBulletGallery).ListTemplates(1).ListLevels(1)
                .NumberFormat = ChrW(61548)
                .TrailingCharacter = wdTrailingTab
                .NumberStyle = wdListNumberStyleBullet
                .NumberPosition = CentimetersToPoints(0.85)
                .Alignment = wdListLevelAlignLeft
                .TextPosition = CentimetersToPoints(1.7)
                .ResetOnHigher = 0
                .StartAt = 1
                .Font.name = "Wingdings"
            End With
            .LinkToListTemplate ListGalleries(wdBulletGallery).ListTemplates(1), 1
        End With
    End If
    If AddMissingStyle("论文无缩序号", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .ParagraphFormat.CharacterUnitFirstLineIndent = 0
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.TabStops.Add 48, wdAlignTabLeft
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .Font.Bold = False
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            With ListGalleries(wdNumberGallery).ListTemplates(7).ListLevels(1)
                .NumberFormat = "(%1)"
                .TrailingCharacter = wdTrailingTab
                .NumberStyle = wdListNumberStyleArabic
                .NumberPosition = CentimetersToPoints(0.85)
                .Alignment = wdListLevelAlignLeft
                .TextPosition = CentimetersToPoints(0)
                '.TabPosition = wdUndefined
                .ResetOnHigher = 0
                .StartAt = 1
                .Font.name = "Times New Roman"
                .LinkedStyle = "论文无缩序号"
            End With
            .LinkToListTemplate ListGalleries(wdNumberGallery).ListTemplates(7), 1
        End With
    End If
    If AddMissingStyle("论文有序列表", wdStyleTypeParagraph, objStyle) Then
        With objStyle
            .BaseStyle = wdStyleNormal
            .ParagraphFormat.CharacterUnitFirstLineIndent = 0
            .ParagraphFormat.CharacterUnitLeftIndent = 0
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0
            .Font.Bold = False
            .ParagraphFormat.LeftIndent = 0
            .ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
            .ParagraphFormat.LineSpacing = 20
            With ListGalleries(wdNumberGallery).ListTemplates(7).ListLevels(1)
                .NumberFormat = "%1."
                .TrailingCharacter = wdTrailingTab
                .NumberStyle = wdListNumberStyleArabic
                .NumberPosition = CentimetersToPoints(0.85)
                .Alignment = wdListLevelAlignLeft
                .TextPosition = CentimetersToPoints(1.7)
                .ResetOnHigher = 0
                .StartAt = 1
                .Font.name = "Times New Roman"
                .LinkedStyle = "论文有序列表"
            End With
            .LinkToListTemplate ListGalleries(wdNumberGallery).ListTemplates(7), 1
        End With
    End If
    
    MsgBox "已经检查了所有的模板样式，并对必要的样式进行了恢复!", vbInformation Or vbOKOnly, C_TITLE
    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "检查并恢复缺失的样式时出错: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Private Sub CheckHeadingListGallery()
    Dim headingLevel As Integer
    Dim listTempl As ListTemplate
    
    ' 设置要检查的标题级别(1-9)
    headingLevel = 1
    
    ' 获取标题级别的列表模板(来自列表库)
    On Error Resume Next
    Set listTempl = ListGalleries(wdOutlineNumberGallery).ListTemplates(headingLevel)
    On Error GoTo 0
    
    If Not listTempl Is Nothing Then
        MsgBox "标题" & headingLevel & "在列表库中关联的列表模板存在"
        ' 可以进一步检查这个模板是否在文档中使用
    Else
        MsgBox "标题" & headingLevel & "在列表库中没有预定义的列表模板"
    End If
End Sub

Private Function StyleExists(ByVal StyleName As String) As Boolean
    Dim objStyle As Style
    
    On Error Resume Next
    'Try To find the style in the document
    Set objStyle = ActiveDocument.Styles(StyleName)
    StyleExists = Not (objStyle Is Nothing)
End Function

Private Function AddMissingStyle(ByVal StyleName As String, ByRef StyleType As WdStyleType, ByRef NewStyle As Style) As Boolean
    Dim i As Long
    
    If Not StyleExists(StyleName) Then
        If StyleType = wdStyleTypeList Then
            'Auto-creation of list styles is Not supported in this version
            Err.Raise ERR_USRMSG, , "列表样式 '" & StyleName & "' 已被删除，无法自动恢复！"
        End If
    Else
        Set NewStyle = ActiveDocument.Styles(StyleName)
        If NewStyle.Type = StyleType Then
            'Style exists And the style type is correct --> Exit
            AddMissingStyle = True
            Exit Function
        ElseIf StyleType = wdStyleTypeList Then
            'Auto-creation of list styles is Not supported in this version
            Err.Raise ERR_USRMSG, , "列表样式 '" & StyleName & "' 已经被修改, 无法自动恢复！"
        Else
            'Style exists, but the style type is incorrect --> rename the existing style
            Do
                'Look For a free name
                If Not StyleExists(StyleName & " backup" & i) Then
                    Exit Do
                End If
                i = i + 1
            Loop
            'Rename the style
            ActiveDocument.Styles(StyleName).NameLocal = StyleName & " backup" & i
        End If
    End If
    'Add a New style As a copy of the normal style
    Set NewStyle = ActiveDocument.Styles.Add(StyleName, StyleType)
    NewStyle.Font = ActiveDocument.Styles(wdStyleNormal).Font
    If StyleType <> wdStyleTypeParagraph Then
        NewStyle.ParagraphFormat = ActiveDocument.Styles(wdStyleNormal).ParagraphFormat
        NewStyle.AutomaticallyUpdate = False
    End If
    AddMissingStyle = Not (NewStyle Is Nothing)
    Exit Function
End Function

Public Sub LoadCuitRibbon_RibbonFun(IRibbon As IRibbonUI)
    On Error Resume Next
    IRibbon.ActivateTab ("CuitTab")
    ' CheckAddIns
End Sub

Private Function ApplyParaStyle(ByVal StyleName As String, ByVal BuiltInStyleID As Integer, ByVal booMultiPara As Boolean) As Boolean
    ' Applies a paragraph style To the current selection
    ' - If the booMultiPara flag is Not active, the Function is cancelled For multi-paragraph selections
    ' - Set the cursor To the beginning of the current paragraph
    Dim objStyle As Style
    
    On Error Resume Next
    If BuiltInStyleID <> 0 Then
        Set objStyle = ActiveDocument.Styles(BuiltInStyleID)
    Else
        Set objStyle = ActiveDocument.Styles(StyleName)
    End If
    On Error GoTo ERROR_HANDLER
    If objStyle Is Nothing Then
        Err.Raise ERR_USRMSG, , "该模版中找不到预定义的段落类型 '" & StyleName & "'." & vbCrLf & _
        "请使用'模板检查恢复'按钮对其进行恢复！"
    End If
    'If objStyle <> "论文正文" Then Exit Function
    With Selection
        'check whether text is highlighted
        If .Start <> .End Then
            'some text is selected
            If (.End > .Paragraphs(1).Range.End) Then
                'multiple paragraphs are selected
                If Not booMultiPara Then
                    'If Not supported, cancel
                    Err.Raise ERR_USRMSG, , "该功能只能应用于一个段落!"
                End If
            End If
        End If
        .ParagraphFormat.Style = objStyle
        'collapse the selection
        .Collapse wdCollapseStart
        'go up, If the cursor is Not at the beginning of the paragraph
        If .Start > .Paragraphs(1).Range.Start Then
            .MoveUp wdParagraph, 1
        End If
    End With
    ApplyParaStyle = True
    Exit Function
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用段落样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
End Function

Private Function ApplyCharStyle(ByVal StyleName As String, ByVal BuiltInStyleID As Integer) As Boolean
    ' Applies a character style To the current selection
    ' If there is no highlighted selection, expand it, Until the Next space of paragraph is found
    Dim objStyle As Style
    
    On Error Resume Next
    If BuiltInStyleID <> 0 Then
        Set objStyle = ActiveDocument.Styles(BuiltInStyleID)
    Else
        Set objStyle = ActiveDocument.Styles(StyleName)
    End If
    On Error GoTo ERROR_HANDLER
    If objStyle Is Nothing Then
        Err.Raise ERR_USRMSG, , "该模版中找不到预定义的字符类型 '" & StyleName & "'." & vbCrLf & _
        "请使用'模板检查恢复'按钮对其进行恢复！"
    End If
    If objStyle <> "论文正文" Then Exit Function
    With Selection
        'If no text is highlighted, expand the selection up To the Next space Or paragraph
        If .Start = .End Then
            .MoveStartUntil " " & vbCrLf, wdBackward
            .MoveEndUntil " " & vbCrLf, wdForward
        End If
        .Style = objStyle
    End With
    ApplyCharStyle = True
    Exit Function
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用字符样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
End Function

Public Sub MakeStandard_RibbonFun(control As IRibbonControl)
    ' 1. Different styles selected -> apply the default paragraph style
    ' 2. The font is Not the paragraph standard -> apply the default character format
    ' 3. Part of a paragraph selection -> apply the default character format
    ' 4. All other cases: apply the default paragraph style, If Not yet present. If present, apply the default character format
    ' The default paragraph format can be "p1a" Or "normal", depending from the context
    Dim ur As UndoRecord
    Dim booApplyCharFormat As Boolean
    Dim objFirstPara As Paragraph
    Dim objRangeSave As Range
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用正文样式"
    If Selection.ParagraphFormat.Style Is Nothing Then
        booApplyCharFormat = False
    ElseIf (Selection.Font.name <> Selection.ParagraphFormat.Style.Font.name) Or _
            (Selection.Font.Italic <> Selection.ParagraphFormat.Style.Font.Italic) Or _
            (Selection.Font.Bold <> Selection.ParagraphFormat.Style.Font.Bold) Then
        booApplyCharFormat = True
    ElseIf Selection.Start = Selection.End Then
        booApplyCharFormat = False
    ElseIf InStr(1, Selection.text, Chr(13)) = 0 Then
        booApplyCharFormat = True
    Else
        booApplyCharFormat = False
    End If
    Set objRangeSave = Selection.Range
    If booApplyCharFormat Then
        ApplyCharStyle "论文正文", 0
    Else
        'NormalSpacing control
        'Separate the first paragraph
        Set objFirstPara = Selection.Paragraphs(1)
        If objFirstPara Is Nothing Then Err.Raise ERR_CANCEL
        'If more than one paragraph is selected, first format the rest of the selection
        If Selection.End > objFirstPara.Range.End Then
            Selection.MoveStart wdParagraph, 1
            ApplyParaStyle "论文正文", 0, True
        End If
        objFirstPara.Range.Select
        If Selection.Style = ActiveDocument.Styles("论文正文").NameLocal Then
            ApplyCharStyle "论文正文", 0
        Else
            ApplyParaStyle "论文正文", 0, True
        End If
    End If
    Application.ScreenRefresh
    ur.EndCustomRecord
    
    objRangeSave.Select
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "应用正文样式时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

Public Sub MakeProgCode_RibbonFun(control As IRibbonControl)
    Dim ur As UndoRecord
    
    On Error Resume Next
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "应用源代码样式"
    ApplyParaStyle "论文程序代码", 0, True
    Application.ScreenRefresh
    ur.EndCustomRecord
End Sub

Sub OpenCrossReferenceDialog()
    '未使用
    With Application.Dialogs(wdDialogInsertCrossReference)
        ' 设置默认参数（可选）
        .ReferenceType = "图"  ' 引用类型：标题、书签、脚注等
        '        SendKeys "%r{down 1}", True
        '        .ReferenceKind = wdNumberRelativeContext
        '        SendKeys "%h", True
        .ReferenceItem = "图1-2 论文结构图"
        .Show
    End With
End Sub

Private Sub CheckAddIns()
    '检查MathType和AxMath是否安装了
    Dim addInTemplate As Template
    
    On Error Resume Next
    mathTypeFound = False
    axMathFound = False
    
    For Each addInTemplate In Application.Templates
        If InStr(1, addInTemplate.name, "AxMath", vbTextCompare) > 0 Then
            axMathFound = True
            Exit For
        ElseIf InStr(1, addInTemplate.name, "MathType", vbTextCompare) > 0 Then
            mathTypeFound = True
            Exit For
        End If
    Next
End Sub

Sub NestedField()
    ' 未使用
    Dim aField As Field, bField As Field
    Dim aRange As Range
    Dim currentRange As Range
    
    Selection.TypeText "图"
    Set currentRange = Selection.Range
    With ActiveDocument
        Set aField = currentRange.Fields.Add(currentRange, wdFieldEmpty, "MACROBUTTON AMMPlaceRM \* MERGEFORMAT -", False) '创建空白域，并写入域代码
        Set aRange = .Range(aField.Code.End - 2, aField.Code.End - 2)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "SEQ AMEqn \h \* MERGEFORMAT", False) '在aRange对象位置插入域
        Set aRange = .Range(aField.Code.End - 2, aField.Code.End - 2)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "SEQ AMChap \c \* Arabic \* MERGEFORMAT", False) '在aRange对象位置插入域
        Set aRange = .Range(aField.Code.End - 1, aField.Code.End)
        Set bField = aRange.Fields.Add(aRange, wdFieldEmpty, "SEQ 表 \* ARABIC \s 1", False) '在aRange对象位置插入域
        aField.Update '更新域，效果等同于显示域结果
    End With
    Selection.TypeText " "
End Sub

Sub ToggleFieldCode()
    ' 未使用
    ActiveDocument.Fields.Update
    ActiveDocument.Fields.ToggleShowCodes
End Sub

Sub FindListTemplateLinkedToHeading()
    ' 未使用
    Dim doc As Document
    Dim listTemp As ListTemplate
    Dim lvl As ListLevel
    Dim targetStyleName As String
    Dim foundTemplate As ListTemplate
    Dim listTempIdx As Integer
    Dim listLvlIdx As Integer
    
    Set doc = ActiveDocument
    listTempIdx = 0
    listLvlIdx = 0
    targetStyleName = "标题 1" ' 替换为需要查找的标题样式名
    
    ' 遍历所有列表模板
    For Each listTemp In doc.ListTemplates
        listTempIdx = listTempIdx + 1
        ' 检查每个列表层级是否链接到目标样式
        listLvlIdx = 0
        For Each lvl In listTemp.ListLevels
            listLvlIdx = listLvlIdx + 1
            If lvl.LinkedStyle = targetStyleName Then
                Set foundTemplate = listTemp
                Exit For
            End If
        Next lvl
        If Not foundTemplate Is Nothing Then Exit For
    Next listTemp
    
    ' 输出结果
    If Not foundTemplate Is Nothing Then
        MsgBox "标题样式 """ & targetStyleName & """ 链接到列表模板: " & foundTemplate.name, vbInformation, C_TITLE
        MsgBox "listTempIdx: " & listTempIdx & "listLvlIdx: " & listLvlIdx, vbInformation, C_TITLE
    Else
        MsgBox "未找到链接到样式 """ & targetStyleName & """ 的列表模板。", vbExclamation, C_TITLE
    End If
End Sub

Sub TestStyle()
    ' 未使用
    Dim StyleName As String
    'StyleName = "论文题目"
    'MsgBox "样式'" & StyleName & "'" & StyleExists(StyleName), vbOKOnly, C_TITLE
    'MsgBox "列表名称：" & ActiveDocument.Styles("标题 1").ListTemplate.name, vbOKOnly, C_TITLE
    '    For Each lt In ListGalleries(wdOutlineNumberGallery).ListTemplates
    '        MsgBox lt.name
    '    Next lt
    
    MsgBox "name: " & ActiveDocument.ListTemplates(1).name, vbInformation, C_TITLE
End Sub

Sub ListAllBookmarks()
    ' 未使用
    Dim bm As bookmark
    Dim i As Integer
    Dim doc As Document
    
    Set doc = ActiveDocument
    i = 1
    
    ' 创建新文档显示结果
    Dim newDoc As Document
    Set newDoc = Documents.Add
    newDoc.Content.text = "文档中共有 " & doc.Bookmarks.Count & " 个书签：" & vbCrLf & vbCrLf
    
    ' 列出所有书签
    For Each bm In doc.Bookmarks
        newDoc.Content.InsertAfter i & ". 书签名称: " & bm.name & vbCrLf
        newDoc.Content.InsertAfter "   位置文本: " & bm.Range.text & vbCrLf & vbCrLf
        i = i + 1
    Next bm
    
    ' 格式化输出文档
    newDoc.Content.Paragraphs.Format.SpaceAfter = 0
    newDoc.Range(0, 0).Select
End Sub

'***** Purpose *************************************************************
'
' Comfortably insert cross references in MSWord
'
'***************************************************************************

'***** Useage **************************************************************
'
' 1) Put the cursor To the location in the document where the
'    crossreference shall be inserted,
'    Then press the keyboard shortcut.
'    A temporary bookmark is inserted
'    (If their display is enabled, grey square brackets will appear).
' 2) Move the cursor To the location To where the crossref shall point.
'    Supported are:
'    * Bookmarks
'        e.g. "[42]" (bibliographic reference)
'    * Headlines
'    * Subtitles of Figures, Abbildungen, Tables, Tabellen, etc.
'        (preferably these should be realised via { SEQ Table} etc.)
'        Examples:
'            - { SEQ Figure}: "Figure 123", "Figure 12-345"
'            - { SEQ Table} : "Table 123", "Table 12-345"
'            - { SEQ Ref}   : "[42]"
'    * More of the above can be configured, see below.
'    Hint: Recommendation For large documents: use the navigation pane
'          (View -> Navigation -> Headlines) To quickly find the
'          destination location.
'    Hint: Cross references To hidden text are Not possible
'    Hint: The macro may fail trying To cross reference To locations that
'          have heavily been edited (deletions / moves) While
'          "track changes" (markup mode) was active.
' 3) Press the keyboard shortcut again.
'    The cursor will jump back To the location of insertion
'    And the crossref will have been inserted. Done!
' 4) Additional Function:
'    Positon the cursor at a cross reference field
'    (If you have configured chained cross reference fields,
'    put the cursor To the last field in the chain).
'    Press the keyboard shortcut.
'    - The field display toggles To the Next configured option,
'      e.g. from "see Chapter 1" To "cf. Introduction".
'    - Subsequently added cross references will use the latest format
'      (persistent Until Word is exited).
'
'    You can configure multiple options on how the cross references
'    shall be inserted,
'       e.g. As "Figure 12" Or "Figure 12 - Overview" etc..
'    See below under "=== Configuration" on how
'    To modify the default configuration.
'    Once configured, you can toggle between the different options
'    one after the other As follows:
'    - put the cursor inside a cross reference field, Or immediately
'      behind it (it is generally recommended To Set <Field shading>
'      To <Always> (see https://wordribbon.tips.net/T006107_Controlling_Field_Shading.html)
'      in order To have fields highlighted by grey background.
'    - press the keyboard shortcut
'    - that field toggles its display To be according To the Next option
'    - the current selection is remembered For subsequent
'      cross reference inserts (persistent Until closure of Word)
'
'***************************************************************************

' ============================================================================================
'
' === Main code / entry point
'
' ============================================================================================
Public Sub InsertCrossReference_RibbonFun(control As IRibbonControl)
    'Private Sub test_InsertCrossReference()
    Dim rng As Range
    
    Call InsertCrossReference_
    
    ' 消除交叉引用的格式，使其和论文正文样式一直并更新
    Set rng = Selection.Range
    If rng.Start > 0 Then
        rng.MoveStart wdCharacter, -1
        If rng.Fields.Count > 0 Then
            rng.Fields(1).Select
            Selection.ClearFormatting
            Selection.Style = ActiveDocument.Styles("论文正文")
            Selection.Collapse Direction:=wdCollapseEnd
        Else
            rng.MoveStart wdCharacter, 1
        End If
    End If
End Sub

Function InsertCrossReference_(Optional isActiveState As Variant)
    ' Preparation:
    ' 1) Make sure, the following References are ticked in the VBA editor:
    '       - Microsoft VBScript Regular Expressions 5.5
    '    How To Do it: https://www.datanumen.com/blogs/add-object-library-reference-vba/
    '    Since 200902, this is no longer necessary.
    ' 2) Put this macro code in a VBA module in your document Or document template.
    '    It is recommended To put it into <normal.dot>,
    '    Then the functionality is available in any document.
    ' 3) Assign a keyboard shortcut To this macro (recommendation: Alt+Q)
    '    This works like this (in Office 2010):
    '      - File -> Options -> Adapt Ribbon -> Keyboard Shortcuts: Modify...
    '      - Select from Categories: Macros
    '      - Select form Macros: [name of the Macro]
    '      - Assign keyboard shortcut...
    ' 4) Alternatively To 3) Or in addition To the shortcut, you can assign this
    '    macro To the ribbon button "Insert -> CrossReference".
    '    However, Then you will Not be able any more To access Word's dialog
    '    For inserting cross references.
    '      To assign this Sub To the ribbon button "Insert -> CrossReference",
    '      just rename this Sub To "InsertCrossReference" (without underscore).
    '      To de-assign, re-rename it To something like
    '      "InsertCrossReference_" (With underscore).
    '
    ' Revision History:
    ' 151204 Beginn der Revision History
    ' 160111 Kann jetzt auch umgehen mit Numerierungen mit Bindestrich ?la "Figure 1-1"
    ' 160112 Jetzt auch Querverweise m鰃lich auf Dokumentenreferenzen ?la "[66]" mit Feld " SEQ Ref "
    ' 160615 Felder werden upgedatet falls n鰐ig
    ' 180710 Support f黵 "Nummeriertes Element"
    ' 181026 Generischerer Code f黵 Figure|Table|Abbildung
    ' 190628 New Function: toggle To insert numeric Or text references ("\r")
    ' 190629 Explanations And UI changed To English
    ' 190705 More complete And better configurable inserts
    ' 190709 Expanded configuration possibilities due To text sequences
    ' 200901 Function <IsInArray()> can cope With empty arrays
    ' 200902 Late binding is used To reference the RegExp-library ()
    ' 201112 Support For the \#0 switch
    
    Static isActive As Boolean ' remember whether we are in insertion mode
    Static cfgPHeadline As Integer ' ptr To current config For Headlines
    Static cfgPBookmark As Integer ' ptr To current config For Bookmarks
    Static cfgPFigureTE As Integer ' ptr To current config For Figures, Tables, ...
    
    Dim paramRefType As Variant ' type of reference (WdReferenceType)
    Dim paramRefKind As Variant ' kind of reference (WdReferenceKind)
    Dim paramRefText As Variant ' content of the field
    Dim novbCrLf As String ' dito, but w/o trailing CrLf
    Dim paramRefRnge As Range                ' range containing the reference
    Dim paramRefReal As String ' which of the three configurations
    
    Dim storeTracking As Variant ' temporarily remember the status of "TrackRevisions"
    Dim prompt As String ' text For msgbox
    Dim Response        As Variant              ' user's response To msgbox
    Dim lastpos As Variant
    Dim retry           As Boolean
    Dim found As Boolean
    Dim Index As Variant
    Dim linktype As Variant
    Dim searchstring As String
    Dim allowed As Boolean
    Dim SEQLettering As String
    Dim SEQCategory As String
    Dim Codetext As String
    
    ' ============================================================================================
    ' === Configuration
    ' (This is the (default) configuration that was used before there was any Preference Management.
    '  We leave this in the code To still be able To run InsertCrossReference
    '  without Preference Management.)
    ' The following defines how the crossreferences are inserted.
    ' You may reconfigure according To your preferences. Or just use the defined defaults.
    '
    ' For a basic understanding, it is helpful To know the hierarchy of configurations:
    '   configurations
    '       options
    '           parts
    '               switches
    '
    ' There are three configurations according To the three types of fields:
    '   cfgHeadline For Headlines
    '   cfgBookmark For Bookmark
    '   cfgFigureTE For Figures/Tables/Equations/...
    '
    ' For each configuration, multiple options can be configured.
    ' These are the options between which you can toggle. Accordingly, a certain reference
    ' would be displayed e.g As
    '    Figure 1 - System overview
    '    Figure 1
    '    System overview
    '    ...
    '
    ' Each option can have multiple parts, where parts are
    '   either Field code sequences          (example: <REF \h \r>)
    '   Or     text       sequences          (example: <see chapter >).
    ' Text sequences must be enclosed in <'> (example: <' - '>).
    ' The <? is used To represent a non-breaking space.
    '
    ' Each Field code sequence can have multiple switches (example: <\h \r>).
    '
    '
    ' When Word is started, the configurations always default To the first option.
    ' When you toggle, you switch To the Next option of the configuration. After the last defined
    ' option, the first reappears. Toggling applies To the selected type of fields only,
    ' thus there are three independent toggles For Headlines, Bookmarks And FigureTEs.
    ' After toggling, the selected option is remembered For subsequent inserts of the
    ' respective type. The memory is persistent Until Word is closed.
    ' The individual options are defined in the configuration string, one after the other,
    ' seperated by the sign <|>.
    '
    ' The meaning of the individual switches is similar To the switches of Word's {REF} field:
    '   Main switches:
    '     <REF>,<R>     element's name
    '     <PAGEREF>,<P> insert pagenumber instead of reference
    '     <' '>         text sequence, allows To combine cross references To things like
    '                   <(see chapter 32 on page 48 BELOW)>
    '   Modifier switches:
    '     <\r>          Number instead of text
    '     <\p>          insert <above> Or <below> (Or whatever it is in your local language)
    '     <\n>          no context                              (Not applicable To cfgFigure)
    '     <\w>          full context                            (Not applicable To cfgFigure)
    '     <\c>          combination of category + number + text (Not applicable To cfgBookmark)
    '     <\h>          insert the cross reference As a hyperlink
    '
    '
    Dim cfgHeadline As String ' configurations For Headlines
    Dim cfgBookmark As String ' configurations For Bookmarks
    Dim cfgFigureTE As String ' configurations For Figures, Tables, ...
    
    ' Configuration For Headlines:
    cfgHeadline = "R \r  |REF |R \r '??R  |'(see chapter 'R \r' on page 'PAGEREF')'|R \r ' on p.?PAGEREF|R \p       "
    '             "number|text|number?皌ext| (see chapter  XX    on page YY       ) |number on p.癤X      |above/below"
    '
    ' Configuration For Bookmarks:
    cfgBookmark = "R    |PAGEREF|R \p       |R  ' (see? R \p    ')'"
    '             "text |pagenr |above/below|text (see癮bove/below) "
    '
    ' Configuration For Figures, Tables, Equations, ...:
    '    cfgFigureTE = "R \r     |R \r    '??R  |R   |P     |R \p       |R \c            |R \#0 "
    '             "Figure xx|Figure xx - desc|desc|pagenr|above/below|Figure xxTabdesc|xx    "
    
    ' Favourite configuration of User1:
    '    cfgHeadline = "R \r|'chapter? R \r|R \r'??R"     ' number | text | number - text
    '    cfgBookmark = "R"                                   ' text   | pagenumber
    cfgFigureTE = "R \r" ' Fig XX | description | combi
    
    ' Here you can define additional default parameters which shall generally be appended:
    ' Here we define
    '   - that cross references shall always be inserted As hyperlinks
    '   - that the /* MERGEFORMAT switch shall be Set
    Dim cfgHeadlineAddDefaults As String ' additional default switches For Headlines
    Dim cfgBookmarkAddDefaults As String ' additional default switches For Bookmarks
    Dim cfgFigureTEAddDefaults As String ' additional default switches For Figures, Tables, ...
    
    cfgHeadlineAddDefaults = "\h \* MERGEFORMAT "
    cfgBookmarkAddDefaults = "\h \* MERGEFORMAT "
    '    cfgFigureTEAddDefaults = "\h \* MERGEFORMAT "
    cfgFigureTEAddDefaults = "\h "
    '
    ' Define here the subtitles that shall be recognised. Add more As you wish:
    Const subtitleTypes = "Figure|Fig.|Abbildung|Abb.|Table|Tab.|Tabelle|Equation|Eq.|Gleichung"
    '
    ' Use regex-Syntax To define how To determine subtitles from headers:
    ' ("? is a special character that will be replaced With the above <subtitleTypes>.)
    Const subtitleRecog = "((^(?)([\s\xa0]+)([-\.\d]+):?([\s\xa0]+)(.*))"
    ' Above example:
    '   To be recognised As a subtitle the string
    '      - must start With one of the keywords in <subtitlTypes>
    '      - be followed by one Or more of (whitespaces Or character xa0=160=&nbsp;)
    '      - be followed by one Or more digits Or dots Or minuses (Or any combination thereof)
    '      - be followed by zero Or one colon
    '      - be followed by one Or more of (whitespaces Or character xa0=160=&nbsp;)
    '      - be followed by zero Or more additional characters
    '
    ' === End of Configuration
    ' ============================================================================================
    
    Dim ur As UndoRecord
    
    ' === Is there a Preference Management?
    ' We want To be able To use this routine With And without a PreferenceMgr, thus:
    Dim tmpVal As String
    Dim obj As Object
    Dim Config As Object
    Set Config = CreateObject("Scripting.Dictionary")
    
    Set obj = Nothing
    On Error Resume Next
    Set obj = UserForms.Add("UF_PreferenceMgr")
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "插入交叉引用"
    
    If obj Is Nothing Then
        ' === There is *no* Preference Management.
        ' === Read hard-coded configuration from above into variables ============================
        '        cfgHeadline = Replace(cfgHeadline, "?, Chr(160))
        '        cfgHeadline = AddDefaults(cfgHeadline, cfgHeadlineAddDefaults)
        '        cfgBookmark = Replace(cfgBookmark, "?, Chr(160))
        '        cfgBookmark = AddDefaults(cfgBookmark, cfgBookmarkAddDefaults)
        '        cfgFigureTE = Replace(cfgFigureTE, "?, Chr(160))
        '        cfgFigureTE = AddDefaults(cfgFigureTE, cfgFigureTEAddDefaults)
        '        cfgAHeadline = Split(CStr(cfgHeadline), "|")
        '        cfgABookmark = Split(CStr(cfgBookmark), "|")
        '        cfgAFigureTE = Split(CStr(cfgFigureTE), "|")
        
        ' === Chapters:
        tmpVal = Replace(cfgHeadline, "?", Chr(160))
        tmpVal = AddDefaults(tmpVal, cfgHeadlineAddDefaults)
        Config("cfgCrRf_Ch_FormatA") = Split(CStr(tmpVal), "|")
        
        ' === Bookmarks:
        tmpVal = Replace(cfgBookmark, "?", Chr(160))
        tmpVal = AddDefaults(tmpVal, cfgBookmarkAddDefaults)
        Config("cfgCrRf_BM_FormatA") = Split(CStr(tmpVal), "|")
        
        ' === Figures, Tables, Equations, ...:
        tmpVal = Replace(cfgFigureTE, "?", Chr(160))
        tmpVal = AddDefaults(tmpVal, cfgFigureTEAddDefaults)
        Config("cfgCrRf_ST_FormatA") = Split(CStr(tmpVal), "|")
        Config("cfgCrRf_ST_KeyWd") = Split(subtitleTypes, "|")
        Config("cfgCrRf_ST_KeyRx") = subtitleRecog
        
    Else
        ' === There *is* Preference Management.
        ' Let him Do his initialisations:
        Call obj.doInit
        
        ' === Read configuration from registry into variables ====================================
        Dim arry() As Variant
        arry = obj.GetConfigValues()
        
        Dim i As Long
        Dim theNam As String
        Dim theVal As String
        Dim varNam As String
        Dim doSplit As Boolean
        Dim withBlanks As Boolean
        For i = 0 To UBound(arry, 2)
            If arry(1, i) = False Then
                MsgBox "Missing registry Setting <" & arry(0, i) & ">. Using default value.", vbOKOnly + vbExclamation, "Registry error"
                Stop ' Not yet implemented
            Else
                theNam = CStr(arry(0, i))
                If theNam Like "*KeyWd" Then
                    withBlanks = False
                Else
                    withBlanks = True
                End If
                theVal = Replace(strPrepare(CStr(arry(1, i)), withBlanks), Chr(13), "|")
                
                Select Case True
                    Case theNam Like "*AddDf"
                        tmpVal = AddDefaults(tmpVal, theVal)
                        varNam = "cfg" & rgex(theNam, "(.*_.*)_", "$1") & "_FormatA"
                        doSplit = True
                    Case theNam Like "*MainS"
                        ' store only temporarily - the real storing is done when the additional defaults follow
                        tmpVal = theVal
                        varNam = ""
                        doSplit = False
                    Case theNam Like "*_KeyWd"
                        tmpVal = theVal
                        varNam = "cfg" & theNam
                        doSplit = True
                    Case Else
                        tmpVal = theVal
                        varNam = "cfg" & theNam
                        doSplit = False
                End Select
                If varNam <> "" Then
                    'Debug.Print
                    'Debug.Print varNam
                    'Debug.Print tmpVal
                    If doSplit = False Then
                        Config(varNam) = tmpVal
                    Else
                        Config(varNam) = Split(CStr(tmpVal), "|")
                    End If
                End If
            End If
        Next
    End If
    
    ActiveWindow.View.ShowFieldCodes = False
    
    'Debug.Print cfgPHeadline
    ' Where To insert the XRef:
    ' ============================================================================================
    ' === Check If we are in Insertion-Mode Or Not ===============================================
    If Not (isActive) Then
        ' ========================================================================================
        ' ===== We are Not in Insertion-Mode!  ==> just remember the position To jump back later
        If ActiveDocument.Bookmarks.Exists("tempforInsert") Then
            ActiveDocument.Bookmarks.item("tempforInsert").Delete
        End If
        
        ' Special Function: If the cursor is inside a wdFieldRef-field, Then
        ' - toggle the display among the configured options
        ' - remember the New status For future inserts.
        Index = CursorInField(Selection.Range) ' would fail, If .View.ShowFieldCodes = True
        If Index <> 0 Then
            ' ====================================================================================
            ' ===== Toggle display:
            Dim myOption As String
            Dim myRefType As Integer
            Dim fText0 As String '
            Dim fText2 As String ' Refnumber
            Dim Element As Variant
            Dim needle As String
            Dim optionstring As String
            Dim idx As Integer
            
            ' == Read And clean the code from the field:
            fText0 = ActiveDocument.Fields(Index).Code ' Original
            fText2 = fText0
            fText2 = Replace(fText2, "PAGE", "")        ' change from PAGEREF To REF
            fText2 = regEx(fText2, "REF\s+(\S+)")       ' Get the reference-name
            needle = Replace(Config("cfgCrRf_ST_KeyRx"), "?", Join(Config("cfgCrRf_ST_KeyWd"), "|"))
            
            Select Case True
                    ' == It is a subtitle:
                Case Left(fText2, 4) = "_Ref" And isSubtitle(fText2, needle)
                    myRefType = wdRefTypeNumberedItem
                    'Debug.Print "Subtitle:", cfgPFigureTE, myOption
                    idx = MultifieldDelete(Config("cfgCrRf_ST_FormatA"), cfgPFigureTE, fText0, Index, needle, True)
                    If idx = -1 Then Exit Function
                    
                    cfgPFigureTE = (idx + 1) Mod (UBound(Config("cfgCrRf_ST_FormatA")) + 1)
                    myOption = Config("cfgCrRf_ST_FormatA")(cfgPFigureTE)
                    Application.StatusBar = "New Cross reference format For Subtitles: <" & myOption & ">."
                    
                    paramRefText = ActiveDocument.Bookmarks(fText2).Range.Paragraphs(1).Range.text
                    ' Call MultifieldDelete(Config("cfgCrRf_ST_FormatA"), cfgPFigureTE, fText0, Index, needle, True)
                    paramRefType = RegExReplace(paramRefText, needle, "$2")
                    found = getXRefIndex(paramRefType, CleanHidden(ActiveDocument.Bookmarks(fText2).Range.Paragraphs(1).Range), Index)
                    Call InsertCrossRefs(1, myOption, paramRefType, Index, , True)
                    
                    ' == It is a headline:
                Case Left(fText2, 4) = "_Ref"
                    myRefType = wdRefTypeHeading
                    'Debug.Print "Headline:", cfgPHeadline, myOption
                    idx = MultifieldDelete(Config("cfgCrRf_Ch_FormatA"), cfgPHeadline, fText0, Index)
                    If idx = -1 Then Exit Function
                    
                    cfgPHeadline = (idx + 1) Mod (UBound(Config("cfgCrRf_Ch_FormatA")) + 1)
                    myOption = Config("cfgCrRf_Ch_FormatA")(cfgPHeadline)
                    Application.StatusBar = "New Cross reference format For Headlines: <" & myOption & ">."
                    Call InsertCrossRefs(2, myOption, myRefType, Index, fText2, True)
                    
                    ' == It is a bookmark:
                Case Else
                    myRefType = wdRefTypeBookmark
                    'Debug.Print "Bookmark:", cfgBookmark, myOption
                    idx = MultifieldDelete(Config("cfgCrRf_BM_FormatA"), cfgPBookmark, fText0, Index)
                    If idx = -1 Then Exit Function
                    
                    cfgPBookmark = (idx + 1) Mod (UBound(Config("cfgCrRf_BM_FormatA")) + 1)
                    myOption = Config("cfgCrRf_BM_FormatA")(cfgPBookmark)
                    Application.StatusBar = "New Cross reference format For Bookmarks: <" & myOption & ">."
                    
                    'debug.print rgex(Trim(fText0), "(REF|PAGEREF)\s+(\S+)", "$2")
                    Call InsertCrossRefs(2, myOption, myRefType, fText2, fText2, True)
            End Select
            Exit Function ' Finished changing the display of the reference.
            
        Else
            ' ====================================================================================
            ' ===== Insert temporary Bookmark:
            ' Remember the current position within the document by putting a bookmark there:
            ActiveDocument.Bookmarks.Add name:="tempforInsert", Range:=Selection.Range
            isActive = True ' remember that we are in Insertion-Mode
            '            Call RibbonControl.setAButtonState("BtnTCrossRef", True)
        End If
        
        ' Stelle, wo die zu referenzierende Stelle ist
    Else
        ' ================================
        ' ===== We ARE in Insertion-Mode! ==> jump back To bookmark And insert the XRef
        '        Call RibbonControl.setAButtonState("BtnTCrossRef", True)    ' Though the user has toggled the button, we still want it To be pressed
        
        ' ===== Find out the type of the element To cross-reference To.
        '       It could be a Headline, Figure, Bookmark, ...
        paramRefType = ""
        Select Case Selection.Paragraphs(1).Range.ListFormat.ListType
            Case wdListSimpleNumbering ' bullet lists, numbered Elements
                paramRefType = wdRefTypeNumberedItem
                paramRefKind = wdNumberRelativeContext
                paramRefText = Selection.Paragraphs(1).Range.ListFormat.ListString & _
                " " & Trim(Selection.Paragraphs(1).Range.text)
                paramRefText = Replace(paramRefText, Chr(13), "")
                found = getXRefIndex(paramRefType, paramRefText, Index)
                
            Case wdListOutlineNumbering ' Headlines
                paramRefType = wdRefTypeHeading
                ' The following two lines of code fix a strange behaviour of Word (Bug?),
                '    see http://www.office-forums.com/threads/word-2007-insertcrossreference-wrong-number.1882212/#post-5869968
                paramRefType = wdRefTypeNumberedItem
                paramRefReal = "Headline"
                paramRefKind = wdNumberRelativeContext
                ' paramRefText = Selection.Paragraphs(1).Range.ListFormat.ListString
                ' Sometimes (probably in documents where things have been deleted With track changes), the command
                '    Selection.Paragraphs(1).Range.ListFormat.ListString
                ' doesn't work correctly. Therefore, we Get the numbering differently:
                Dim oDoc As Document
                Dim oRange As Range
                Set oDoc = ActiveDocument
                Set oRange = oDoc.Range(Start:=Selection.Range.Start, End:=Selection.Range.End)
                'Debug.Print oRange.ListFormat.ListString
                paramRefText = oRange.ListFormat.ListString
                found = getXRefIndex(paramRefType, paramRefText, Index)
                
            Case wdListNoNumbering ' SEQ-numbered items, Bookmarks And Figure/Table/Equation/...
                'paramRefText = Trim(Selection.Paragraphs(1).Range.text)
                Set paramRefRnge = Selection.Paragraphs(1).Range
                paramRefText = Trim(paramRefRnge.text)
                With Selection.Paragraphs(1)
                    ' There could be different fields. We look For the first of type <wdFieldSequence>:
                    For i = 1 To .Range.Fields.Count
                        If .Range.Fields(i).Type = wdFieldSequence Then
                            Exit For
                        End If
                    Next
                    If i > .Range.Fields.Count Then
                        paramRefType = ""
                        found = False
                        GoTo trybookmark
                    End If
                    Codetext = UnCAPS(.Range.Fields(i).Code)
                    If ((Left(Codetext, 8) = " SEQ Ref") And _
                            (.Range.Bookmarks.Count = 1)) Then
                        ' == a) SEQ-numbered item, a bibliographic reference ?la <[32] Jackson, 1939, page 37>:
                        paramRefType = wdRefTypeBookmark
                        paramRefKind = wdContentText
                        paramRefReal = "Bookmark"
                        paramRefText = .Range.Bookmarks(1).name
                        found = getXRefIndex(paramRefType, paramRefText, Index)
                    Else
                        ' Bookmark Or Figure/Table/Equation/...
                        ' Get the Lettering:
                        paramRefRnge.End = paramRefRnge.Fields(i).result.End
                        SEQLettering = Trim(paramRefRnge.text)
                        ' *) Hyphen in something like "Figure 1-2" is strangely chr(30), thus this correction:
                        SEQLettering = Replace(SEQLettering, Chr(30), "-")
                        'SEQLettering = Replace(SEQLettering, Chr(160), "")
                        ' Get the category:
                        Set paramRefRnge = Selection.Paragraphs(1).Range
                        ' Extract the Category, e.g. in " SEQ Fig. \* ARABIC" that is "Fig.":
                        'SEQCategory = Trim(paramRefRnge.Fields(i).Code.Words(3))
                        SEQCategory = regEx(paramRefRnge.Fields(i).Code, "\S+\s+(\S+)")
                        
                        ' Try To insert it As a Figure/Table/...
                        ' == b) Figure/Table/...
                        paramRefReal = "FigureTE"
                        paramRefType = SEQCategory
                        paramRefKind = wdOnlyLabelAndNumber
                        found = getXRefIndex(paramRefType, SEQLettering, Index)
                        
trybookmark:
                        If found = False Then
                            ' OK, it was Not a Figure/Table/Equation/...
                            ' Let's check If we are in a bookmark:
                            
                            ' Bookmarks can overlap. Therefore we need an iteration.
                            ' For user experience, it is best If we Select the innermost bookmark (= the shortest):
                            Dim bname As String
                            Dim bmlen As Variant
                            Dim bmlen2 As Long
                            bmlen = ""
                            For Each Element In Selection.Bookmarks
                                bmlen2 = Len(Element.Range.text)
                                If bmlen2 < bmlen Or bmlen = "" Then
                                    bname = Element.name
                                    bmlen = Len(Element.Range.text)
                                End If
                            Next
                            If bmlen <> "" Then
                                ' == c) bookmark
                                paramRefReal = "Bookmark"
                                paramRefType = wdRefTypeBookmark
                                paramRefText = bname
                                found = getXRefIndex(paramRefType, paramRefText, Index)
                            End If
                        End If
                    End If
                End With
            Case Else ' Everything Else
                ' nothing To Do
        End Select ' Now we know what element it is
        
        ' ===== Check, If we can cross-reference To this element:
cannot:
        If paramRefType = "" Then
            ' Sorry, we cannot...
            prompt = "无法在此处插入交叉引用。" & vbNewLine & "请尝试在其他位置插入交叉引用，或者取消。"
            Response = MsgBox(prompt, 1)
            If Response = vbCancel Then
                Selection.GoTo what:=wdGoToBookmark, name:="tempforInsert"
                If ActiveDocument.Bookmarks.Exists("tempforInsert") Then
                    ActiveDocument.Bookmarks.item("tempforInsert").Delete
                End If
                isActive = False
                '                Call RibbonControl.setAButtonState("BtnTCrossRef", False)
            End If
            GoTo CleanExit
        End If
        
        
        ' ===== Insert the cross-reference:
        retry = False
retryfinding:
        If (found = False) And (retry = False) Then
            ' Refresh, ohne dass es als 膎derung getracked wird:
            storeTracking = ActiveDocument.TrackRevisions
            ActiveDocument.TrackRevisions = False
            Selection.HomeKey Unit:=wdStory
            
            Do ' alle SEQ-Felder abklappern
                lastpos = Selection.End
                Selection.GoTo what:=wdGoToField, name:="SEQ"
                'On Error Resume Next
                Debug.Print "Err.Number = " & Err.Number
                allowed = False
                If paramRefType = wdRefTypeNumberedItem Then
                    allowed = True
                Else
                    If IsInArray(paramRefType, Config("cfgCrRf_ST_KeyWd")) Then
                        allowed = True
                        searchstring = " SEQ " & linktype
                        If Left(Selection.NextField.Code.text, Len(searchstring)) = searchstring Then
                            Selection.Fields.Update
                        End If
                    End If
                End If
                If allowed = False Then
                    novbCrLf = regEx(paramRefText, "([^\n\r]+)")
                    prompt = "无法插入这个交叉引用。" & vbCrLf & _
                        vbCrLf & _
                        "尝试插入指向" & vbCrLf & _
                        "   <" & novbCrLf & ">" & vbCrLf & _
                        "的引用无效，请检查无效的引用信息。" & vbCrLf & _
                        vbCrLf & _
                        "诊断数据:" & vbCrLf & _
                        "   paramRefType = <" & paramRefType & ">" & vbCrLf & _
                        "   paramRefKind = <" & paramRefKind & ">" & vbCrLf & _
                        "   paramRefText = <" & novbCrLf & ">"
                    MsgBox prompt, vbOKOnly, "错误 - 无法插入交叉引用"
                    GoTo CleanExit
                End If
                
            Loop While (lastpos <> Selection.End)
            retry = True
            ActiveDocument.TrackRevisions = storeTracking
            GoTo retryfinding
        End If
        
        ' Jetzt das eigentliche Einf黦en des Querverweises an der urspr黱glichen Stelle:
        Selection.GoTo what:=wdGoToBookmark, name:="tempforInsert"
        If found = True Then
            ' Read the correct array the currently selected options:
            Select Case paramRefReal
                Case "Headline"
                    optionstring = Config("cfgCrRf_Ch_FormatA")(cfgPHeadline)
                    ' paramRefType = Not 1, but 0
                Case "Bookmark"
                    optionstring = Config("cfgCrRf_BM_FormatA")(cfgPBookmark)
                    ' paramRefType = 2
                Case Else
                    optionstring = Config("cfgCrRf_ST_FormatA")(cfgPFigureTE)
                    ' paramRefType = 0
            End Select
            
            Call InsertCrossRefs(1, optionstring, paramRefType, Index)
        Else
            If paramRefText = False Then
                paramRefText = paramRefRnge.text
            End If
            prompt = ""
            prompt = vbCrLf & prompt & "paramRefType = <" & paramRefType & ">" & _
                vbCrLf & prompt & "paramRefKind = <" & paramRefKind & ">" & _
                vbCrLf & prompt & "paramRefText = <" & paramRefText & ">"
            MsgBox prompt, vbOKOnly, "Error - Reference Not found:"
            Stop
        End If
        
        isActive = False
        '        Call RibbonControl.setAButtonState("BtnTCrossRef", False)
        
       On Error Resume Next
        If ActiveDocument.Bookmarks.Exists("tempforInsert") Then
            ActiveDocument.Bookmarks.item("tempforInsert").Delete
        End If
       On Error GoTo 0
    End If 'If Not (isActive) Then
CleanExit:
    isActiveState = CBool(isActive)
    
    ur.EndCustomRecord
    Exit Function
    
ERROR_HANDLER:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "插入交叉引用时发生错误: " & Err.Description, vbCritical, C_TITLE
    End If
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Function

'Sub trial()
'    Dim thing As Variant
'    Dim i As Integer
'    Dim StryRng  As Object
'
'    Debug.Print ActiveDocument.StoryRanges.count
'
'    Debug.Print UBound(ActiveDocument.GetCrossReferenceItems("Figure"))
'    thing = ActiveDocument.GetCrossReferenceItems("Figure")(1)
'
'
'    Dim pRange As Range ' The story range, To Loop through each story in the document
'    Dim sShape As Shape ' For the text boxes, which Word considers shapes
'    Dim strText As String
'
'    For Each pRange In ActiveDocument.StoryRanges    'Loop through all of the stories
'        Debug.Print pRange.StoryType, pRange.storyLength
'        Debug.Print UBound(pRange.GetCrossReferenceItems("Figure"))
'    Next
'
'
'For i = 1 To 12
'    If StryRng Is Nothing Then 'First Section object's Header range
'        Set StryRng = ActiveDocument.StoryRanges.item(1)
'    Else
'        Set StryRng = StryRng.NextStoryRange 'ie. Next Section's Header
'    End If
'    With StryRng
'        Debug.Print i, .StoryType, .storyLength
'    End With
'Next
'    'Debug.Print i, ActiveDocument.StoryRanges.item(i).StoryType, ActiveDocument.StoryRanges(i).storyLength
'
'
'    'Debug.Print UBound(ActiveDocument.StoryRanges.item(wdMainTextStory).GetCrossReferenceItems("Figure"))
'    thing = ActiveDocument.GetCrossReferenceItems("Figure")(1)
'    'thing = ActiveDocument.GetCrossReferenceItems("Figure").
'    Debug.Print ActiveDocument.StoryRanges.Application.ActiveDocument.name
'
'
'    'debug.Print ActiveDocument.StoryRanges.
'    Debug.Print ActiveDocument.StoryRanges.item(wdTextFrameStory).text
'    'Debug.Print UBound(ActiveDocument.StoryRanges.item(wdTextFrameStory).GetCrossReferenceItems(3))
'
'End Sub

' ============================================================================================
'
' === Worker routines
'
' ============================================================================================
Private Function CleanHidden(RangeIn As Range) As String
    Dim Range2 As Range
    Dim Range4 As Range
    Dim thetext As String
    
    'Set Range1 = Selection.Range
    
    ' 1.) Remove all but 1st paragraph
    'Debug.Print RangeIn.Paragraphs.Count
    Set Range2 = RangeIn.Duplicate   ' clone, Not a ptr !
    With Range2.TextRetrievalMode
        .IncludeHiddenText = True ' include it, even If currently hidden
        .IncludeFieldCodes = False
    End With
    '    If Range2.Paragraphs.Count > 1 Then
    '        Debug.Print "More than 1 paragraph!"
    '    End If
    Set Range4 = Range2.Paragraphs(1).Range
    
    ' 2.) Remove hidden text
    Range4.TextRetrievalMode.IncludeHiddenText = False
    
    ' *) Remove that strange hidden character at the end
    thetext = Range4.text
    '    thetext = Left(thetext, Len(thetext) - 1)
    
    CleanHidden = thetext
End Function

Private Function AddDefaults(ByRef thestring, tobeAdded As String) As String
    'AddDefaults = RegExReplace(theString, "(R(EF)?|P(AGEREF)?)", "$1" & " " & tobeAdded)
    ' https://regex101.com/r/QT00K9/1
    AddDefaults = RegExReplace(thestring, "(R(EF)?[^|']*|P(AGEREF)?[^|']*)", "$1" & " " & tobeAdded)
    
    '    theString = theString & "|"
    '    tobeAdded = " " & tobeAdded
    '    AddDefaults = Replace(theString, "|", tobeAdded & "|")
    '    AddDefaults = Left(AddDefaults, Len(AddDefaults) - 1)
End Function

Private Function InsertCrossRefs(mode As Integer, _
        optionstring As String, _
        ByVal paramRefType As Variant, _
        Index As Variant, _
        Optional ByVal refcode As String = "", _
    Optional moveCursor As Boolean = False)
    ' Parameters:
    '   <mode>=0    update by manipulating switches
    '         =1    insert via .InsertCrossReference
    '         =2    insert via .Fields.Add
    '   <optionstring>: current option (possibly With multiple parts And multiple switches)
    '   <paramRefType>: the Type of cross reference:
    '                       2 For bookmarks
    '                       0 For everything Else
    '   <index>       : the index of source in Word's internal table Or
    '                   the name of the bookmark
    '   <refcode>     : the reference's name, e.g. _REF6537428
    
    Dim i As Integer '
    Dim thepart As Variant
    Dim isCode As Boolean
    Dim thePartOld As Variant
    Dim isCodeOld As Boolean
    Dim refcode2 As String
    
    thePartOld = ""
    i = 0
    If Len(optionstring) = 0 Then
        MsgBox "InsertCrossRefs detected a non-valid option: <" & optionstring & ">."
        Exit Function
    End If
    Do
        thePartOld = thepart
        isCodeOld = isCode
        i = i + 1
        ' Get the Next part (there could be multiple...)
        thepart = GetPart(optionstring, i, isCode)
        If thepart = Error Then
            ' We have reached the last part!
            
            ' One thing before we return:
            ' If we got the parameter <moveCursor>
            ' (which will be the Case If we have done a replacement, rather than a New
            ' insert) Then we want the cursor positioned behind the last field, even If
            ' the last part was a text - like this the user can continue To toggle.
            ' Hence we have To move the cursor a bit back:
            If (moveCursor = True) And (Len(thePartOld) > 0) And (isCodeOld = False) Then
                Selection.MoveLeft wdCharacter, Len(thePartOld)
            End If
            Exit Do
        End If
        
        ' If it's a text, insert it
        If isCode = False Then
            Application.Selection.InsertAfter thepart
            Application.Selection.Move wdCharacter, 1
        Else
            ' It is a code sequence:
            ' When we modify With method = <0>, we have received a fieldcode <refcode>.
            ' We use this code To Do the modification.
            ' If there are any additional insertions, these must be done With the .Fields.Add-method.
            ' Therefore, we have To prepare a proper fieldcode For that method.
            Call ReplaceAbbrev(thepart)
            If mode = 2 Then
                ' The complete code must be provided in refcode2. The other params are unused.
                refcode2 = " " & regEx(thepart, "(PAGEREF|REF|P|R)")
                refcode2 = refcode2 & " " & refcode
                refcode2 = refcode2 & " " & rgex(CStr(thepart), "(PAGEREF|REF|P|R)(.*)", "$2")
            Else
                ' we must provide:
                ' 1) paramRefType
                '   nothing To Do
                
                ' 2) index
                '   nothing To Do
                
                ' 3) thePart3: Fieldcode w/o RefNr w/ switches
                '   nothing To Do
                
                ' 4)
                refcode2 = "Not used"
                
            End If
            
            If Insert1CrossRef(mode, paramRefType, Index, thepart, refcode2) = False Then
                Exit Do
            End If
            
            If mode = 0 Then
                ' We have modified the first field.
                ' Any additional fields shall be inserted With the .Fields.Add-method.
                mode = 2
            End If
        End If
        
    Loop While True
End Function

Private Function Insert1CrossRef(mode As Integer, Optional param1 As Variant, _
        Optional param2 As Variant, _
        Optional param3 As Variant, _
        Optional param4 As Variant) As Boolean
    ' Parameter <mode>=0    update by manipulating switches     ==>
    '                 =1    insert via .InsertCrossReference
    '                       ==> param1: wdReferenceKind
    '                       ==> param2: wdReferenceItem / RefNr
    '                       ==> param3: Fieldcode w/o RefNr w/ switches
    '                       ==> param4: Not used
    '                 =2    insert via .Fields.Add
    '                       ==> param1: Not used
    '                       ==> param2: Not used
    '                       ==> param3: Not used
    '                       ==> param4: Fieldcode w/ RefNr w/ switches
    ' Returns: True : upon successful completion
    '          False: when an invalid paramter was detected
    Dim inclHyperlink As Boolean
    Dim inclPosition As Boolean
    Dim param0 As Variant
    Dim myCode As String
    Dim idx As Integer
    
    Select Case mode
        Case 0 ' update by manipulating switches
            With ActiveDocument.Fields(param2)
                If InStr(1, param3, "PAGEREF") Then
                    myCode = "PAGEREF "
                    param3 = Replace(param3, "PAGEREF", "")
                Else
                    myCode = "REF "
                End If
                myCode = myCode & param4 & " " & param3
                .Code.text = " " & myCode & " "
                
                ' If the cursor is now behind the field, it must be moved back:
                If Selection.End > .result.End Then
                    Selection.Move wdCharacter, -1
                End If
                Selection.Fields.Update
                ' Now, the cursor will be exactly behind the field. That's fine.
                
                ' If the cursor is now in front of the field, it must be moved forward:
                If Selection.Start < .result.Start Then
                    Selection.Start = .result.End
                End If
            End With
            
        Case 1 ' Insert New via .InsertCrossReference
            Dim mainSwitches() As String
            Dim mainFound As Boolean
            Dim Element As Variant
            Dim rmatch As Variant
            ' Check the main switches:
            mainSwitches = Split("PAGEREF|P|REF|R", "|")
            mainFound = False
            For Each Element In mainSwitches
                rmatch = regEx(param3, "(\b" & Trim(Element) & "\b)")
                If rmatch <> False Then
                    'If InStr(1, param3, element) Then
                    mainFound = True
                    param3 = Trim(Replace(param3, rmatch, ""))
                    If Left(Element, 1) = "R" Then
                        param3 = "REF " & param3
                        param0 = wdContentText
                        If Not (IsNumeric(param1)) Then param0 = wdOnlyCaptionText
                    Else
                        param3 = "PAGEREF " & param3
                        param0 = wdPageNumber
                    End If
                    Exit For
                End If
            Next
            If mainFound = False Then
                MsgBox "Insert1CrossRef: non-valid code encountered: <" & param3 & ">"
                Insert1CrossRef = False
                Exit Function
            End If
            
            ' ===== Check the modifier switches:
            If InStr(1, param3, "\n") Then
                param0 = wdNumberNoContext
                param3 = Replace(param3, "\n", "")
            ElseIf InStr(1, param3, "\w") Then
                param0 = wdNumberFullContext
                param3 = Replace(param3, "\w", "")
            ElseIf InStr(1, param3, "\c") Then
                param0 = wdEntireCaption
                param3 = Replace(param3, "\c", "")
                If Not (IsNumeric(param1)) Then param0 = wdEntireCaption
            ElseIf InStr(1, param3, "\r") Then
                param0 = wdNumberRelativeContext
                param3 = Replace(param3, "\r", "")
                If Not (IsNumeric(param1)) Then param0 = wdOnlyLabelAndNumber
            End If
            
            If InStr(1, param3, "\h") Then
                inclHyperlink = True
                param3 = Replace(param3, "\h", "")
            Else
                inclHyperlink = False
            End If
            If InStr(1, param3, "\p") Then
                inclPosition = True
                param3 = Replace(param3, "\p", "")
            Else
                inclPosition = False
            End If
            If InStr(1, param3, "\#0") Then
                param0 = wdOnlyLabelAndNumber
            End If
            
            ' ===== Insert the cross reference, Not all parameters might already be correct:
            '                                  RefType, RefKind, RefIndx, hyperlink,  position     sepNr , seperator
            Call Selection.InsertCrossReference(param1, param0, param2, inclHyperlink, inclPosition, False, "")
            
            param3 = Replace(param3, "PAGEREF", "")
            param3 = Replace(param3, "REF", "")
            
            ' Make sure, the cursor is still in the field
            Do
                idx = CursorInField(Selection.Range)
                If idx <> 0 Then Exit Do
                Selection.MoveLeft wdCharacter, 1
            Loop While True
            
            
            ' ===== Append any leftover switches:
            ' Unfortunately, the order DOES matter in some cases (\#0 must be the FIRST switch), thus:
            If InStr(1, param3, "\#0") Then
                param3 = Replace(param3, "\#0", "")
                'Debug.Print RegExReplace(ActiveDocument.Fields(idx).Code.text, "(Ref\d+)(\s?)", "$1 \#0 ")
                ActiveDocument.Fields(idx).Code.text = RegExReplace(ActiveDocument.Fields(idx).Code.text, "(Ref\d+)(\s?)", "$1 \#0 ")
                ActiveDocument.Fields(idx).Update
            End If
            ActiveDocument.Fields(idx).Code.text = ActiveDocument.Fields(idx).Code.text & " " & param3 & " "
            'Application.StatusBar = "Cross Reference inserted of type <" & param3 & ">."
            
        Case 2 ' Insert New via .Fields.Add
            Selection.Fields.Add Range:=Selection.Range, Type:=wdFieldEmpty, PreserveFormatting:=False
            Selection.TypeText text:=Trim(param4)
            Selection.Fields.Update
            
            ' Put Cursor behind the New field:
            Selection.Move wdCharacter, 1
            Selection.Fields.Update
            'Application.StatusBar = "Cross Reference inserted <" & param4 & ">."
            
        Case Else
            Stop
    End Select
    
    Insert1CrossRef = True
End Function

Private Function MultifieldDelete(optionArray As Variant, _
        ByRef optionPtr As Integer, _
        myCode As String, _
        ByRef Index, _
        Optional checkbyText As String = "", _
        Optional includeLast As Boolean = False) As Integer
    ' Returns: -1 On Error
    '          Else the index of the found format
    ' The Parameters:
    '   optionArray  (in): Array of the configured options
    '   optioPtrr    (IO): ptr To the current Option within the array
    '   myCode           : Not used
    '   index        (in): index of the field Or name of the bookmark
    '   checkbyText  (in): For the type FigureTE
    '   includeLast  (in): Not used
    
    Dim i As Integer ' Loop over options
    Dim j As Integer ' Loop over parts
    Dim theCode As String ' the field's code
    Dim myRange As Range            ' a moving range used To verify the content of each part
    Dim theOption As String ' the current option
    Dim thepart As String ' the current part
    Dim isCode As Boolean ' whether the part is field code Or text
    Dim matchfound As Boolean ' To keep track of the check results
    Dim idx As Integer ' index of the current field in Word's array
    Dim theEnd As Long ' To remember the end of the current multithing
    Dim lOptionPtr As Integer ' a local copy of the optionPtr, which we increment While searching
    Dim textOK As Boolean
    Dim fulltext As String
    Dim expctd As String
    Dim thetext As String
    
    MultifieldDelete = -1
    
    ' Create a dummy Range:
    Set myRange = ActiveDocument.Range
    
    ' Loop over the options:
    For i = 0 To UBound(optionArray)
        lOptionPtr = (optionPtr + i) Mod (UBound(optionArray) + 1)
        theOption = optionArray(lOptionPtr)
        'Debug.Print theOption
        If Len(theOption) = 0 Then
            MsgBox "MultifieldDelete has encountered an invalid option <" & theOption & ">."
            Exit Function
        End If
        
        ' Loop over the parts of one option:
        j = 0
        Do While True
            matchfound = True
            j = j - 1
            thepart = GetPart(theOption, j, isCode)
            If j = 0 Then
                ' There is no Next part, so Exit
                Exit Do
            End If
            If thepart = Error Then
                ' We have checked all parts
                Exit Do
            End If
            
            If j = -1 Then
                ' Make sure, the Cursor is immediately behind the field:
                ActiveDocument.Fields(Index).Update
                Set myRange = Selection.Range
                
                ' If it is a multipart thingy, include the last text in our Range:
                If isCode = False Then
                    myRange.MoveEnd wdCharacter, Len(thepart)
                    myRange.Start = myRange.End
                End If
                
                ' Remember the end of the multithing:
                theEnd = myRange.End
                
            End If
            
            If isCode = False Then
                myRange.MoveStart wdCharacter, -Len(thepart)
                If myRange.text = thepart Then
                    ' OK, matched
                Else
                    ' Mismatch
                    matchfound = False
                    Exit Do
                End If
                myRange.MoveEnd wdCharacter, -Len(thepart)
            Else
                ' It is a Code
                ' Check, If there is a field:
                Call ReplaceAbbrev(thepart)
                idx = CursorInField(myRange)
                If idx = 0 Then
                    matchfound = False
                    Exit Do
                End If
                
                If checkbyText <> "" Then
                    ' This is For the type FigureTE.
                    ' Here, the switches \r, \c are Not applicable,
                    ' rather the Ref points To different bookmarks.
                    fulltext = ActiveDocument.Bookmarks(regEx(myCode, "(_Ref\d+)")).Range.Paragraphs(1).Range.text
                    fulltext = CleanHidden(ActiveDocument.Bookmarks(regEx(myCode, "(_Ref\d+)")).Range.Paragraphs(1).Range)
                    textOK = False
                    If InStr(1, thepart, "PAGEREF") Then
                        textOK = True ' because there is no real check
                    ElseIf InStr(1, thepart, "\p") Then ' above/below
                        textOK = True ' because there is no real check
                    ElseIf InStr(1, thepart, "\r") Then
                        ' Category And number
                        thepart = Trim(Replace(thepart, "\r", ""))
                        expctd = RegExReplace(fulltext, checkbyText, "$3$4$5")
                    ElseIf InStr(1, thepart, "\c") Then
                        ' Full subtitle
                        thepart = Trim(Replace(thepart, "\c", ""))
                        expctd = fulltext
                    ElseIf InStr(1, thepart, "\#0") Then
                        'thePart = Trim(Replace(thePart, "\#0", ""))
                        expctd = regEx(fulltext, "\D(\d+)")
                        ' Number only
                    Else
                        ' Description only
                        expctd = RegExReplace(fulltext, checkbyText, "$7")
                    End If
                    If textOK = False Then
                        thetext = ActiveDocument.Fields(idx).result.text
                        expctd = Replace(expctd, Chr(13), "")
                        If StrComp(thetext, expctd, vbTextCompare) = 0 Then
                            textOK = True
                        End If
                    End If
                    'Else
                    theCode = ActiveDocument.Fields(idx).Code.text
                    theCode = Trim(RegExReplace(theCode, "(REF|PAGEREF)\s+(\S+)", "$1")) ' Remove the bookmark name
                    'End If
                Else
                    textOK = True ' because there is no check
                    theCode = ActiveDocument.Fields(idx).Code
                    theCode = Trim(RegExReplace(theCode, "(REF|PAGEREF)\s+(\S+)", "$1")) ' Remove the bookmark name
                End If
                
                ' Check, If the Field codes match
                ' (present code from options (thePart) vs what's in the document (theCode):
                If CodesComply(theCode, thepart) = False Or textOK = False Then
                    matchfound = False
                    Exit Do
                End If
                myRange.MoveStart wdCharacter, -Len(ActiveDocument.Fields(idx).result.text)
                myRange.MoveEnd wdCharacter, -Len(ActiveDocument.Fields(idx).result.text)
            End If
            
        Loop ' over the parts
        If matchfound Then
            MultifieldDelete = lOptionPtr
            optionPtr = lOptionPtr
            Exit For ' no need To check the other options
        End If
    Next ' over the options
    
    If matchfound = False Then
        ' Not successful in finding the pattern
        Exit Function
    End If
    
    If Abs(j) > 1 Then
        ' It was a multifield
    Else
        ' It was a single field
    End If
    
    ' Delete the whole thing:
    ' Word may try To be smart by removing a lonely blank before Or after the cut-out part.
    ' As we Do Not want that, it gets a bit complicated:
    Dim theStart As Long
    Dim LenStory As Long
    Dim LenCut As Long
    myRange.End = theEnd
    theStart = myRange.Start
    LenStory = ActiveDocument.StoryRanges(wdMainTextStory).StoryLength
    LenCut = theEnd - theStart
    myRange.Cut
    ' Because Word may try To be smart by removing a lonely blank:
    If Selection.Start < theStart Then
        Selection.InsertBefore (" ")
        Selection.Move wdCharacter, 1
    End If
    If ActiveDocument.StoryRanges(wdMainTextStory).StoryLength < LenStory - LenCut Then
        Selection.InsertAfter (" ")
        Selection.Move wdCharacter, -1
    End If
End Function

Private Function strPrepare(string1 As String, Optional withBlanks As Boolean = True) As String
    ' Prepare a configuration string from registry/textbox For use in <InsertCrossReference>.
    ' Therefore, we have To Do the following:
    '   a) strip away comments
    '   b) remove line breaks
    '   c) break into individual configs
    '   d) Treat the special character <? vs <\?
    '   e) Treat escaped <apo>s
    '   f) Trim To have one blank at beginning And end of each line
    
    '   a) strip away comments
    string1 = strRemoveComments(string1)
    
    '   b) remove line breaks
    string1 = Replace(string1, vbNewLine, "")
    
    '   c) break into individual configs, one per line
    '       We can be sure, that there are no more vbNewLines.
    '       Thus we replace the <|> (If they are Not literals) by <vbNewLine>
    string1 = strReplaceNonLits(string1, "|", vbNewLine)
    
    '   d) Treat the special character <? / <\?
    '       Treat <? (representing protected blank Chr(160)) And <\? (representing literal <?):
    string1 = Replace(string1, "?", Chr(160))           ' replace <? by protected blank
    string1 = Replace(string1, "\" & Chr(160), "?")     ' If the <? was escaped (<\?), restore it back To the pound <?
    
    '   e) Treat escaped <apo>s
    string1 = Replace(string1, "\" & "'", "'")
    
    '   f) Trim To have one blank at beginning And end of each line
    If withBlanks = True Then
        string1 = RegExReplace(string1, "\n *", " ")                ' Replace multiple blanks after linebreak by exactly one
        string1 = RegExReplace(string1, "(\S)([\r\n])", "$1 $2")    ' If a line ends Not on a blank, add one
        string1 = RegExReplace(string1, "(\S) {2,}([\r])", "$1 $2") ' Reduce multiple blanks at end of line To one
        string1 = " " & Trim(string1) & " "
    Else
        string1 = Trim(string1) ' Remove possible blanks at beginning & end
        string1 = RegExReplace(string1, "[\r\n]+ *", "|")           ' Replace linebreak And single Or multiple blanks after it by the divider "|"
        'string1 =
    End If
    
    strPrepare = string1
End Function

Private Function CodesComply(ByVal CodeToBCheck As String, ByVal CodeExpected As String) As Boolean
    Dim Element As Variant
    
    CodeToBCheck = Trim(CodeToBCheck)
    CodeExpected = Trim(CodeExpected)
    Call ReplaceAbbrev(CodeExpected)
    
    ' Code complies, If there are exactly the same elements. Order is arbitrary.
    ' Extract the individual words With a regex:
    Dim extract As Object
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    re.Global = True
    re.Pattern = "\S+"                  ' Word by Word
    Set extract = re.Execute(CodeExpected)
    
    For Each Element In extract
        If InStr(1, CodeToBCheck, Element) > 0 Then
            CodeToBCheck = Trim(Replace(CodeToBCheck, Element, ""))
        Else
            CodesComply = False
            Exit Function
        End If
    Next
    
    ' If there is nothing leftover now in theCode except "REF", we have a match:
    If Len(CodeToBCheck) > 0 Then
        CodesComply = False
        Exit Function
    Else
        CodesComply = True
    End If
End Function

Private Function getXRefIndex(RefType, text, Index As Variant) As Boolean
    
    Dim thisitem As String
    Dim i As Integer
    
    text = Trim(text)
    If Right(text, 1) = Chr(13) Then
        text = Left(text, Len(text) - 1)
    End If
    
    getXRefIndex = False
    If RefType = wdRefTypeBookmark Then
        ' The "index" is the bookmark name:
        Index = text
        getXRefIndex = True
    Else
        ' In all other cases, we need To find the index
        ' by searching through Word's CrossReferenceItems(RefType):
        Index = -1
        For i = 1 To UBound(ActiveDocument.GetCrossReferenceItems(RefType))
            thisitem = Trim(Left(Trim(ActiveDocument.GetCrossReferenceItems(RefType)(i)), Len(text)))
            text = Replace(text, Chr(160), " ")
            If StrComp(thisitem, text, vbTextCompare) = 0 Then
                getXRefIndex = True
                Index = i
                Exit For
            End If
        Next
        
        ' Regarding the issue that crossrefs are only found in the document body,
        ' but Not If they are within Textboxes:
        '
        ' Microsoft says (https://learn.microsoft.com/en-gb/office/vba/api/word.selection.insertcrossreference)
        ' that <Selection.InsertCrossReference()> can be used With <ReferenceItem> where
        ' "this argument specifies the item number Or name in the Reference type box in the Cross-reference dialog box".
        ' We have found out, that the dialog box
        ' - first lists the cross references in the document body (wdMainTextStory ?)
        ' - Then  lists the cross references in the text boxes (wdTextFrameStory).
        ' This is True independent of the order of the elements in the document, so:
        ' first all Xrefs in the document, Then all Xrefs in the TextFrames.
        '
        ' Idea therefore: If the XRef was Not found in the document body, Then search in the TextFrames.
        ' This is what the below code does. It finds the XRef And returns the index,
        ' i.e. the ordinal of the XRef within the TextFrame XRefs.
        ' Then If we add this ordinal To the number of XRefs in the document body,
        ' we exactly Get the "item number [..] in the Cross-reference dialog box", i.e. the Index. Sounds good so far.
        ' However, when pass this Index To <Selection.InsertCrossReference()> To actually insert the XRef,
        ' Then this Function crashes, apparently believing that the given index is out-of-bounds.
        ' So there is currently no solution. :-(
        '        If Index = -1 Then
        '            ' Not yet found, Then it is probably in another story:
        '            Index = UBound(ActiveDocument.GetCrossReferenceItems(RefType))
        '            Dim ctr As Integer
        '            Dim pRange As Object
        '            i = 0
        '            For Each pRange In ActiveDocument.StoryRanges    'Loop through all of the stories (https://www.msofficeforums.com/word-vba/38383-Loop-through-all-shapes-all-stories-Not.html#post125397)
        '                Debug.Print pRange.StoryType, pRange.storyLength, Left(pRange.text, 80)
        '                If pRange.StoryType = wdTextFrameStory Then
        '                    i = i + 1
        '                    thisitem = Trim(Left(Trim(pRange.text), Len(text)))
        '                    text = Replace(text, Chr(160), " ")
        '                    If StrComp(thisitem, text, vbTextCompare) = 0 Then
        '                        getXRefIndex = True
        '                        Index = Index + i
        '                        Exit For
        '                    End If
        '                End If
        '            Next
        '        End If
        
    End If
End Function

Private Function isSubtitle(bookmark As String, regexneedle As String) As Boolean
    Dim thetext As String
    
    isSubtitle = False
    
    If ActiveDocument.Bookmarks.Exists(bookmark) = False Then
        Exit Function
    End If
    
    thetext = ActiveDocument.Bookmarks(bookmark).Range.Paragraphs(1).Range.text
    thetext = Replace(thetext, Chr(160), " ")
    If regEx(thetext, regexneedle) <> False Then
        isSubtitle = True
    End If
End Function

Private Function ReplaceAbbrev(thestring) As Boolean
    Dim rmatch As Variant
    Dim needle As String
    Dim repl As String
    
    ReplaceAbbrev = False
    
    needle = "(PAGEREF|P|REF|R)\b" '\b is For word boundary
    'needle = "(OKKLJLK|O|ZUI|Z)\b"
    rmatch = regEx(thestring, needle)
    If rmatch = False Then
        MsgBox "Expected keyword Not found in <" & thestring & ">."
        Exit Function
    End If
    
    If Left(rmatch, 1) = "P" Then
        repl = "PAGEREF"
    Else
        repl = "REF"
    End If
    thestring = RegExReplace(thestring, needle, repl)
    thestring = Trim(thestring)
    
    ReplaceAbbrev = True
End Function

Private Sub ChangeFields()
    Dim objDoc As Document
    Dim objFld As Field
    Dim sFldStr As String
    Dim i As Long, lFldStart As Long
    
    Set objDoc = ActiveDocument
    ' Loop through fields in the ActiveDocument
    For Each objFld In objDoc.Fields
        ' If the field is a cross-ref, Do something To it.
        If objFld.Type = wdFieldRef Then
            Debug.Print objFld.result.text
            GoTo skipsome
            'Make sure the code of the field is visible. You could also just toggle this manually before running the macro.
            objFld.ShowCodes = True
            'I hate using Selection here, but it's probably the most straightforward way To Do this. Select the field, find its start, And Then move the cursor over so that it sits right before the 'R' in REF.
            objFld.Select
            Selection.Collapse wdCollapseStart
            Selection.MoveStartUntil "R"
            'Type 'PAGE' To turn 'REF' into 'PAGEREF'. This turns a text reference into a page number reference.
            Selection.TypeText "PAGE"
            'Update the field so the change is reflected in the document.
            objFld.Update
            objFld.ShowCodes = True
skipsome:
        End If
    Next objFld
End Sub

' ============================================================================================
'
' === Helper routines
'
' ============================================================================================
' ============================================================================================
' === Navigation
' ============================================================================================
Private Function CursorInField(theRange As Range) As Long
    ' If the cursor is currently positioned in a Word field of type wdFieldRef,
    ' Then this Function returns the index of this field.
    ' Else it returns 0.
    
    Dim item As Variant
    
    CursorInField = 0
    'Debug.Print Selection.Start
    
    ' There is Selection.Fields Or Range.Fields, which looks promising To find
    ' the field over which the cursor stands.
    ' But the fields are only listed If the range Or selection overlaps them fully,
    ' Not on partly overlap.
    ' Therefore, we just iterate over all the fields And check their start- And end-position
    ' against the position of the cursor.
    For Each item In ActiveDocument.Fields
        'If Item.index < 50 Then Debug.Print Item.index, Item.Type, Item.Result.Start, Item.Result.End, Item.Result.Case
        If item.Type = wdFieldRef Or item.Type = wdFieldPageRef Then ' wdFieldRef:=3
            If item.result.Start <= theRange.Start And _
                item.result.End >= theRange.Start - 1 Then ' -1 allows that the cursor may stand immediately behind the field
                CursorInField = item.Index
                'Debug.Print "CursorInField: yes"
                Exit Function
            End If
        End If
    Next
End Function

' ============================================================================================
' === Use of arrays
' ============================================================================================
Private Function IsInArray(ByVal stringToBeFound As String, arr As Variant, Optional CaseInsensitive As Boolean = False) As Boolean
    Dim i
    Dim dummy As Integer
    
    ' First check, If the array is possibly empty:
    On Error Resume Next
    dummy = UBound(arr) ' this throws an error on empty arrays, source: https://stackoverflow.com/questions/26290781/check-If-array-is-empty-vba-excel/26290860
    If Err.Number <> 0 Then
        ' The Array is empty!
        IsInArray = False
        Exit Function
    End If
    On Error GoTo 0
    
    If CaseInsensitive = True Then
        For i = LBound(arr) To UBound(arr)
            If LCase(arr(i)) = LCase(stringToBeFound) Then
                IsInArray = True
                Exit Function
            End If
        Next i
    Else
        For i = LBound(arr) To UBound(arr)
            If arr(i) = stringToBeFound Then
                IsInArray = True
                Exit Function
            End If
        Next i
    End If
    IsInArray = False
End Function

' ============================================================================================
' === Regex
' ============================================================================================
Private Function RegExReplace(Quelle As Variant, Expression As Variant, replacement As Variant, Optional multiline As Boolean = False) As String
    ' Beispiel f黵 einen Aufruf:
    ' (w黵de bei mehrfachen Backslashes hintereinander jeweils den ersten wegnehmen)
    ' result = RegExReplace(input, "\\(\\+)", "$1")
    
    'Dim re     As New RegExp
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    
    re.Global = True
    re.multiline = multiline
    re.Pattern = Expression
    RegExReplace = re.Replace(Quelle, replacement)
End Function

Private Function regEx(Quelle As Variant, Expression As String) As Variant
    Dim extract As Object
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    
    re.Global = True
    re.Pattern = Expression
    Set extract = re.Execute(Quelle)
    On Error Resume Next
    regEx = extract.item(0).submatches.item(0)
    If Error <> "" Then
        regEx = False
    End If
End Function

Private Function rgex(strInput As String, matchPattern As String, _
    Optional ByVal outputPattern As String = "$0", Optional ByVal behaviour As String = "") As Variant
    ' How it works:
    ' It takes 2-3 parameters.
    '    A text To use the regular expression on.
    '    A regular expression.
    '    A format string specifying how the result should look. It can contain $0, $1, $2, And so on.
    '         $0 is the entire match, $1 And up correspond To the respective match groups in the regular expression.
    '         Defaults To $0.
    '    If the expression matches multiple times, by default only the first match ("0") is considered.
    '         This can be modified by the Optional parameter.
    '         It can contain 1, 2, 3, ... For the 1st, 2nd, 3rd, ... match Or "*" To return the complete array of matches.
    '
    ' Some examples
    ' Extracting an email address:
    ' =rgex("Peter Gordon: some@email.com, 47", "\w+@\w+\.\w+")
    ' =rgex("Peter Gordon: some@email.com, 47", "\w+@\w+\.\w+", "$0")
    ' Results in: some@email.com
    ' Extracting several substrings:
    ' =rgex("Peter Gordon: some@email.com, 47", "^(.+): (.+), (\d+)$", "E-Mail: $2, Name: $1")
    ' Results in: E-Mail: some@email.com, Name: Peter Gordon
    ' To take apart a combined string in a single cell into its components in multiple cells:
    ' =rgex("Peter Gordon: some@email.com, 47", "^(.+): (.+), (\d+)$", "$" & 1)
    ' =rgex("Peter Gordon: some@email.com, 47", "^(.+): (.+), (\d+)$", "$" & 2)
    ' Results in: Peter Gordon some@email.com ...
    '
    ' Prerequisites: Verweis auf
    ' Microsoft VBScript Regular Expressions 5.5  |c:\windows\SysWOW64\vbscript.dll\3
    '
    ' Modified from source: https://stackoverflow.com/questions/22542834/how-To-use-regular-expressions-regex-in-microsoft-excel-both-in-cell-And-loops/22542835
    
    'Dim inputRegexObj As New VBScript_RegExp_55.RegExp, outputRegexObj As New VBScript_RegExp_55.RegExp, outReplaceRegexObj As New VBScript_RegExp_55.RegExp
    Dim inputRegexObj As Object
    Dim outputRegexObj As Object
    Dim outReplaceRegexObj As Object
    Dim inputMatches As Object
    Dim replaceMatches As Object
    Dim replaceMatch As Object
    Dim replaceNumber As Integer
    Dim ixi As Integer
    Dim ixf As Integer
    Dim ix As Integer
    Dim sepr As String
    Dim outputres As String
    
    Set inputRegexObj = CreateObject("vbscript.regexp")
    Set outputRegexObj = CreateObject("vbscript.regexp")
    Set outReplaceRegexObj = CreateObject("vbscript.regexp")
    
    With inputRegexObj
        .Global = True
        .multiline = True
        .IgnoreCase = False
        .Pattern = matchPattern
    End With
    With outputRegexObj
        .Global = True
        .multiline = True
        .IgnoreCase = False
        .Pattern = "\$(\d+)"
    End With
    With outReplaceRegexObj
        .Global = True
        .multiline = True
        .IgnoreCase = False
    End With
    
    Select Case behaviour
        Case "":  ' No parameter given: by default use match 0
            ixi = 0 ' index initial
            ixf = ixi ' index final
        Case IsNumeric(behaviour):  ' They want a specific match
            ixi = CInt(behaviour) - 1 ' => 1st match has index 0
            ixf = ixi
        Case "*":  ' They want all matches
            ixi = 0
            ixf = -1  ' preliminary value
        Case Else
            MsgBox "errorhasoccured"
    End Select
    
    Set inputMatches = inputRegexObj.Execute(strInput)
    If inputMatches.Count = 0 Then ' Nothing found
        rgex = False
    ElseIf (ixi + 1 > inputMatches.Count) Then ' There is no x-th match-group
        rgex = False
    Else ' Something was found
        rgex = ""
        sepr = ""
        If ixf = -1 Then  ' Now we can determine, how many matches To return
            ixf = inputMatches.Count - 1
            ' Outputformat will be: "{Nr of results}|{Result 1}|{Result 2}|..|{Result N}"
            sepr = "|"
            rgex = CStr(inputMatches.Count)
        End If
        
        ' Reduce results To the requested match-group:
        Set replaceMatches = outputRegexObj.Execute(outputPattern)
        
        For ix = ixi To ixf
            For Each replaceMatch In replaceMatches
                replaceNumber = replaceMatch.submatches(0)
                outReplaceRegexObj.Pattern = "\$" & replaceNumber
                
                If replaceNumber = 0 Then
                    outputres = outReplaceRegexObj.Replace(outputPattern, inputMatches(ix).Value)
                Else
                    If replaceNumber > inputMatches(ix).submatches.Count Then
                        'rgex = "A To high $ tag found. Largest allowed is $" & inputMatches(0).SubMatches.Count & "."
                        rgex = Error 'CVErr(vbErrValue)
                        Exit Function
                    Else
                        outputres = outReplaceRegexObj.Replace(outputPattern, inputMatches(ix).submatches(replaceNumber - 1))
                    End If
                End If
            Next
            rgex = rgex & sepr & outputres
        Next
    End If
End Function

Private Function GetPart(thestring, thePosition, Optional ByRef isCode As Boolean) As Variant
    ' thePosition counts  1, 2, ...
    '                 Or -1, -2, ... To find from behind
    ' It is an I/O-parameter And will be reset To 0 in Case of error.
    
    Dim extract As Object
    Dim idx As Integer
    Dim re As Object
    Set re = CreateObject("vbscript.regexp")
    
    ' 1) Get the different parts
    thestring = Trim(thestring)
    re.Global = True
    re.Pattern = "[^']+(?='|$)" '"[^'|$]+(?='|^)"
    Set extract = re.Execute(thestring)
    
    ' 2) Check If the index is out of bounds:
    If Abs(thePosition) > extract.Count Then
        thePosition = 0
        GetPart = ""
        Exit Function
    End If
    If thePosition = 0 Then
        MsgBox "GetPart() has received invalid index <" & thePosition & ">."
        Stop
    End If
    If thePosition < 0 Then
        ' Find from behind
        idx = extract.Count + thePosition
    Else
        idx = thePosition - 1
    End If
    
    ' 3) Extract the desired item
    GetPart = extract.item(idx)
    If Len(GetPart) = 0 Then
        MsgBox "GetPart() has encountered invalid part: <" & GetPart & ">."
        Stop
    End If
    
    ' 4) As an additional information, return whether this is a string Or a code
    If (Left(thestring, 1) = "'") Then
        isCode = False
    Else
        isCode = True
    End If
    If ((idx Mod 2) > 0) Then
        isCode = Not (isCode)
    End If
End Function

' ============================================================================================
' === String manipulation
' ============================================================================================
Private Function strReplaceNonLits(string1, tbremoved, tbinserted) As String
    Const apo = "'" ' special character For literals
    
    Dim p0, p1 As Long
    Dim l1 As Long
    Dim s2 As String
    
    p0 = 1
    Do While p0 <> 0
        '        Debug.Print "===== " & p0
        '        Debug.Print string1
        ' Find first
        p1 = InStr(p0, string1, tbremoved, vbTextCompare)
        If p1 = 0 Then Exit Do
        
        ' b) Check If it is Not enclosed in <apos>:
        s2 = Mid(string1, 1, p1) ' extract from start To the <cmt>
        s2 = Replace(s2, "\" & apo, "\@")                               ' transform escaped apos (<\apo>) To <\@> For the Next step
        l1 = Len(s2)
        s2 = Replace(s2, apo, "")                                       ' remove the remaining <apo>s, these are the non-escaped ones
        If (l1 - Len(s2)) Mod 2 = 0 Then
            ' outside of <apo>, Then Do the replacement
            string1 = Mid(string1, 1, p1 - 1) & tbinserted & Mid(string1, p1 + Len(tbremoved))
            p0 = p1 + Len(tbinserted)
        Else
            ' inside  of <apo>, Then Do nothing
            p0 = p1 + 1
        End If
    Loop
    strReplaceNonLits = string1
End Function

Private Function strRemoveComments(string1) As String
    Const eol = vbNewLine ' end of line; (vbNewLine=chr(13)+chr(10) )
    Const apo = "'" ' special character For literals
    Const cmt = """" ' special character For comments
    
    Dim p0, p1, p2, p3 As Long ' positions
    Dim l1 As Long ' lengths
    Dim s2 As String ' string To check For <apo>s
    
    p0 = 1
    Do While p0 <> 0
        ' a) Find first <cmt>.
        p1 = InStr(p0, string1, cmt, vbTextCompare)
        If p1 = 0 Then Exit Do
        
        ' b) Check If it is Not enclosed in <apos>:
        p2 = InStrRev(string1, eol, p1, vbTextCompare) ' find start of line (= previous eol+len(eol))
        If p2 = 0 Then ' no previous start of line,
            p2 = 1 '   Then the line starts at 1
        Else
            p2 = p2 + Len(eol) '   Else the line starts after the eol
        End If
        s2 = Mid(string1, p2, p1 - p2) ' extract from start To the <cmt>
        s2 = Replace(s2, "\" & apo, "\@")                               ' transform escaped apos (<\apo>) To <\@> For the Next step
        l1 = Len(s2)
        s2 = Replace(s2, apo, "")                                       ' remove the remaining <apo>s, these are the non-escaped ones
        If (l1 - Len(s2)) Mod 2 = 0 Then ' If the number is even, we are outside <' '>
            ' outside of <apo>, Then remove <cmt> And rest of line
            p3 = InStr(p1, string1, eol, vbTextCompare) ' find Next end of line
            string1 = Mid(string1, 1, p1 - 1) & Mid(string1, p3) ' remove from <cmt> (incl) To end of line (excl)
            p0 = (p1 - 1) + 1 + Len(eol) ' where To continue the search
        Else
            ' inside  of <apo>, Then leave the <cmt>
            p0 = p1 + 1
            
        End If
    Loop
    
    strRemoveComments = string1
End Function

Private Function UnCAPS(aInput As Variant) As String
    Dim result As String
    
    aInput.Font.AllCaps = False
    result = aInput.text
    
    UnCAPS = result
End Function

' ============================================================================================
'
' === The end
'
' ============================================================================================

Public Sub About_RibbonFun(ByVal control As IRibbonControl)
    MsgBox TEXT_AppName + vbCrLf _
         + vbCrLf _
         + TEXT_Description + vbCrLf _
         + TEXT_NonCommecialPrompt + vbCrLf + vbCrLf _
         + TEXT_VersionPrompt + Version + vbCrLf _
         + TEXT_Author + vbCrLf _
         + TEXT_GiteeUrl + vbCrLf _
         + TEXT_GithubUrl, _
        vbOKOnly + vbInformation, C_TITLE
End Sub

Public Sub GetLatestVersion_Github_RibbonFun(ByVal control As IRibbonControl)
    On Error GoTo errHandle
    
    Shell "explorer.exe " & TEXT_GithubUrl
    
    Exit Sub
    
errHandle:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "发生错误 (MakeStandard): " & Err.Description, vbCritical, C_TITLE
    End If
End Sub

Public Sub GetLatestVersion_Gitee_RibbonFun(ByVal control As IRibbonControl)
    On Error GoTo errHandle
    
    Shell "explorer.exe " & TEXT_GiteeUrl
    
    Exit Sub
    
errHandle:
    If Err.Number = ERR_USRMSG Then
        MsgBox Err.Description, vbExclamation, C_TITLE
    ElseIf Err.Number <> ERR_CANCEL Then
        MsgBox "发生错误 (MakeStandard): " & Err.Description, vbCritical, C_TITLE
    End If
End Sub

Public Sub RemoveSpaces_RibbonFun(ByVal control As IRibbonControl)
    Dim rng As Range
    Dim i As Long, j As Long, k As Long
    Dim prevChar As String, nextChar As String
    Dim isChinesePrev As Boolean, isChineseNext As Boolean, foundSpace As Boolean
    Dim ur As UndoRecord
    
    On Error GoTo ERROR_HANDLER
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "删除中英文字符间的空格"

    Set rng = Selection.Range
    If rng.text = "" Then
        Exit Sub
    End If
    
    ' 从后往前处理避免索引变化
    For i = rng.Characters.Count - 1 To 2 Step -1
        j = i
        Do While rng.Characters(j).text = " "
            j = j - 1
            foundSpace = True
        Loop
        
        If foundSpace Then
            prevChar = rng.Characters(j).text
            nextChar = rng.Characters(i + 1).text
            
            isChinesePrev = IsChineseCharacter(prevChar)
            isChineseNext = IsChineseCharacter(nextChar)
            
            ' 如果两端不同或者都为中文，则删除连续的空格
            If (isChinesePrev <> isChineseNext) Or (isChinesePrev And isChineseNext) Then
                For k = i To j + 1 Step -1
                    rng.Characters(k).text = ""
                Next k
            End If
            i = j + 1
            foundSpace = False
        End If
    Next i

    Application.ScreenRefresh
    ur.EndCustomRecord
    Exit Sub ' 正常退出点，避免进入错误处理程序
    
ERROR_HANDLER:
    MsgBox "发生错误: " & vbCrLf & vbCrLf & Err.Description, vbCritical, C_TITLE
    If Not (ur Is Nothing) Then ur.EndCustomRecord
End Sub

' 辅助函数：判断是否为中文字符
Private Function IsChineseCharacter(char As String) As Boolean
    Dim charCode As Long
    
    charCode = AscW(char)
    If charCode < 0 Then charCode = charCode And 65535
    ' 基本汉字 + 标点 + 全角符号 + 扩展汉字
    If (charCode >= CLng(&H4E00) And charCode <= (CLng(&H9FFF) And 65535)) Or _
       (charCode >= CLng(&H3000) And charCode <= CLng(&H303F)) Or _
       (charCode >= (CLng(&HFF00) And 65535) And charCode <= (CLng(&HFFEF) And 65535)) Or _
       (charCode >= CLng(&H3400) And charCode <= CLng(&H4DBF)) Or _
       (charCode >= CLng(&H20000) And charCode <= CLng(&H2FFFF)) Then
        IsChineseCharacter = True
    Else
        IsChineseCharacter = False
    End If
End Function
