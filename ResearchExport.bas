Attribute VB_Name = "ResearchExport"
Option Explicit
Private dpiValue As Long
Public Sub SetDpi(control As IRibbonControl)
    dpiValue = CLng(control.Tag)
End Sub
Public Sub CustomDpi(control As IRibbonControl)
    Dim s As String, d As Double
    s = InputBox("输入 DPI（36–2400）", "自定义 DPI", CStr(CurrentDpi()))
    If s = "" Then Exit Sub
    If Not IsNumeric(s) Then Exit Sub
    d = CDbl(s)
    If d < 36 Or d > 2400 Or d <> Fix(d) Then
        MsgBox "请输入 36–2400 的整数。": Exit Sub
    End If
    dpiValue = CLng(d)
End Sub
Private Function CurrentDpi() As Long
    If dpiValue = 0 Then dpiValue = 600
    CurrentDpi = dpiValue
End Function
Public Sub ExportSlide(control As IRibbonControl)
    RunExport False, False
End Sub
Public Sub ExportSelection(control As IRibbonControl)
    RunExport True, False
End Sub
Public Sub ExportTransparent(control As IRibbonControl)
    RunExport True, True
End Sub
Private Function Q(ByVal s As String) As String
    Q = Chr$(34) & s & Chr$(34)
End Function
Private Sub RunExport(ByVal selected As Boolean, ByVal transparent As Boolean)
    Dim src As Presentation, tmp As Presentation, sl As Slide
    Dim sr As ShapeRange, pasted As ShapeRange, obj As Shape
    Dim w As Double, h As Double, canvasW As Double, canvasH As Double
    Dim canvasPx As Long, canvasPy As Long
    Dim px As Long, py As Long, dpi As Long, fd As FileDialog
    Dim dest As String, raw As String, helper As String, cmd As String
    Dim shell As Object, rc As Long, errText As String
    On Error GoTo Failed
    If Application.Windows.Count = 0 Then Err.Raise 5, , "请先打开演示文稿。"
    Set src = ActivePresentation
    Set sl = ActiveWindow.View.Slide
    dpi = CurrentDpi()
    If selected Then
        If ActiveWindow.Selection.Type <> ppSelectionShapes Then Err.Raise 5, , "请用选择框选中对象或组合（不要进入文字编辑状态）。"
        Set sr = ActiveWindow.Selection.ShapeRange
        'Set canvas BEFORE paste. Never resize a pasted group or its children.
        w = sr.Width: h = sr.Height
        canvasW = src.PageSetup.SlideWidth
        canvasH = src.PageSetup.SlideHeight
        If canvasW < w Then canvasW = w
        If canvasH < h Then canvasH = h
        If canvasW < 72 Then canvasW = 72
        If canvasH < 72 Then canvasH = 72
        sr.Copy
        Set tmp = Presentations.Add(msoFalse)
        tmp.PageSetup.SlideWidth = canvasW
        tmp.PageSetup.SlideHeight = canvasH
        Set sl = tmp.Slides.Add(1, ppLayoutBlank)
        Set pasted = sl.Shapes.Paste
        If pasted.Count > 1 Then
            Set obj = pasted.Group
        Else
            Set obj = pasted(1)
        End If
        w = obj.Width: h = obj.Height
        If w <= 0 Or h <= 0 Then Err.Raise 5, , "选区尺寸无效。"
        obj.Left = 0: obj.Top = 0
        sl.FollowMasterBackground = msoFalse
        sl.Background.Fill.Solid
        sl.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
    Else
        w = src.PageSetup.SlideWidth: h = src.PageSetup.SlideHeight
        canvasW = w: canvasH = h
    End If
    px = CLng(Fix(w / 72# * dpi + 0.5))
    py = CLng(Fix(h / 72# * dpi + 0.5))
    canvasPx = CLng(Fix(canvasW / 72# * dpi + 0.5))
    canvasPy = CLng(Fix(canvasH / 72# * dpi + 0.5))
    If canvasPx > 32767 Or canvasPy > 32767 Or CDbl(canvasPx) * canvasPy > 100000000# Then Err.Raise 5, , "渲染画布过大，请降低 DPI 或在较小页面中制作组图。"
    If px < 1 Or py < 1 Or px > 32767 Or py > 32767 Then Err.Raise 5, , "像素边长须在 1–32767 内；请降低 DPI 或尺寸。"
    If CDbl(px) * py > 100000000# Then Err.Raise 5, , "超过一亿像素，请降低 DPI。"
    If MsgBox("输出尺寸：" & Format(w / 72# * 25.4, "0.00") & " × " & Format(h / 72# * 25.4, "0.00") & " mm" & vbCrLf & px & " × " & py & " 像素，" & dpi & " DPI" & vbCrLf & "高 DPI 不会恢复原始低清图片的细节。继续？", vbOKCancel) <> vbOK Then GoTo Clean
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.Title = "保存科研 PNG（自动使用 .png 扩展名）"
    fd.InitialFileName = "科研图片.png"
    If fd.Show <> -1 Then GoTo Clean
    dest = fd.SelectedItems(1)
    If LCase$(Right$(dest, 4)) <> ".png" Then dest = dest & ".png"
    If Len(Dir$(dest)) > 0 Then
        If MsgBox("文件已存在，覆盖？" & vbCrLf & dest, vbYesNo Or vbQuestion) <> vbYes Then GoTo Clean
    End If
    helper = Environ$("APPDATA") & "\ResearchExport\Finalize-Png.ps1"
    If Len(Dir$(helper)) = 0 Then Err.Raise 53, , "缺少 PNG 校验组件，请重新运行安装脚本。"
    Set shell = CreateObject("WScript.Shell")
    raw = Environ$("TEMP") & "\ResearchExport-" & CreateObject("Scripting.FileSystemObject").GetTempName & ".png"
    If transparent Then
        obj.Export raw, ppShapeFormatPNG, px, py, ppRelativeToSlide
    Else
        sl.Export raw, "PNG", canvasPx, canvasPy
    End If
    cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & Q(helper) & " -InputPath " & Q(raw) & " -OutputPath " & Q(dest) & " -Dpi " & dpi & " -Width " & px & " -Height " & py
    If selected And Not transparent Then cmd = cmd & " -CanvasWidth " & canvasPx & " -CanvasHeight " & canvasPy
    If transparent Then cmd = cmd & " -RequireAlpha"
    rc = shell.Run(cmd, 0, True)
    If rc <> 0 Then Err.Raise 5, , "PNG 尺寸、透明通道或保存校验失败，未替换目标文件。请尝试普通选区导出或降低 DPI。日志：" & raw & ".error.txt"
    MsgBox "已导出：" & dest & vbCrLf & px & " × " & py & " 像素；已写入 DPI 元数据。", vbInformation
Clean:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Saved = msoTrue: tmp.Close
    If Len(raw) > 0 Then
        If Len(Dir$(raw)) > 0 Then Kill raw
    End If
    Exit Sub
Failed:
    errText = Err.Description
    MsgBox "导出未完成：" & errText, vbExclamation
    Resume Clean
End Sub

'v1.2: optional movable crop frame; original export commands are unchanged.
Public Sub AddCropFrame(control As IRibbonControl)
    Dim sw As String, sh As String
    On Error GoTo Failed
    If Application.Windows.Count = 0 Then Err.Raise 5, , "请先打开演示文稿。"
    sw = InputBox("裁剪输出宽度（毫米）", "插入裁剪框", "55")
    If sw = "" Then Exit Sub
    sh = InputBox("裁剪输出高度（毫米）", "插入裁剪框", "50")
    If sh = "" Then Exit Sub
    If Not IsNumeric(sw) Or Not IsNumeric(sh) Then Err.Raise 5, , "宽高必须是毫米数值。"
    InsertCropFrame CDbl(sw), CDbl(sh)
    Exit Sub
Failed:
    MsgBox "未插入裁剪框：" & Err.Description, vbExclamation
End Sub
Private Sub InsertCropFrame(ByVal widthMm As Double, ByVal heightMm As Double)
    Dim slideW As Double, slideH As Double, fw As Double, fh As Double
    Dim frame As Shape, currentSlide As Slide
    slideW = ActivePresentation.PageSetup.SlideWidth
    slideH = ActivePresentation.PageSetup.SlideHeight
    fw = widthMm / 25.4 * 72#: fh = heightMm / 25.4 * 72#
    If fw <= 0 Or fh <= 0 Or fw > slideW + 0.01 Or fh > slideH + 0.01 Then Err.Raise 5, , "裁剪框宽高须大于零且不超过当前幻灯片。"
    If fw > slideW Then fw = slideW
    If fh > slideH Then fh = slideH
    Set currentSlide = ActiveWindow.View.Slide
    Set frame = currentSlide.Shapes.AddShape(msoShapeRectangle, (slideW - fw) / 2, (slideH - fh) / 2, fw, fh)
    frame.Tags.Add "ResearchExportCrop", "1"
    frame.Name = "科研裁剪框_" & CStr(frame.Id)
    frame.Fill.Visible = msoFalse
    frame.Line.Visible = msoTrue
    frame.Line.ForeColor.RGB = RGB(220, 30, 70)
    frame.Line.Weight = 1
    frame.Line.DashStyle = msoLineDash
    frame.LockAspectRatio = msoFalse
    frame.Select
End Sub
Public Sub ExportCropFrame(control As IRibbonControl)
    RunCropExport
End Sub
Private Sub RunCropExport()
    Dim src As Presentation, tmp As Presentation, originalSlide As Slide, croppedSlide As Slide
    Dim frame As Shape, fd As FileDialog, pastedSlides As SlideRange
    Dim w As Double, h As Double, x As Double, y As Double, cw As Double, ch As Double
    Dim px As Long, py As Long, canvasPx As Long, canvasPy As Long, cropX As Long, cropY As Long
    Dim dpi As Long, i As Long, dest As String, raw As String, helper As String, cmd As String
    Dim shell As Object, rc As Long, errorText As String
    On Error GoTo Failed
    If Application.Windows.Count = 0 Then Err.Raise 5, , "请先打开演示文稿。"
    If ActiveWindow.Selection.Type <> ppSelectionShapes Then Err.Raise 5, , "请先选中新增的裁剪框（不要将框与组图组合）。"
    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then Err.Raise 5, , "请只选中一个裁剪框。"
    Set frame = ActiveWindow.Selection.ShapeRange(1)
    If frame.Tags("ResearchExportCrop") <> "1" Then Err.Raise 5, , "请先点“插入裁剪框”，移动框后单独选中它，再导出。"
    If Abs(frame.Rotation) > 0.01 Then Err.Raise 5, , "裁剪框不能旋转，请将旋转角度设为 0。"
    Set src = ActivePresentation
    Set originalSlide = ActiveWindow.View.Slide
    cw = src.PageSetup.SlideWidth: ch = src.PageSetup.SlideHeight
    x = frame.Left: y = frame.Top: w = frame.Width: h = frame.Height
    If w <= 0 Or h <= 0 Or x < -0.01 Or y < -0.01 Or x + w > cw + 0.01 Or y + h > ch + 0.01 Then Err.Raise 5, , "请将整个裁剪框放在幻灯片内。可用“形状格式 → 大小”精确调整宽高。"
    dpi = CurrentDpi()
    px = CLng(Fix(w / 72# * dpi + 0.5)): py = CLng(Fix(h / 72# * dpi + 0.5))
    canvasPx = CLng(Fix(cw / 72# * dpi + 0.5)): canvasPy = CLng(Fix(ch / 72# * dpi + 0.5))
    If px < 1 Or py < 1 Or canvasPx > 32767 Or canvasPy > 32767 Or CDbl(canvasPx) * canvasPy > 100000000# Then Err.Raise 5, , "渲染尺寸超限，请降低 DPI。"
    cropX = CLng(Fix(x / cw * canvasPx + 0.5)): cropY = CLng(Fix(y / ch * canvasPy + 0.5))
    If cropX < 0 Then cropX = 0
    If cropY < 0 Then cropY = 0
    If cropX + px > canvasPx Then cropX = canvasPx - px
    If cropY + py > canvasPy Then cropY = canvasPy - py
    If MsgBox("裁剪输出：" & Format(w / 72# * 25.4, "0.00") & " × " & Format(h / 72# * 25.4, "0.00") & " mm" & vbCrLf & px & " × " & py & " 像素，" & dpi & " DPI" & vbCrLf & "输出框内全部可见内容和幻灯片背景；所有科研裁剪框会从输出中去除。", vbOKCancel) <> vbOK Then Exit Sub
    Set fd = Application.FileDialog(msoFileDialogSaveAs)
    fd.Title = "保存裁剪区域 PNG"
    fd.InitialFileName = "科研裁剪.png"
    If fd.Show <> -1 Then Exit Sub
    dest = fd.SelectedItems(1)
    If LCase$(Right$(dest, 4)) <> ".png" Then dest = dest & ".png"
    If Len(Dir$(dest)) > 0 Then
        If MsgBox("文件已存在，覆盖？" & vbCrLf & dest, vbYesNo Or vbQuestion) <> vbYes Then Exit Sub
    End If
    helper = Environ$("APPDATA") & "\ResearchExport\Finalize-Png.ps1"
    If Len(Dir$(helper)) = 0 Then Err.Raise 53, , "缺少 PNG 处理组件，请重新安装。"
    originalSlide.Copy
    Set tmp = Presentations.Add(msoFalse)
    tmp.PageSetup.SlideWidth = cw: tmp.PageSetup.SlideHeight = ch
    Set pastedSlides = tmp.Slides.Paste(1)
    Set croppedSlide = pastedSlides(1)
    For i = croppedSlide.Shapes.Count To 1 Step -1
        If croppedSlide.Shapes(i).Tags("ResearchExportCrop") = "1" Then croppedSlide.Shapes(i).Delete
    Next i
    Set shell = CreateObject("WScript.Shell")
    raw = Environ$("TEMP") & "\ResearchExport-" & CreateObject("Scripting.FileSystemObject").GetTempName & ".png"
    croppedSlide.Export raw, "PNG", canvasPx, canvasPy
    cmd = "powershell.exe -NoProfile -NonInteractive -ExecutionPolicy Bypass -File " & Q(helper) & " -InputPath " & Q(raw) & " -OutputPath " & Q(dest) & " -Dpi " & dpi & " -Width " & px & " -Height " & py & " -CanvasWidth " & canvasPx & " -CanvasHeight " & canvasPy & " -CropX " & cropX & " -CropY " & cropY
    rc = shell.Run(cmd, 0, True)
    If rc <> 0 Then Err.Raise 5, , "裁剪或保存校验失败，未替换目标文件。日志：" & raw & ".error.txt"
    MsgBox "裁剪已导出：" & dest & vbCrLf & px & " × " & py & " 像素；已写入 DPI 元数据。", vbInformation
Clean:
    On Error Resume Next
    If Not tmp Is Nothing Then tmp.Saved = msoTrue: tmp.Close
    If Len(raw) > 0 Then
        If Len(Dir$(raw)) > 0 Then Kill raw
    End If
    Exit Sub
Failed:
    errorText = Err.Description
    MsgBox "裁剪未完成：" & errorText, vbExclamation
    Resume Clean
End Sub
