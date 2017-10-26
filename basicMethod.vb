
Sub main()
    
  On Error Resume Next
  Dim oPres As Presentation
  Set oPres = Application.ActivePresentation
  Dim oSlide As Slide
  Dim oShape As Shape
  Dim tr As TextRange
  Dim sText As String
  Dim i As Long, j As Long
  keyStr = "www.1ppt.com"
  oSlideCounts = oPres.Slides.Count
    '获取每一页
  For i = 1 To oPres.Slides.Count
    Set oSlide = oPres.Slides.Item(i)
    '遍历shape
    For j = 1 To oSlide.Shapes.Count
        Set oShape = oSlide.Shapes.Item(j)
        '判断类型，找到文字
        If oShape.TextFrame.HasText = msoTrue Then
            Set tr = oShape.TextFrame.TextRange
            sText = tr.Text
            '转换为小写
            sText = LCase(sText)
            MsgBox sText
            
            oPos = InStr(sText, keyStr)
            If oPos > 0 And (i <> oSlideCounts) Then
                '在非最后一页找到字符串，并删除文本框
                MsgBox "del textRange"
                oShape.Delete
            End If
            
            If oPos > 0 And (i = oSlideCounts) Then
                MsgBox "del slide"
                oSlide.Delete
                GoTo GoDone
            End If
            
        Set tr = Nothing
        End If
        Set oShape = Nothing
      Next
      Set oSlide = Nothing
    Next
    Set oPres = Nothing
   
GoDone:

End Sub
