Attribute VB_Name = "SelectShapes"
Option Explicit
Dim Count As Integer

Public Sub Shapes()
    ' Select a Shape in Power Point design mode
    ' Execute the macro that will count all shapes alike
    ' Counted shapes a rotated
    
    Count = 0
    ClearInmediateWindow
    Dim s1 As Slide: Set s1 = ActiveWindow.Selection.SlideRange(1)

    Dim shp0 As Shape: Set shp0 = ActiveWindow.Selection.ShapeRange.Item(1)
    Dim shp1 As Shape
    
    For Each shp1 In s1.Shapes
        
            If CompareShapes(shp0, shp1) Then
                    shp1.Rotation = 30
                    Count = Count + 1
            End If
 
    Next

    MsgBox (Str(Count) + " Elements Found")
    ShowAll

    'Debug.Print Str(Count)

End Sub

Private Function CompareShapes(s1 As Shape, s2 As Shape) As Boolean

Dim cond As Boolean: cond = True
 
        CompareShapes = False

        cond = cond And s1.Height = s2.Height
        cond = cond And s1.Width = s2.Width
        cond = cond And s1.Type = s2.Type
        cond = cond And s1.AlternativeText = s2.AlternativeText
        cond = cond And s1.HasTextFrame = s2.HasTextFrame
        cond = cond And s1.Nodes.Count = s2.Nodes.Count
        cond = cond And (s1.Fill.ForeColor = s2.Fill.ForeColor)
        cond = cond And (s1.Fill.BackColor = s2.Fill.BackColor)
        
        If Not cond Then Exit Function
        
        If s1.HasTextFrame Then
            cond = cond And s1.TextFrame.TextRange.Text = s2.TextFrame.TextRange.Text
            cond = cond And s1.TextFrame2.TextRange.Text = s2.TextFrame2.TextRange.Text
        End If
        
        If Not cond Then Exit Function
        
        If (s1.Type = msoGroup) Then
                    
            If (s1.GroupItems.Count = s2.GroupItems.Count) Then
            Dim n As Integer
                For n = 1 To s1.GroupItems.Count
                    cond = cond And CompareShapes(s1.GroupItems(n), s2.GroupItems(n))
                Next n
            Else
                cond = False
            End If
        
        ElseIf (s1.Type = msoGraphic) Then
            
            cond = cond And s1.Reflection.Size = s2.Reflection.Size
        
        End If
         
        CompareShapes = cond

End Function

Private Sub ShowAll()
    'Debug procedure
    Dim shp1 As Shape

    For Each shp1 In ActiveWindow.Selection.SlideRange(1).Shapes
        shp1.Visible = True
        shp1.Rotation = 0
    Next
End Sub

Private Sub ClearInmediateWindow()
    'For debugging
   VBA.SendKeys "^g^a{DEL}", True
End Sub


