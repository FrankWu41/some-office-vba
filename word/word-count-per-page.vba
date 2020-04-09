Sub WordCountPerPage()
    Dim NumSec As Integer
    Dim BreakCount As Integer
    Dim S As Integer
    Dim Summary As String
    Dim total As Integer
    NumSec = ActiveDocument.ActiveWindow.Panes(1).Pages.Count
    total = 0
    For S = 1 To NumSec
        Selection.GoTo What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=S
        total = total + Selection.Bookmarks("\Page").Range.ComputeStatistics(wdStatisticWords)
        Summary = Summary & "PageNo: " & S & " Current:" & Selection.Bookmarks("\Page").Range.ComputeStatistics(wdStatisticWords) & " Total：" & total & vbCrLf
    Next
    MsgBox Summary
End Sub