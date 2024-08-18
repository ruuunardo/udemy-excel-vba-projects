Attribute VB_Name = "TableOfContent"
Option Explicit

Sub Add_TOC()
    Dim cellStart As Range
    Dim cellThis As Range
    
On Error GoTo ErrorHandle
    Set cellStart = Excel.Application.InputBox("Input starting cell", , , , , , , 8)
    Set cellThis = cellStart.Cells(1, 1) 'limit to one cell
    
    Dim cntSheets As Integer
    cntSheets = Worksheets.Count
    
    Dim cellEnd As Range
    Set cellEnd = cellThis.Offset(cntSheets - 1, 1)
    
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Range " & cellStart.Address & ":" & cellEnd.Address _
            & " will be overwritten" & vbNewLine & "Are you sure to proceed?" _
            , vbOKCancel + vbDefaultButton2 + vbQuestion)
    
    
    Dim sh As Worksheet
    Dim wsName As String
    Dim thisWsName As String
    thisWsName = ActiveSheet.Name
    
    For Each sh In ActiveWorkbook.Worksheets
        wsName = sh.Name
        If wsName <> thisWsName Then
            ActiveSheet.Hyperlinks.Add Anchor:=cellThis, Address:="", SubAddress:= _
            wsName & "!A1", TextToDisplay:=wsName, ScreenTip:="Go to " & wsName
            
            cellThis.Offset(0, 1).Value = sh.Range("A1")
            Set cellThis = cellThis.Offset(1, 0)
            End If
    Next sh
    
ErrorHandle:
    Exit Sub
End Sub
