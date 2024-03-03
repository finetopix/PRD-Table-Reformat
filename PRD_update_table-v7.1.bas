Attribute VB_Name = "Module11"
Sub Table_Format_Update()
Dim totalrows As Long
Dim rng As range
Dim matchedcount As Integer
Dim row As Integer
Dim nextStr As String
Dim RefStr As String
Dim tableindex As Integer
Dim tbl As Word.Table
Dim iRow As Integer
' v5: remove line breaks, added operator names for -J NR3500 TUs; added space after slash for Joint owner names. 10/11/2023
' v6  2024.01.11 revert to full Str compare
' v7: power value scripts updated,some antenna model name move to second line 05/02/2024
' v7.1 bugfix:two rows are same antenna
If Selection.Information(wdWithInTable) = True Then
    tableindex = ThisTableNumber
    totalrows = FindNumberofRows() ' get the total rows in selected table
Else
    MsgBox "Selection is not in a table."
        Exit Sub
End If

MsgBox "Table updates will take a while, Please Click OK to start..."

Set tbl = ActiveDocument.Tables(tableindex)
'replace text
Call Replace_text(tableindex)
'change font size to 2 in the table
Call Change_font_size(tableindex, 2)

RefStr = tbl.cell(1, 1).range.Text  ' reference string
nextStr = tbl.cell(2, 1).range.Text
matchedcount = 0
        row = 1
        '******place cursor at start
        Set rng = tbl.cell(row, 1).range
        rng.Collapse Direction:=wdCollapseStart
        rng.Select
 
Start:
        
        iRow = row
        Do While nextStr = RefStr '2024.01.11 revert to full Str compare
            matchedcount = matchedcount + 1
            If iRow + 1 >= totalrows Then
            Exit Do
            Else
            nextStr = tbl.cell(iRow + 1, 1).range.Text
            End If
            
            iRow = iRow + 1
        Loop
        If (mathchedcount + row < totalrows) Then
        RefStr = tbl.cell(matchedcount + row, 1).range.Text
        nextStr = tbl.cell(matchedcount + row, 1).range.Text
        End If
        
        If matchedcount >= 2 Then
           Call Merge_rows(tableindex, matchedcount)
        End If
        
        row = matchedcount + row
       
        matchedcount = 0
        
        If row >= totalrows Then
        'change back font size to 11
        Call replace_text_after_merge(tableindex)
        Call Change_font_size(tableindex, 11)
        MsgBox "Table updates done"
        
        Exit Sub
        Else
         Set rng = tbl.cell(row, 1).range
            rng.Collapse Direction:=wdCollapseStart
            rng.Select
        GoTo Start
        End If
End Sub

Function FindNumberofRows() As Long
    Dim rows As Long
    'MsgBox (Selection.Information(wdMaximumNumberOfRows))
    FindNumberofRows = Selection.Information(wdMaximumNumberOfRows)
    'MsgBox (rows)
End Function

Function ThisTableNumber() As Integer
    Dim CurrentSelection As Long
    Dim T_Start As Long
    Dim T_End As Long
    Dim oTable As Table
    Dim j As Long
    CurrentSelection = Selection.range.Start
    For Each oTable In ActiveDocument.Tables
        T_Start = oTable.range.Start
        T_End = oTable.range.End
        j = j + 1
        'ThisTableNumber = "Couldn't determine table number" ' Added error message
        If CurrentSelection >= T_Start And _
            CurrentSelection <= T_End Then ' added "="
            ThisTableNumber = j
            Exit For
        End If
    Next
End Function
Sub Replace_text(tableindex)
    Dim tbl As Table
    Dim cell As cell
    Dim row As Integer
    Dim vendor As String
    Dim height As Double
    Dim port_item As Variant
    Dim result  As String
    
    Set tbl = ActiveDocument.Tables(tableindex)
    For Each cell In tbl.range.Cells
        Select Case cell.ColumnIndex
            Case 1 'column 1 process
              If InStr(cell.range.Text, "-J") Then
                row = cell.rowIndex
                vendor = tbl.cell(row, 2).range.Text
                'Modify system/sector column to add vendor info
                If InStr(vendor, "Vodafone") Then
                    If InStr(tbl.cell(row, 9), "NR") Then
                    'add operator name
                    tbl.cell(row, 9).range.Text = Replace(tbl.cell(row, 9).range.Text, "NR", "TPG NR")
                    Else
                       tbl.cell(row, 9).range.Text = Replace(tbl.cell(row, 9).range.Text, "LTE", "TPG NR/LTE")
                       tbl.cell(row, 9).range.Text = Replace(tbl.cell(row, 9).range.Text, "3.5GHz", "TPG NR 3500")
                    End If
                    'tbl.cell(row, 9).Range.Text = Replace(tbl.cell(row, 9).Range.Text, "LTE", vendor & " NR/LTE")
                ElseIf InStr(vendor, "TPG") Then
                        If InStr(tbl.cell(row, 9), "NR") Then
                        'add operator name
                        tbl.cell(row, 9).range.Text = Replace(tbl.cell(row, 9).range.Text, "NR", "TPG NR")
                        Else
                            tbl.cell(row, 9).range.Text = Replace(tbl.cell(row, 9).range.Text, "LTE", "TPG NR/LTE")
                        End If
                Else
                        'if vendor is Optus, just add vendor in front of the TU
                        tbl.cell(row, 9).range.Text = Replace(tbl.cell(row, 9).range.Text, tbl.cell(row, 9).range.Text, vendor & " " & tbl.cell(row, 9).range.Text)
                End If
                'Change column 2 to joint venture name
                tbl.cell(row, 2).range.Text = "Optus/ Vodafone Joint Venture"
              ElseIf InStr(cell.range.Text, "-V") Then
                row = cell.rowIndex
                vendor = tbl.cell(row, 2).range.Text
                If InStr(tbl.cell(row, 9), "NR") Then
                        'nothing to do
                ElseIf InStr(tbl.cell(row, 9), "LTE") Then
                    tbl.cell(row, 9).range.Text = Replace(tbl.cell(row, 9).range.Text, "LTE", "NR/LTE")
                End If
              End If
            
            Case 9
                row = cell.rowIndex
                If InStr(cell.range.Text, "3.64GHz") Then
                    'cell.Range.Text = "NR 3500"
                    cell.range.Text = Replace(cell.range.Text, "3.64GHz", "NR 3500")
                ElseIf InStr(cell.range.Text, "3.5GHz") Then
                    'cell.Range.Text = "NR 3500"
                    cell.range.Text = Replace(cell.range.Text, "3.5GHz", "NR 3500")
                ElseIf InStr(cell.range.Text, "3.56GHz") Then
                    'cell.Range.Text = "NR 3500"
                    cell.range.Text = Replace(cell.range.Text, "3.56GHz", "NR 3500")
                ElseIf InStr(cell.range.Text, "Wimax 2300") Then
                    'cell.Range.Text = "NR 2300"
                    cell.range.Text = Replace(cell.range.Text, "3.64GHz", "NR 3500")
                End If
                'remove line breaks
                cell.range.Text = Replace(cell.range.Text, Chr(13), "")
            
            Case 10
                port_item = ""
                result = ""
                'if cell value with multiple ports,then divided by "+" to remove zeros, then merge again
                If InStr(cell.range.Text, "+") Then
                    port_item = Split(Replace(cell.range.Text, "", ""), "+")
                    For Each item In port_item
                        If item <> "0" Then
                            'Debug.Print cell.Range.Text
                            If result = "" Then
                                result = item 'first port power value assign to cell directly
                            Else
                                result = result + "+" + item '2nd and after connected with "+"
                            End If
                        End If
                    Next item
                End If
                
                'remove line breaks and assign result to the cell
                
                cell.range.Text = Replace(result, Chr(13), "")
                
        End Select
    Next cell
End Sub
Sub Merge_rows(tableindex, matchedcount)
    Dim row_count As Integer
    Dim n, m, x, y As Integer
    Dim selectedRange As range
    Dim tbl As Word.Table
    Dim rowIndex, colIndex As Integer
    Dim newCell As cell
    Set tbl = ActiveDocument.Tables(tableindex)
    
    n = matchedcount
    m = matchedcount - 1
    
    Selection.MoveDown
    
    If n = 2 Then
       'Selection.MoveRight Unit:=wdCharacter, Count:=6, Extend:=wdExtend
       Set selectedRange = Selection.range
       rowIndex = Selection.Cells(1).rowIndex
       colIndex = 6 ' from column diagram ref to Column Mech Tilt
       For i = 2 To 6
        Set newCell = tbl.cell(rowIndex, i)
        selectedRange.Expand Unit:=wdCell
        selectedRange.SetRange Start:=selectedRange.Start, End:=newCell.range.End
       Next i
       selectedRange.Select
    Else
       x = n - 2
       Selection.MoveDown Unit:=wdLine, Count:=x, Extend:=wdExtend
       Selection.MoveRight Unit:=wdCharacter, Count:=5, Extend:=wdExtend
    End If
    
    
    Selection.Delete
    Selection.MoveUp
    Selection.MoveDown Unit:=wdLine, Count:=m, Extend:=wdExtend
    Selection.Cells.Merge
    
    For y = 1 To 5
        Selection.MoveRight
        'Selection.Delete
        Selection.MoveDown Count:=m, Extend:=wdExtend
        Selection.Cells.Merge
    Next y
            
End Sub
Sub replace_text_after_merge(tableindex)
    Dim tbl As Table
    Dim cell As cell
    Dim line1, line2 As String
    
    Set tbl = ActiveDocument.Tables(tableindex)
    For Each cell In tbl.Columns(3).Cells
        If InStr(cell.range.Text, "RR2VV-6533D-R6-[O]") Then
            line1 = Replace(cell.range.Text, "RR2VV-6533D-R6-[O]", "")
            line2 = "RR2VV-6533D-R6-[O]"
            cell.range.Text = line1 & vbCrLf & line2
            GoTo p1
        End If
        If InStr(cell.range.Text, "RRZZVVT4S4-65D-R8 [Service Beam]-[O]") Then
            line1 = Replace(cell.range.Text, "RRZZVVT4S4-65D-R8 [Service Beam]-[O]", "")
            line2 = "RRZZVVT4S4-65D-R8 [Service Beam]-[O]"
            cell.range.Text = line1 & vbCrLf & line2
            GoTo p1
        End If
        If InStr(cell.range.Text, "RRV4-65B-R6-[O]") Then
            line1 = Replace(cell.range.Text, "RRV4-65B-R6-[O]", "")
            line2 = "RRV4-65B-R6-[O]"
            cell.range.Text = line1 & vbCrLf & line2
            GoTo p1
        End If
        If InStr(cell.range.Text, "RRV4-65D-R6-[O]") Then
            line1 = Replace(cell.range.Text, "RRV4-65D-R6-[O]", "")
            line2 = "RRV4-65D-R6-[O]"
            cell.range.Text = line1 & vbCrLf & line2
            GoTo p1
        End If
        If InStr(cell.range.Text, "AEQP_I CS7801002 IPAA 2.7m") Then
            line1 = Replace(cell.range.Text, "AEQP_I CS7801002 IPAA 2.7m", "")
            line2 = "AEQP_I CS7801002 IPAA 2.7m"
            cell.range.Text = line1 & vbCrLf & line2
            GoTo p1
        End If
        If InStr(cell.range.Text, "RVVPX310.11B-T2") Then
            line1 = Replace(cell.range.Text, "RVVPX310.11B-T2", "")
            line2 = "RVVPX310.11B-T2"
            cell.range.Text = line1 & vbCrLf & line2
            GoTo p1
        End If
        If InStr(cell.range.Text, "Ericsson/UKY-21077SC11") Then
            line1 = Replace(cell.range.Text, "Ericsson/UKY-21077SC11", "")
            line2 = "Ericsson/UKY-21077SC11"
            cell.range.Text = line1 & vbCrLf & line2
            GoTo p1
        End If
p1:
     Next cell
End Sub
Sub Change_font_size(tableindex, size)
Dim fontSize As Integer
Dim tbl As Table
Dim cell As cell

fontSize = size
Set tbl = ActiveDocument.Tables(tableindex)
    For Each cell In tbl.range.Cells
        cell.range.Font.size = fontSize
    Next cell

End Sub



