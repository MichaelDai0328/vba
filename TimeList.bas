Attribute VB_Name = "Module1"
Function time_list(ByRef Arr, ByVal start_time, ByVal end_time)

    For i = LBound(Arr) To UBound(Arr)
        For j = LBound(Arr, 2) To UBound(Arr, 2)
            Arr(i, j) = Rnd() * (end_time - start_time) + start_time
        Next j
    Next i
    
End Function

Function location_list(ByRef location, ByVal place1, ByVal place2)

    For i = LBound(location) To UBound(location)
        For j = LBound(location, 2) To UBound(location, 2)
            location(i, j) = Round(Rnd())
        Next j
    Next i
    
    For i = LBound(location) To UBound(location)
        For j = LBound(location, 2) To UBound(location, 2)
            If location(i, j) = 1 Then
                location(i, j) = place1
             Else
                location(i, j) = place2
             End If
        Next j
    Next i

End Function

Function print_list(ByVal list_name, _
    Optional ByVal skip_row As Integer = 1, _
    Optional ByVal skip_col As Integer = 1)
    
    For i = LBound(list_name) To UBound(list_name)
        For j = LBound(list_name, 2) To UBound(list_name, 2)
            ActiveSheet.Cells(i + skip_row, j * 2 + skip_col) = list_name(i, j)
        Next j
    Next i

End Function

Function sort_time(ByVal first_row As Integer, _
    ByVal first_col As Integer, ByVal dimention As Integer, _
    ByVal iscale As Integer)
    
    For i = 0 To iscale - 1
        ActiveSheet.Range(Cells(first_row, first_col + i * 2), _
            Cells(first_row + dimention - 1, first_col + i * 2)).Sort _
            Key1:=ActiveSheet.Cells(first_row, first_col + i * 2)
        
    Next
    
End Function

Sub generate_table()
    Dim day_list(1 To 8, 1 To 31)
    Dim night_list(1 To 3, 1 To 31)
    Dim day_location(1 To 4, 1 To 31)
    Dim night_location(1 To 2, 1 To 31)
    day_start = 2 / 3 '16:00
    day_end = 7 / 9 '18:40
    night_start = 0.805555555555555 '19:20
    night_end = 0.875 '21:00

    Call time_list(day_list, day_start, day_end)
    
    Call location_list(day_location, "西小门", "西快速通道")
    
    Call time_list(night_list, night_start, night_end)
    
    Call location_list(night_location, "西小门", "西快速通道")
    
    For i = 0 To 30
        For j = 0 To 3
            ActiveSheet.Cells(15 + j, 7 + i * 2) = "西大门"
        Next j
    Next i
        
    For i = 0 To 30
        ActiveSheet.Cells(21, 7 + i * 2) = "西大门"
    Next i
    
    Call print_list(day_list, 10, 4)

    Call print_list(day_location, 10, 5)
   
    Call print_list(night_list, 18, 4)
    
    Call print_list(night_location, 18, 5)
    
    Call sort_time(11, 6, 4, 31)
    
    Call sort_time(15, 6, 4, 31)
    
    Call sort_time(19, 6, 2, 31)
    
End Sub


