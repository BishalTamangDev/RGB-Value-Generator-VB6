Attribute VB_Name = "Module1"
Type colour
    red As Integer
    green As Integer
    blue As Integer
End Type

Public Function return_file_size() As Integer
Dim file_size As Integer
Dim file_data As colour

file_size = 1
Open "RGBvalues.txt" For Random As #1 Len = 6
    Do While EOF(1) = False
        Get #1, file_size, file_data
        file_size = file_size + 1
    Loop
    file_size = file_size - 2
Close 1
return_file_size = file_size
End Function

