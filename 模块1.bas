Attribute VB_Name = "Ä£¿é1"
Sub test()
Dim myArr As Variant
Dim i As String
Dim j As Integer

'Dim Item As String

myArr = Array("192.168.255.1", "192.168.255.3", "192.168.255.13", "192.168.255.14", "192.168.255.5", "192.168.255.7", "192.168.255.9", "192.168.255.10")


j = 1
'For j = 1 To 4
For Each Item In myArr
        
        'MsgBox Item
    
    Cells(j, 1).Value = Item
    j = j + 1
    'Next
    
Next

End Sub
