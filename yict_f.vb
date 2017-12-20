Sub YIC_Tinset_image()
'将指定目录下的图片导入到excel指定的单元格中。

Dim myArr As Variant 'vb不能定义常量数组，要先声明
Dim i As String
Dim j As Integer

myArr = Array("192.168.255.1", "192.168.255.3", "192.168.255.13", "192.168.255.14", "192.168.255.5", "192.168.255.7", "192.168.255.9", "192.168.255.10", "192.168.255.15", "192.168.255.16", "192.168.9.4", "192.168.230.2", "172.16.0.1-0.3", "172.16.1.4", "172.16.1.5", "172.16.1.6", "172.16.1.7", "172.16.1.8", "172.16.1.9",)

j = 3 
For Each Item In myArr
        
        'MsgBox Item
    
    'Cells(j, 1).Value = Item
	Range("B" & j).Select
	ActiveSheet.Pictures.Insert("D:\test\img\" & Item & "- CPU - last7day.png"). _
        Select
	Range("D" & j).Select
	ActiveSheet.Pictures.Insert( _
        "D:\test\img\" & Item & " - memory - last7day.png").Select
		
    j = j + 2
     
Next


		
