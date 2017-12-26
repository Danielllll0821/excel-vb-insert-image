Sub YIC_Tinset_image()
'将指定目录下的图片导入到excel指定的单元格中。

Dim myArr As Variant 'vb不能定义常量数组，要先声明
Dim i As String
Dim j As Integer

myArr = Array("192.168.255.1", "192.168.255.3", "192.168.255.13", "192.168.255.14", "192.168.255.5", "192.168.255.7", "192.168.255.9", "192.168.255.10", "192.168.255.15", "192.168.255.16", "192.168.9.4", "192.168.230.2", "172.16.0.1-0.3", "172.17.1.4", "172.17.1.5", "172.17.1.6", "172.17.1.7", "172.17.1.8", "172.17.1.9")

j = 3 '定义第一张图片的行
For Each Item In myArr
        
        'MsgBox Item
    
    'Cells(j, 1).Value = Item
    PicLeft1 = Range("B" & j).Left '图片左边位置，即单元格的左边
	PicTop1 = Range("B" & j).Top  '图片顶部位置，即单元格的顶部
	
    Range("B" & j).Select
	'-1表示图片按原文件大小插入https://msdn.microsoft.com/zh-cn/vba/excel-vba/articles/shapes-addpicture-method-excel
    ActiveSheet.Shapes.AddPicture("D:\test\img\" & Item & "- CPU - last7day.png", False, True, PicLeft1, PicTop1, -1, -1).Select
    PicLeft2 = Range("D" & j).Left '图片左边位置
	PicTop2 = Range("D" & j).Top  '图片顶部位置    
    Range("D" & j).Select
    ActiveSheet.Shapes.AddPicture("D:\test\img\" & Item & " - memory - last7day.png", False, True,PicLeft2, PicTop2, -1, -1).Select
		
    j = j + 2
     
Next


		
