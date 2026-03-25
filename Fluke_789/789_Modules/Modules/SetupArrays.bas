Attribute VB_Name = "SetupArrays"
'Sub ArraySetup()

 
    
 '   ranges = Array("16:19", "21:23", "27:30", "34:36", "38:40", "44:46", "48:51", "55:55", "59:59", "63:65", "67:69", "71:73", "77:79", "82:84", "87:89", "93:95", "97:99", "103:105", "107:109", "111:114", "118:120", "124:126", "130:132", "136:138", "142:142", "143:143", "144:144")
  '  ranges2 = Array()
   ' ranges3 = Array()
    

'End Sub


Sub ArraySetup()
    Set ranges = CreateObject("Scripting.Dictionary")
    
    
    ranges(Tab1) = Array("13:14", "16:24", "27:29", "32:39", "41:42", "45:46", "48:54", "56:57", "59:60", "62:63", "66:67", "70:72", "75:76", "79:80", "83:85", "88:88", "89:89", "90:90", "91:91")
    ranges(Tab2) = Array()
    ranges(Tab3) = Array()
    ranges(Tab4) = Array()
End Sub
