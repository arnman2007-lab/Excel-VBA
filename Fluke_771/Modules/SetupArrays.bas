Attribute VB_Name = "SetupArrays"
'Sub ArraySetup()

 
    
 '   ranges = Array("16:19", "21:23", "27:30", "34:36", "38:40", "44:46", "48:51", "55:55", "59:59", "63:65", "67:69", "71:73", "77:79", "82:84", "87:89", "93:95", "97:99", "103:105", "107:109", "111:114", "118:120", "124:126", "130:132", "136:138", "142:142", "143:143", "144:144")
  '  ranges2 = Array()
   ' ranges3 = Array()
    

'End Sub


Sub ArraySetup()
    Set ranges = CreateObject("Scripting.Dictionary")

    ' Fluke 771 Milliamp Process Clamp Meter
    ' Row 14-17: Operational checks (Backlight, Display, Keypad, Spotlight)
    ' Row 20-25: DC Current 20.99 mA range (±4, ±12, ±20 mA)
    ' Row 27-28: DC Current 99.9 mA range (±100 mA)

    ranges(Tab1) = Array("14:17", "20:25", "27:28")
    ranges(Tab2) = Array()
    ranges(Tab3) = Array()
    ranges(Tab4) = Array()
End Sub
