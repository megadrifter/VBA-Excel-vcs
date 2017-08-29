Attribute VB_Name = "Sketches"
Global blinkrunning


Sub blink_loader()
If blinkrunning = 1 Then
    Range("A1").Value = 1
    blinkrunning = 0
Else
    Blink
End If
End Sub


Sub Blink()
blinkrunning = 1
Set cell_01 = Range("A1")
Set cell_02 = Range("B1")
Set cell_03 = Range("C1")
lasttime = millis()
cell_01.Interior.Color = 65535
cell_02.Interior.Color = 65535


' main loop
While (1)
    nowtime = millis()
    flag_strobo = cell_03.Value
    
    
    'processing flag_strobo
    If (flag_strobo = 1) Then
        lag = (nowtime - lasttime)
        'manage first cell
        If lag > 900 Then
        cell_01.Interior.Color = 65535
        ElseIf lag > 600 Then
        cell_01.Interior.Color = 65535
        ElseIf lag > 400 Then
        cell_01.Interior.Color = 255
        ElseIf lag > 200 Then
        cell_01.Interior.Color = 65535
        Else
        cell_01.Interior.Color = 255
        End If
        
        'manage second cell
        If lag > 1800 Then
        cell_02.Interior.Color = 65535
        lasttime = nowtime  ' reset cycle
        ElseIf lag > 1400 Then
        cell_02.Interior.Color = 65535
        ElseIf lag > 1200 Then
        cell_02.Interior.Color = 255
        ElseIf lag > 1000 Then
        cell_02.Interior.Color = 65535
        ElseIf lag > 800 Then
        cell_02.Interior.Color = 255
        End If
    Else
    cell_01.Interior.Color = 65535
    cell_02.Interior.Color = 65535
    End If
    
    
    
    
DoEvents ' not to hang up Excel
If cell_01.Value = 1 Then
    cell_01.Value = ""
    Exit Sub
End If
Wend
End Sub


