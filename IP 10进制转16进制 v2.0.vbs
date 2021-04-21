'拿到用户字符串 进行字符串分割
'判断分割后的字符串是否是四组字符串
'分别对四组字符串进行字符串转数值类型
'对四个数字进行十六进制的转化
'将四组十六进制数字转成字符串
'拼接四个字符串
'@Author 粥蛋
'@Date 20201210

Dim ipInputText     '用户输入的值
Dim AppTitle        '程序标题
Dim IpArrayStr      'ip分割后的数组 字符串类型
Dim IpArrayInt(3)   'ip分割后的数组 数值类型
Dim IpStrH          'ip转化后的16进制拼接的字符串

AppInit()
AppStart()

Sub AppInit()
    AppTitle = "IP地址转16进制小工具 2.0"
End Sub

Sub AppStart()
    ipInputText = InputBox("请输入你的ip",AppTitle)
    ipInputText = Replace(ipInputText," ","")
    If ipInputText = "" Then
    Else
        IpFormat()
    End If
End Sub

Sub IpFormat()
    IpArrayStr = Split(ipInputText,".")
    If UBound(IpArrayStr)+1 = 4 Then
        For i=0 To 3 Step 1
            IpArrayInt(i) = CInt(IpArrayStr(i))
        Next
        ArrayD2H()
    End If
End Sub

Sub ArrayD2H()
    For i=0 To 3 Step 1
        '2.0 升级说明
        '在这里需要判断一下 16进制转后字符串长度
        '如果字符串长度等于1 则字符串前补位0
        If Len(CStr(Hex(IpArrayInt(i)))) = 1 Then
            IpStrH = IpStrH & "0" & CStr(Hex(IpArrayInt(i)))
        Else
            IpStrH = IpStrH & CStr(Hex(IpArrayInt(i)))
        End If
    Next
    result = InputBox("计算结果",AppTitle,IpStrH)
End Sub