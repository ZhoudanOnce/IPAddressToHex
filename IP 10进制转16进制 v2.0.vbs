'�õ��û��ַ��� �����ַ����ָ�
'�жϷָ����ַ����Ƿ��������ַ���
'�ֱ�������ַ��������ַ���ת��ֵ����
'���ĸ����ֽ���ʮ�����Ƶ�ת��
'������ʮ����������ת���ַ���
'ƴ���ĸ��ַ���
'@Author �൰
'@Date 20201210

Dim ipInputText     '�û������ֵ
Dim AppTitle        '�������
Dim IpArrayStr      'ip�ָ������� �ַ�������
Dim IpArrayInt(3)   'ip�ָ������� ��ֵ����
Dim IpStrH          'ipת�����16����ƴ�ӵ��ַ���

AppInit()
AppStart()

Sub AppInit()
    AppTitle = "IP��ַת16����С���� 2.0"
End Sub

Sub AppStart()
    ipInputText = InputBox("���������ip",AppTitle)
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
        '2.0 ����˵��
        '��������Ҫ�ж�һ�� 16����ת���ַ�������
        '����ַ������ȵ���1 ���ַ���ǰ��λ0
        If Len(CStr(Hex(IpArrayInt(i)))) = 1 Then
            IpStrH = IpStrH & "0" & CStr(Hex(IpArrayInt(i)))
        Else
            IpStrH = IpStrH & CStr(Hex(IpArrayInt(i)))
        End If
    Next
    result = InputBox("������",AppTitle,IpStrH)
End Sub