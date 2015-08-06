'指定程序路径在第几列
Const SORPATHNUM As Integer = 1
'指定处理后路径在第几列
Const TARPATHNUM As Integer = 2
'指定是否为java文件在第几列
Const ISJAVANUM As Integer = 3
'指定检查路径结果列
Const CHECKRESULT As Integer = 4
'指定起始行
Const STARTROW As Integer = 11
'指定生成结果列
Const MAKERESULT As Integer = 5

'指定程序路径在第几列
Const SORPATHNUM_ADMIN3 As Integer = 1
'指定处理后路径在第几列
Const TARPATHNUM_ADMIN3 As Integer = 3
'指定程序类型在第几列
Const ISJAVANUM_ADMIN3 As Integer = 4
'指定检查路径结果列
Const CHECKRESULT_ADMIN3 As Integer = 5
'指定生成结果列
Const MAKERESULT_ADMIN3 As Integer = 6
'指定起始行
Const STARTROW_ADMIN3 As Integer = 7

'指定匹配java文件的字符
Const JAVASOR As String = "$javaSource"
'指定匹配非java文件的字符
Const NOTJAVASOR As String = "$notJavaSource"
'指定匹配admin3文件的字符
Const ADMIN3 As String = "$admin3"
'组件标识
Const CMPNAME_CS2002 As String = "Cmp_CTIE-CS2002"
Const CMPNAME_GFT As String = "Cmp_CTIE-GFT"
Const CMPNAME_CMG As String = "Cmp_CTIE-CMG\CTP4.0\CMGEAR"
Const CMPCOMPILE As String = "Cmp_CTIE-COMPILE3"
Const CMPNAME_ADMIN3 As String = "Cmp_CTIE-ADMIN3"

Private cs2002list() As String, gftList() As String, cmgList() As String, admin3list() As String


'生成java全量(其他增量)编译脚本
Private Sub CommandButton1_Click()
    '表格的最大行数
    Dim finalRow As Integer
    If Cells(65536, SORPATHNUM).End(xlUp).row < Cells(65536, TARPATHNUM).End(xlUp).row Then
        finalRow = Cells(65536, TARPATHNUM).End(xlUp).row
    Else
        finalRow = Cells(65536, SORPATHNUM).End(xlUp).row
    End If
    '清空对应单元格日志
    For n = STARTROW To finalRow
        For m = 2 To 10
            Cells(n, m).Value = ""
        Next m
    Next n
    MakeBuildXml_cs2002 "buildCS2002_template.xml", "buildCS2002.xml"
End Sub

'生成全量编译脚本
Private Sub CommandButton2_Click()
    '表格的最大行数
    Dim finalRow As Integer
    If Cells(65536, SORPATHNUM).End(xlUp).row < Cells(65536, TARPATHNUM).End(xlUp).row Then
        finalRow = Cells(65536, TARPATHNUM).End(xlUp).row
    Else
        finalRow = Cells(65536, SORPATHNUM).End(xlUp).row
    End If
    '清空对应单元格日志
    For n = STARTROW To finalRow
        For m = 1 To 10
            Cells(n, m).Value = ""
        Next m
    Next n
    Cells(STARTROW, SORPATHNUM).Value = "\vobs\V_CTIE\Cmp_CTIE-CS2002\**"
    MakeBuildXml_cs2002 "buildCS2002_template.xml", "buildCS2002_full.xml"
End Sub

Private Sub MakeBuildXml_cs2002(buildTemplateVar As String, buildVar As String)
Dim arr() As String, name As String, n As Integer
    '表格的最大行数
    'Dim finalRow As Integer
    'finalRow = Cells(65536, SORPATHNUM).End(xlUp).row
    finalRow = UBound(cs2002list)
    
    'MsgBox ("开始生成编译脚本")
    
    Dim myFile As Object
    Set myFile = CreateObject("scripting.filesystemobject")
    
    Dim des As String
    des = ThisWorkbook.Path
    des = Replace(des, "/", "\")
    des = Replace(des, "\" + CMPCOMPILE, "")
    
     '对每一行循环处理
    'For n = STARTROW To finalRow
    For i = 0 To UBound(cs2002list)
        arr = Split(cs2002list(i), "|")
        n = arr(0)
        name = arr(1)
        '对路径进行处理
        'MsgBox (InStr("Cmp_CTIE-CTIE", Cells(n, SORPATHNUM).Value))
        '输出处理后的路径
        'If InStr(Cells(n, SORPATHNUM).Value, CMPNAME_CS2002) > 0 Then
        If InStr(cs2002list(i), CMPNAME_CS2002) > 0 Then
            'Cells(n, TARPATHNUM).Value = Replace(Right(Cells(n, SORPATHNUM).Value, Len(Cells(n, SORPATHNUM).Value) - InStr(Cells(n, SORPATHNUM).Value, CMPNAME_CS2002) - Len(CMPNAME_CS2002)), "\", "/")
            Cells(n, TARPATHNUM).Value = Replace(Right(name, Len(name) - InStr(name, CMPNAME_CS2002) - Len(CMPNAME_CS2002)), "\", "/")
        Else
        End If
        '判断是否为java文件
        If Len(Cells(n, TARPATHNUM).Value) > 4 Then
            Cells(n, ISJAVANUM).Value = (Mid(Cells(n, TARPATHNUM).Value, Len(Cells(n, TARPATHNUM).Value) - 4, 5) = ".java")
        Else
            Cells(n, ISJAVANUM).Value = False
        End If
        '判断文件路径是否填写正确
        If myFile.fileexists(des + "\" + CMPNAME_CS2002 + "\" + Cells(n, TARPATHNUM).Value) Then
            Cells(n, CHECKRESULT).Value = "检查通过"
        ElseIf InStr(Cells(n, TARPATHNUM).Value, "/**") And myFile.folderexists(des + "\" + CMPNAME_CS2002 + "\" + Mid(Cells(n, TARPATHNUM).Value, 1, Len(Cells(n, TARPATHNUM).Value) - 2)) Then
            Cells(n, CHECKRESULT).Value = "检查通过"
        Else
            Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
        End If
    'Next n
     Next i

    Dim buildTemplate As String
    buildTemplate = des + "\" + CMPNAME_CS2002 + "\" + buildTemplateVar
    Dim build As String
    build = des + "\" + CMPNAME_CS2002 + "\" + buildVar
    '删除旧的build.xml文件
    If myFile.fileexists(build) Then
        Kill build
    End If
    Open buildTemplate For Input As #1
    Open build For Output As #2
    Do While Not EOF(1)
        Dim tmp1, tmp2 As String
        Line Input #1, tmp1
        If InStr(tmp1, JAVASOR) > 0 Then

        ElseIf InStr(tmp1, NOTJAVASOR) > 0 Then
            i = 0
            'For n = STARTROW To finalRow
            For n1 = 0 To finalRow
            
                arr = Split(cs2002list(n1), "|")
                n = arr(0)
        
                If Cells(n, ISJAVANUM).Value = True Then
                    
                Else
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    i = 1
                    'Print #2, Chr(13)
                End If
            Next n1
            If i = 0 Then
                tmp2 = "        <include name=""""/>"
                Print #2, tmp2
            End If
        Else
            tmp2 = tmp1
            Print #2, tmp2
        End If
    Loop
    Close #1
    Close #2
    'MsgBox ("已经生成编译脚本为：" + build)
End Sub

Private Sub MakeBuildXml_gft(buildTemplateVar As String, buildVar As String)
Dim arr() As String, name As String, n As Integer
    '表格的最大行数
'    Dim finalRow As Integer
'    finalRow = Cells(65536, SORPATHNUM).End(xlUp).row
    finalRow = UBound(gftList)
    
    'MsgBox ("开始生成编译脚本")
    
    Dim myFile As Object
    Set myFile = CreateObject("scripting.filesystemobject")
    
    Dim des As String
    des = ThisWorkbook.Path
    des = Replace(des, "/", "\")
    des = Replace(des, "\" + CMPCOMPILE, "")
    
     '对每一行循环处理
    'For n = STARTROW To finalRow
    For i = 0 To UBound(gftList)
        arr = Split(gftList(i), "|")
        n = arr(0)
        name = arr(1)
        '对路径进行处理
        'MsgBox (InStr("Cmp_CTIE-CTIE", Cells(n, SORPATHNUM).Value))
        '输出处理后的路径
        'If InStr(Cells(n, SORPATHNUM).Value, CMPNAME_GFT) > 0 Then
        If InStr(gftList(i), CMPNAME_GFT) > 0 Then
            'Cells(n, TARPATHNUM).Value = Replace(Right(Cells(n, SORPATHNUM).Value, Len(Cells(n, SORPATHNUM).Value) - InStr(Cells(n, SORPATHNUM).Value, CMPNAME_GFT) - Len(CMPNAME_GFT)), "\", "/")
            Cells(n, TARPATHNUM).Value = Replace(Right(name, Len(name) - InStr(name, CMPNAME_GFT) - Len(CMPNAME_GFT)), "\", "/")
        Else
        End If
        '判断是否为java文件
        If Len(Cells(n, TARPATHNUM).Value) > 4 Then
            Cells(n, ISJAVANUM).Value = (Mid(Cells(n, TARPATHNUM).Value, Len(Cells(n, TARPATHNUM).Value) - 4, 5) = ".java")
        Else
            Cells(n, ISJAVANUM).Value = False
        End If
        '判断文件路径是否填写正确
        If myFile.fileexists(des + "\" + CMPNAME_GFT + "\" + Cells(n, TARPATHNUM).Value) Then
            Cells(n, CHECKRESULT).Value = "检查通过"
        ElseIf InStr(Cells(n, TARPATHNUM).Value, "/**") And myFile.folderexists(des + "\" + CMPNAME_GFT + "\" + Mid(Cells(n, TARPATHNUM).Value, 1, Len(Cells(n, TARPATHNUM).Value) - 2)) Then
            Cells(n, CHECKRESULT).Value = "检查通过"
        Else
            Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
        End If
    'Next n
    Next i

    Dim buildTemplate As String
    buildTemplate = des + "\" + CMPNAME_GFT + "\" + buildTemplateVar
    Dim build As String
    build = des + "\" + CMPNAME_GFT + "\" + buildVar
    '删除旧的build.xml文件
    If myFile.fileexists(build) Then
        Kill build
    End If
    Open buildTemplate For Input As #1
    Open build For Output As #2
    Do While Not EOF(1)
        Dim tmp1, tmp2 As String
        Line Input #1, tmp1
        If InStr(tmp1, JAVASOR) > 0 Then

        ElseIf InStr(tmp1, NOTJAVASOR) > 0 Then
            i = 0
            'For n = STARTROW To finalRow
            For n1 = 0 To finalRow
            
                arr = Split(gftList(n1), "|")
                n = arr(0)
        
                If Cells(n, ISJAVANUM).Value = True Then
                    
                Else
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    i = 1
                    'Print #2, Chr(13)
                End If
            Next n1
            If i = 0 Then
                tmp2 = "        <include name=""""/>"
                Print #2, tmp2
            End If
        Else
            tmp2 = tmp1
            Print #2, tmp2
        End If
    Loop
    Close #1
    Close #2
    'MsgBox ("已经生成编译脚本为：" + build)
End Sub

Private Sub MakeBuildXml_cmg(buildTemplateVar As String, buildVar As String)
Dim arr() As String, name As String, n As Integer
    '表格的最大行数
'    Dim finalRow As Integer
'    finalRow = Cells(65536, SORPATHNUM).End(xlUp).row
     finalRow = UBound(cmgList)
    
    'MsgBox ("开始生成编译脚本")
    
    Dim myFile As Object
    Set myFile = CreateObject("scripting.filesystemobject")
    
    Dim des As String
    des = ThisWorkbook.Path
    des = Replace(des, "/", "\")
    des = Replace(des, "\" + CMPCOMPILE, "")
    
     '对每一行循环处理
    'For n = STARTROW To finalRow
    For i = 0 To UBound(cmgList)
        arr = Split(cmgList(i), "|")
        n = arr(0)
        name = arr(1)
        '对路径进行处理
        'MsgBox (InStr("Cmp_CTIE-CTIE", Cells(n, SORPATHNUM).Value))
        '输出处理后的路径
        'If InStr(Replace(Cells(n, SORPATHNUM).Value, "/", "\"), CMPNAME_CMG + "\ctpAuthWeb\JavaSource") > 0 Then
        If InStr(Replace(cmgList(i), "/", "\"), CMPNAME_CMG + "\ctpAuthWeb\JavaSource") > 0 Then
            'Cells(n, TARPATHNUM).Value = Replace(Right(Cells(n, SORPATHNUM).Value, Len(Cells(n, SORPATHNUM).Value) - InStr(Replace(Cells(n, SORPATHNUM).Value, "/", "\"), CMPNAME_CMG) - Len(CMPNAME_CMG + "\ctpAuthWeb\JavaSource")), "\", "/")
            Cells(n, TARPATHNUM).Value = Replace(Right(name, Len(name) - InStr(Replace(name, "/", "\"), CMPNAME_CMG) - Len(CMPNAME_CMG + "\ctpAuthWeb\JavaSource")), "\", "/")
        Else
        End If
        '判断是否为java文件
        If Len(Cells(n, TARPATHNUM).Value) > 4 Then
            Cells(n, ISJAVANUM).Value = (Mid(Cells(n, TARPATHNUM).Value, Len(Cells(n, TARPATHNUM).Value) - 4, 5) = ".java")
        Else
            Cells(n, ISJAVANUM).Value = False
        End If
        '判断文件路径是否填写正确
        If myFile.fileexists(des + "\" + CMPNAME_CMG + "\ctpAuthWeb\JavaSource" + "\" + Cells(n, TARPATHNUM).Value) Then
            Cells(n, CHECKRESULT).Value = "检查通过"
        ElseIf InStr(Cells(n, TARPATHNUM).Value, "/**") And myFile.folderexists(des + "\" + CMPNAME_CMG + "\ctpAuthWeb\JavaSource" + "\" + Mid(Cells(n, TARPATHNUM).Value, 1, Len(Cells(n, TARPATHNUM).Value) - 2)) Then
            Cells(n, CHECKRESULT).Value = "检查通过"
        Else
            Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
        End If
    'Next n
    Next i

    Dim buildTemplate As String
    buildTemplate = des + "\" + CMPNAME_CMG + "\" + buildTemplateVar
    Dim build As String
    build = des + "\" + CMPNAME_CMG + "\" + buildVar
    '删除旧的build.xml文件
    If myFile.fileexists(build) Then
        Kill build
    End If
    Open buildTemplate For Input As #1
    Open build For Output As #2
    Do While Not EOF(1)
        Dim tmp1, tmp2 As String
        Line Input #1, tmp1
        If InStr(tmp1, JAVASOR) > 0 Then
            i = 0
            'For n = STARTROW To finalRow
            For n1 = 0 To finalRow
                arr = Split(cmgList(n1), "|")
                n = arr(0)

                If Cells(n, ISJAVANUM).Value = True Then
                    If i = 0 Then
                        tmp2 = "           includes=""" + Cells(n, TARPATHNUM).Value + ","
                        Print #2, tmp2
                        'Print #2, Chr(13)
                        i = 1
                    Else
                        tmp2 = "             " + Cells(n, TARPATHNUM).Value + ","
                        Print #2, tmp2
                    End If
                Else
                    
                End If
            Next n1
            If i = 0 Then
                tmp2 = "             includes=""$caution:no java compile"""
                Print #2, tmp2
            Else
                tmp2 = "             """
                Print #2, tmp2
            End If
        ElseIf InStr(tmp1, NOTJAVASOR) > 0 Then
            Cells(n, MAKERESULT).Value = "非java程序，请手工加入编译脚本"
        Else
            tmp2 = tmp1
            Print #2, tmp2
        End If
    Loop
    Close #1
    Close #2
    'MsgBox ("已经生成编译脚本为：" + build)
End Sub

Private Sub MakeBuildXml_admin3(buildTemplateVar As String, buildVar As String)
Dim arr() As String, name As String, n As Integer
    '表格的最大行数
'    Dim finalRow As Integer
'    finalRow = Cells(65536, SORPATHNUM_ADMIN3).End(xlUp).row
     finalRow = UBound(admin3list)
    
    'MsgBox ("开始生成编译脚本")
    
    Dim myFile As Object
    Set myFile = CreateObject("scripting.filesystemobject")
    
    Dim des As String
    des = ThisWorkbook.Path
    des = Replace(des, "/", "\")
    des = Replace(des, "\" + CMPCOMPILE, "")

    Dim tmpString As String
    
     '对每一行循环处理
   ' For n = STARTROW_ADMIN3 To finalRow
    For i = 0 To UBound(admin3list)
        arr = Split(admin3list(i), "|")
        n = arr(0)
        name = arr(1)
        '对路径进行处理
        '输出处理后的路径
        'If InStr(Cells(n, SORPATHNUM_ADMIN3).Value, CMPNAME_ADMIN3) > 0 Then
        If InStr(admin3list(i), CMPNAME_ADMIN3) > 0 Then
            'tmpString = Replace(Right(Cells(n, SORPATHNUM_ADMIN3).Value, Len(Cells(n, SORPATHNUM_ADMIN3).Value) - InStr(Cells(n, SORPATHNUM_ADMIN3).Value, CMPNAME_ADMIN3) - Len(CMPNAME_ADMIN3)), "\", "/")
            tmpString = Replace(Right(name, Len(name) - InStr(name, CMPNAME_ADMIN3) - Len(CMPNAME_ADMIN3)), "\", "/")
        Else
        End If
        '判断是否为java文件tmpString
        If InStr(tmpString, "ctieadmin/src/") > 0 Then
            If (Mid(tmpString, Len(tmpString) - 4, 5) = ".java") Then
                Cells(n, ISJAVANUM_ADMIN3).Value = "com_java"
            Else
                Cells(n, ISJAVANUM_ADMIN3).Value = "com_notjava"
            End If
        Else
            Cells(n, ISJAVANUM_ADMIN3).Value = "others"
        End If
        '判断文件路径是否填写正确
        If myFile.fileexists(des + "\" + CMPNAME_ADMIN3 + "\" + tmpString) Then
            Cells(n, CHECKRESULT_ADMIN3).Value = "检查通过"
        ElseIf InStr(tmpString, "/**") And myFile.folderexists(des + "\" + CMPNAME_ADMIN3 + "\" + Mid(tmpString, 1, Len(tmpString) - 2)) Then
            Cells(n, CHECKRESULT_ADMIN3).Value = "检查通过"
        Else
            Cells(n, CHECKRESULT_ADMIN3).Value = "检查不通过，文件不存在，请检查填写是否正确"
        End If
        If Cells(n, CHECKRESULT_ADMIN3).Value = "检查通过" Then
            If InStr(tmpString, "ctieadmin/WebContent/") > 0 Then
                Cells(n, TARPATHNUM_ADMIN3).Value = Replace(tmpString, "ctieadmin/WebContent/", "")
            ElseIf InStr(tmpString, "ctieadmin/src/") > 0 Then
                If Cells(n, ISJAVANUM_ADMIN3).Value = "com_java" Then
                    Cells(n, TARPATHNUM_ADMIN3).Value = Replace(Replace(tmpString, "ctieadmin/src/", ""), ".java", "")
                Else
                    Cells(n, TARPATHNUM_ADMIN3).Value = Replace(tmpString, "ctieadmin/src/", "")
                End If
            Else
                Cells(n, TARPATHNUM_ADMIN3).Value = "'"
            End If
        Else
            Cells(n, TARPATHNUM_ADMIN3).Value = ""
        End If
    'Next n
    Next i

    Dim buildTemplate As String
    buildTemplate = des + "\" + CMPNAME_ADMIN3 + "\ctieadmin\" + buildTemplateVar
    Dim build As String
    build = des + "\" + CMPNAME_ADMIN3 + "\ctieadmin\" + buildVar
    '删除旧的build.xml文件
    If myFile.fileexists(build) Then
        Kill build
    End If
    Open buildTemplate For Input As #1
    Open build For Output As #2
    Do While Not EOF(1)
        Dim tmp1, tmp2 As String
        Line Input #1, tmp1
        If InStr(tmp1, ADMIN3) > 0 Then
            'For n = STARTROW_ADMIN3 To finalRow
            For n1 = 0 To finalRow
                arr = Split(admin3list(n1), "|")
                n = arr(0)
        
                If Cells(n, CHECKRESULT_ADMIN3).Value = "检查通过" And Cells(n, TARPATHNUM_ADMIN3).Value <> "" Then
                    If Cells(n, ISJAVANUM_ADMIN3).Value = "com_java" Then
                        tmp2 = "        <include name=""" + "WEB-INF/classes/" + Cells(n, TARPATHNUM_ADMIN3).Value + ".class" + """/>"
                        Print #2, tmp2
                        tmp2 = "        <include name=""" + "WEB-INF/classes/" + Cells(n, TARPATHNUM_ADMIN3).Value + "$*.class" + """/>"
                        Print #2, tmp2
                    ElseIf Cells(n, ISJAVANUM_ADMIN3).Value = "com_notjava" Then
                        tmp2 = "        <include name=""" + "WEB-INF/classes/" + Cells(n, TARPATHNUM_ADMIN3).Value + """/>"
                        Print #2, tmp2
                    Else
                        tmp2 = "        <include name=""" + Cells(n, TARPATHNUM_ADMIN3).Value + """/>"
                        Print #2, tmp2
                    End If
                End If
            Next n1
        Else
            tmp2 = tmp1
            Print #2, tmp2
        End If
    Loop
    Close #1
    Close #2
    'MsgBox ("已经生成编译脚本为：" + build)
End Sub

Private Sub CommandButton3_Click()
Dim rowValue As String, row As Integer, k1 As Integer, k2 As Integer, k3 As Integer, k4 As Integer
    '表格的最大行数
    Dim finalRow As Integer
    If Cells(65536, SORPATHNUM).End(xlUp).row < Cells(65536, TARPATHNUM).End(xlUp).row Then
        finalRow = Cells(65536, TARPATHNUM).End(xlUp).row
    Else
        finalRow = Cells(65536, SORPATHNUM).End(xlUp).row
    End If
    '清空对应单元格日志
    For n = STARTROW To finalRow
        For m = 2 To 10
            Cells(n, m).Value = ""
        Next m
    Next n
    'MakeBuildXml_cs2002 "buildCS2002_template.xml", "buildCS2002.xml"
    '将数据分类存到数据中
    row = STARTROW
    k1 = 0
    k2 = 0
    k3 = 0
    k4 = 0
    Do
       rowValue = Trim(Cells(row, SORPATHNUM).Value)  '从第11行第1列开始遍历
       If rowValue = "" Then Exit Do
       
       If InStr(rowValue, CMPNAME_CS2002) > 0 Then
            ReDim Preserve cs2002list(k1)
            cs2002list(k1) = row & "|" & rowValue
            k1 = k1 + 1
       End If
       
       If InStr(rowValue, CMPNAME_GFT) > 0 Then
            ReDim Preserve gftList(k2)
            gftList(k2) = row & "|" & rowValue
            k2 = k2 + 1
       End If
       
       If InStr(rowValue, CMPNAME_CMG) > 0 Then
            ReDim Preserve cmgList(k3)
            cmgList(k3) = row & "|" & rowValue
            k3 = k3 + 1
       End If
       
       If InStr(rowValue, CMPNAME_ADMIN3) > 0 Then
            ReDim Preserve admin3list(k4)
            admin3list(k4) = row & "|" & rowValue
            k4 = k4 + 1
       End If
       
       row = row + 1
    Loop
    
    MsgBox ("开始生成编译脚本")
    If k1 > 0 Then
        'MsgBox UBound(cs2002list)
         MakeBuildXml_cs2002 "buildCS2002_template.xml", "buildCS2002.xml"
    End If
    
    If k2 > 0 Then
       ' MsgBox UBound(gftList)
        MakeBuildXml_gft "buildGFT_template.xml", "buildGFT.xml"
    End If
    
    If k3 > 0 Then
        'MsgBox UBound(cmgList)
        MakeBuildXml_cmg "buildCMGEAR_template.xml", "buildCMGEAR.xml"
    End If
    
    If k4 > 0 Then
        'MsgBox UBound(cmgList)
        MakeBuildXml_admin3 "buildADMIN3_template.xml", "buildADMIN3.xml"
    End If
    
    MsgBox ("编译完成")

End Sub
















