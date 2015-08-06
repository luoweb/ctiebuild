


'指定程序路径在第几列
Const SORPATHNUM As Integer = 1
'指定处理后路径在第几列
Const TARPATHNUM As Integer = 3
'指定程序类型在第几列
Const SORTYPE As Integer = 4
'指定检查路径结果列
Const CHECKRESULT As Integer = 5
'指定生成结果列
Const MAKERESULT As Integer = 6
'指定起始行
Const STARTROW As Integer = 7
'指定匹配plugins的字符
Const PLUGINS As String = "$plugins"
'指定匹配plugins_COM的字符
Const PLUGINSCOM As String = "$plugin_com"
'指定CTIE3其他的字符
Const CTIE3OTHERS As String = "$ctie3others"
'指定dbinstall的字符
Const DBINSTALL As String = "$dbinstall"
'指定solution的字符
Const SOLUTION As String = "$solution"
'指定device的字符
Const DEVICE As String = "$device"
'指定DB的字符
Const DB As String = "$db"
'指定script的字符
Const SCRIPT As String = "$script"
'指定config的字符
Const CONFIG As String = "$config"
'组件标识
Const CMPNAME As String = "Cmp_CTIE-CTIE3"
Const CMPCOMPILE As String = "Cmp_CTIE-COMPILE"

'生成增量编译脚本
Private Sub CommandButton1_Click()
    '表格的最大行数
    Dim finalRow As Integer
    If Cells(65536, SORPATHNUM).End(xlUp).Row < Cells(65536, TARPATHNUM).End(xlUp).Row Then
        finalRow = Cells(65536, TARPATHNUM).End(xlUp).Row
    Else
        finalRow = Cells(65536, SORPATHNUM).End(xlUp).Row
    End If
    
    '清空对应单元格日志
    For n = STARTROW To finalRow
        For m = 3 To 10
            Cells(n, m).Value = ""
        Next m
    Next n
    MakeBuildXml "buildCTIE3_template.xml", "buildCTIE3.xml"
End Sub

'生成全量编译脚本
Private Sub CommandButton4_Click()
    '表格的最大行数
    Dim finalRow As Integer
    If Cells(65536, SORPATHNUM).End(xlUp).Row < Cells(65536, TARPATHNUM).End(xlUp).Row Then
        finalRow = Cells(65536, TARPATHNUM).End(xlUp).Row
    Else
        finalRow = Cells(65536, SORPATHNUM).End(xlUp).Row
    End If
    '清空对应单元格日志
    For n = STARTROW To finalRow
        For m = 1 To 10
            Cells(n, m).Value = ""
        Next m
    Next n
    MakeBuildXml "buildCTIE3_template.xml", "buildCTIE3_full.xml"
End Sub

Private Sub MakeBuildXml(buildTemplateVar As String, buildVar As String)
    '表格的最大行数
    Dim finalRow As Integer
    finalRow = Cells(65536, SORPATHNUM).End(xlUp).Row
    
    MsgBox ("开始生成编译脚本")
    
    Dim myFile As Object
    Set myFile = CreateObject("scripting.filesystemobject")
    
    Dim des As String
    des = ThisWorkbook.Path
    des = Replace(des, "/", "\")
    des = Replace(des, "\" + "Cmp_CTIE-COMPILE3", "")
    MsgBox des
    
    Dim tmpString, tmpString1 As String
    tmpString = ""
    tmpString1 = ""
     '对每一行循环处理
    For n = STARTROW To finalRow
        '对路径进行处理
        'MsgBox (InStr("Cmp_CTIE-CTIE", Cells(n, SORPATHNUM).Value))
         '判断文件路径是否填写正确
         
        tmpString = Replace(Cells(n, SORPATHNUM).Value, "\", "/")

        If InStr(tmpString, "Cmp_CTIE-CTIE3/") Then
            tmpString = Right(tmpString, Len(tmpString) - InStr(tmpString, "Cmp_CTIE-CTIE3/") + 1)
            If InStr(tmpString, "Cmp_CTIE-CTIE3/src/") Then
                If (InStr(Mid(tmpString, 20), "/") = 0) And myFile.folderexists(des + "/" + tmpString) Then
                    Cells(n, CHECKRESULT).Value = "检查通过"
                    Cells(n, SORTYPE).Value = "plugins"
                    tmpString1 = tmpString
                    Cells(n, TARPATHNUM).Value = "plugins" + Replace(tmpString1, "Cmp_CTIE-CTIE3/src", "") + "_*"
                ElseIf myFile.fileexists(des + "/" + tmpString) Then
                    If InStr(tmpString, "Cmp_CTIE-CTIE3/src/com.icbc/src/com/") Then
                        Cells(n, CHECKRESULT).Value = "检查通过"
                        Cells(n, SORTYPE).Value = "plugins_com"
                        tmpString1 = tmpString
                        tmpString1 = Replace(tmpString1, "Cmp_CTIE-CTIE3/src/com.icbc/src/", "")
                        tmpString1 = Replace(tmpString1, ".java", "")
                        Cells(n, TARPATHNUM).Value = tmpString1
                    Else
                        Cells(n, CHECKRESULT).Value = "检查通过"
                        Cells(n, SORTYPE).Value = "plugins"
                        tmpString1 = tmpString
                        tmpString1 = Replace(tmpString1, "Cmp_CTIE-CTIE3/src/", "")
                        tmpString1 = Mid(tmpString1, 1, InStr(tmpString1, "/") - 1)
                        Cells(n, TARPATHNUM).Value = "plugins/" + tmpString1 + "_*"
                    End If
                ElseIf InStr(tmpString, "Cmp_CTIE-CTIE3/src/com.icbc/src/com/") Then
                    If InStr(tmpString, "/**") Then
                        tmpString1 = tmpString
                        tmpString1 = Replace(tmpString1, "/**", "")
                        If myFile.folderexists(des + "/" + tmpString1) Then
                            Cells(n, CHECKRESULT).Value = "检查通过"
                            Cells(n, SORTYPE).Value = "plugins_com"
                            tmpString1 = Replace(tmpString1, "Cmp_CTIE-CTIE3/src/com.icbc/src/", "")
                            Cells(n, TARPATHNUM).Value = tmpString1 + "/**"
                        Else
                            Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
                        End If
                    Else
                        Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
                    End If
                Else
                    Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
                End If
            ElseIf InStr(tmpString, "Cmp_CTIE-CTIE3/dbinstall/") Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "dbinstall"
                Cells(n, TARPATHNUM).Value = "install/create_config.jar"
            ElseIf InStr(tmpString, "Cmp_CTIE-CTIE3/config/") And myFile.fileexists(des + "/" + tmpString) Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "config"
                tmpString1 = tmpString
                Cells(n, TARPATHNUM).Value = Replace(tmpString1, "Cmp_CTIE-CTIE3/", "")
            ElseIf InStr(tmpString, "Cmp_CTIE-CTIE3/DB/") Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "DB"
                Cells(n, TARPATHNUM).Value = "DB/**"
            ElseIf InStr(tmpString, "Cmp_CTIE-CTIE3/script/") Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "script"
                Cells(n, TARPATHNUM).Value = "script/**"
            'ElseIf InStr(tmpString, "Cmp_CTIE-CTIE3/NCTS/plugins") Or InStr(tmpString, "Cmp_CTIE-CTIE3/NCTB/plugins") Then
                'Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
            ElseIf InStr(tmpString, "Cmp_CTIE-CTIE3/NCTS/resource_update/ctie") Or InStr(tmpString, "Cmp_CTIE-CTIE3/NCTB/DeviceServer") Then
                Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
            ElseIf myFile.fileexists(des + "/" + tmpString) Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "ctie3others"
                tmpString1 = tmpString
                Cells(n, TARPATHNUM).Value = Replace(tmpString1, "Cmp_CTIE-CTIE3/", "")
            ElseIf InStr(tmpString, "/**") And myFile.folderexists(des + "/" + Mid(tmpString, 1, Len(tmpString) - 2)) Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "ctie3others"
                tmpString1 = tmpString
                Cells(n, TARPATHNUM).Value = Replace(tmpString1, "Cmp_CTIE-CTIE3/", "")
            Else
                Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
            End If
        ElseIf InStr(tmpString, "Cmp_CTIE-DEVICE3/DeviceServer/result/") Then
            tmpString = Right(tmpString, Len(tmpString) - InStr(tmpString, "Cmp_CTIE-DEVICE3/") + 1)
            If myFile.fileexists(des + "/" + tmpString) Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "device3"
                tmpString1 = tmpString
                Cells(n, TARPATHNUM).Value = Replace(tmpString1, "Cmp_CTIE-DEVICE3/DeviceServer/result", "NCTB/DeviceServer")
            ElseIf InStr(tmpString, "/**") And myFile.folderexists(des + "/" + Mid(tmpString, 1, Len(tmpString) - 2)) Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "device3"
                tmpString1 = tmpString
                Cells(n, TARPATHNUM).Value = Replace(tmpString1, "Cmp_CTIE-DEVICE3/DeviceServer/result", "NCTB/DeviceServer")
            Else
                Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确1"
            End If
        ElseIf InStr(tmpString, "Cmp_CTIE-SOLUTION3/ctie/") Then
            tmpString = Right(tmpString, Len(tmpString) - InStr(tmpString, "Cmp_CTIE-SOLUTION3/") + 1)
                        'Cells(n, 5).Value = tmpString
            If myFile.fileexists(des + "/" + tmpString) Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                If Len(tmpString) > 4 Then
                    If (Mid(tmpString, Len(tmpString) - 4, 5) = ".java") Then
                        Cells(n, SORTYPE).Value = "solution3_java"
                        tmpString1 = tmpString
                        Cells(n, TARPATHNUM).Value = "NCTS/resource_update/" + Replace(Replace(tmpString1, "Cmp_CTIE-SOLUTION3/", ""), ".java", "")
                    Else
                        Cells(n, SORTYPE).Value = "solution3_notjava"
                        tmpString1 = tmpString
                        Cells(n, TARPATHNUM).Value = "NCTS/resource_update/" + Replace(tmpString1, "Cmp_CTIE-SOLUTION3/", "")
                    End If
                Else
                    Cells(n, SORTYPE).Value = "solution3_notjava"
                    tmpString1 = tmpString
                    Cells(n, TARPATHNUM).Value = "NCTS/resource_update/" + Replace(tmpString1, "Cmp_CTIE-SOLUTION3/", "")
                End If
            ElseIf InStr(tmpString, "/**") And myFile.folderexists(des + "\" + Mid(tmpString, 1, Len(tmpString) - 2)) Then
                Cells(n, CHECKRESULT).Value = "检查通过"
                Cells(n, SORTYPE).Value = "solution3_notjava"
                tmpString1 = tmpString
                Cells(n, TARPATHNUM).Value = "NCTS/resource_update/" + Replace(tmpString1, "Cmp_CTIE-SOLUTION3/", "")
            Else
                Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
            End If
        ElseIf myFile.folderexists(des + "/Cmp_CTIE-CTIE3/src/" + tmpString) And InStr(tmpString, "/") = 0 Then
            Cells(n, CHECKRESULT).Value = "检查通过"
            Cells(n, SORTYPE).Value = "plugins"
            Cells(n, TARPATHNUM).Value = "plugins" + "/" + tmpString + "_*"
        Else
            Cells(n, CHECKRESULT).Value = "检查不通过，文件不存在，请检查填写是否正确"
        End If
    Next n
    
    Dim buildTemplate As String
    buildTemplate = des + "/" + "Cmp_CTIE-COMPILE3" + "/" + buildTemplateVar
    Dim build As String
    build = des + "/" + "Cmp_CTIE-COMPILE3" + "/" + buildVar
    '删除旧的build.xml文件
    If myFile.fileexists(build) Then
        Kill build
    End If
    Open buildTemplate For Input As #1
    Open build For Output As #2
    Do While Not EOF(1)
        Dim tmp1, tmp2, tmp3 As String
        Dim i As Integer
        Line Input #1, tmp1
        If InStr(tmp1, PLUGINS) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "plugins" Then
                    tmp2 = "        <include name=""" + "NCTS/" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    tmp2 = "        <include name=""" + "NCTB/" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "NCTS&NCTB"
                Else
                    
                End If
            Next n
        ElseIf InStr(tmp1, PLUGINSCOM) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "plugins_com" Then
                    If InStr(Cells(n, TARPATHNUM).Value, "/**") Then
                        tmp2 = "        <include name=""" + "NCTS/" + Cells(n, TARPATHNUM).Value + """/>"
                        Print #2, tmp2
                        Cells(n, MAKERESULT).Value = "NCTS"
                    Else
                        tmp2 = "        <include name=""" + "NCTS/" + Cells(n, TARPATHNUM).Value + ".class_e" + """/>"
                        Print #2, tmp2
                        tmp2 = "        <include name=""" + "NCTS/" + Cells(n, TARPATHNUM).Value + "$*.class_e" + """/>"
                        Print #2, tmp2
                        Cells(n, MAKERESULT).Value = "NCTS"
                    End If
                Else

                End If
            Next n
        ElseIf InStr(tmp1, CTIE3OTHERS) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "ctie3others" And (InStr(Cells(n, TARPATHNUM).Value, "NCTS") = 1 Or InStr(Cells(n, TARPATHNUM).Value, "NCTB") = 1) Then
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "NCTS/NCTB"
                Else
                    
                End If
            Next n
        ElseIf InStr(tmp1, DEVICE) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "device3" Then
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "NCTB"
                Else
                    
                End If
            Next n
        ElseIf InStr(tmp1, DBINSTALL) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "dbinstall" Then
                    tmp2 = "        <include name=""" + "NCTS/" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    tmp2 = "        <include name=""" + "NCTB/" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "NCTS&NCTB"
                    Exit For
                Else
                    
                End If
            Next n

        ElseIf InStr(tmp1, SOLUTION) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "solution3_java" Then
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + ".class" + """/>"
                    Print #2, tmp2
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + "$*.class" + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "NCTS"
                ElseIf Cells(n, SORTYPE).Value = "solution3_notjava" Then
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "NCTS"
                Else

                End If
            Next n
         ElseIf InStr(tmp1, CONFIG) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "config" Then
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "config"
                Else

                End If
            Next n
         ElseIf InStr(tmp1, DB) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "DB" Then
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "DB"
                    Exit For
                Else

                End If
            Next n
          ElseIf InStr(tmp1, SCRIPT) > 0 Then
            For n = STARTROW To finalRow
                If Cells(n, SORTYPE).Value = "script" Then
                    tmp2 = "        <include name=""" + Cells(n, TARPATHNUM).Value + """/>"
                    Print #2, tmp2
                    Cells(n, MAKERESULT).Value = "script"
                Else

                End If
            Next n
          Else
            tmp2 = tmp1
            Print #2, tmp2
        End If
    Loop
    Close #1
    Close #2
    MsgBox ("已经生成编译脚本为：" + build)
End Sub

