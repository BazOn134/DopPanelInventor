Imports Inventor
Imports System.Windows.Forms
'Imports System.IO

Module MCopys

    Dim gl_OpenFiles(0) As String

    Public Sub InsexCopys()
        If Not gl_OpenFiles Is Nothing Then gl_OpenFiles = Nothing
        If m_inventorApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then MsgBox("Правило необходимо запускать в сборке", 48, "Ошибка") : Exit Sub
        Dim FFName As String = m_inventorApplication.ActiveDocument.FullFileName 'Left(m_inventorApplication.ActiveDocument.FullFileName, InStrRev(m_inventorApplication.ActiveDocument.FullFileName, "\"))
        Dim OldFFName As String
        Dim OldDisplName As String
        Dim NewFFName As String
        Dim PapkaCBOR As String = Left(FFName, FFName.LastIndexOf("\") + 1)
        Dim ImjaCBOR As String = Mid(FFName, FFName.LastIndexOf("\") + 1)

        MsgBox("FFName= " & FFName & Chr(13) & "PapkaCBOR= " & PapkaCBOR & Chr(13) & "ImjaCBOR= " & ImjaCBOR) '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        Dim sIndex As String = InputBox("Введите разделитель и индекс", "Ввод данных", "-11") : If sIndex = "" Then MsgBox("Индекс не выбран. Работа прекращена", 48, "Отмена") : Exit Sub

        For Each oOcur In m_inventorApplication.ActiveDocument.SelectSet
            'Try
            Dim ADocType As Object : ADocType = oOcur.Type
            If ADocType = 100669440 Or ADocType = 100669696 Then
                If oOcur.ParentComponents.Item(1).Type = 100669440 Or oOcur.ParentComponents.Item(1).Type = 100669696 Then 'oOcur = oOcur.ParentComponents.Item(i)
                    For i = 1 To oOcur.ParentComponents.Item(1).ParentComponents.Count
                        OldDisplName = oOcur.ParentComponents.Item(1).ParentComponents.Item(i).Name
                        OldFFName = oOcur.ParentComponents.Item(1).ParentComponents.Item(i).ReferencedFileDescriptor.FullFileName
                        Zamena(PapkaCBOR, OldFFName, OldDisplName, sIndex)
                        ''=====================================================
                        'MsgBox("перекрас нач111")
                        'Dim oCurrentRenderStyle As RenderStyle = oOcur.RenderStyle
                        'If Not oCurrentRenderStyle Is Nothing Then
                        '    oOcur.RenderStyle = Nothing
                        'Else
                        '    Dim oNewRenderStyle As RenderStyle
                        '    oNewRenderStyle = oOcur.RenderStyles.Item("Темно-красный")
                        '    oOcur.RenderStyle = oNewRenderStyle
                        'End If
                        'MsgBox("перекрас кон111")
                        ''=====================================================
                    Next i
                Else
                    For i = 1 To oOcur.ParentComponents.Count
                        OldDisplName = oOcur.ParentComponents.Item(i).Name
                        OldFFName = oOcur.ParentComponents.Item(i).ReferencedFileDescriptor.FullFileName
                        Zamena(PapkaCBOR, OldFFName, OldDisplName, sIndex)
                        ''=====================================================
                        'MsgBox("перекрас нач222")
                        'Dim oCurrentRenderStyle As RenderStyle = oOcur.RenderStyle
                        'If Not oCurrentRenderStyle Is Nothing Then
                        '    oOcur.RenderStyle = Nothing
                        'Else
                        '    Dim oNewRenderStyle As RenderStyle
                        '    oNewRenderStyle = oOcur.RenderStyles.Item("Темно-красный")
                        '    oOcur.RenderStyle = oNewRenderStyle
                        'End If
                        'MsgBox("перекрас кон222")
                        ''=====================================================
                    Next i
                End If
            End If
            If ADocType = ObjectTypeEnum.kComponentOccurrenceObject Then
                OldDisplName = oOcur.Name
                OldFFName = oOcur.Definition.Document.FullFileName
                NewFFName = Zamena(PapkaCBOR, OldFFName, OldDisplName, sIndex)
                ''=====================================================
                'MsgBox("перекрас нач333")
                'Dim oCurrentRenderStyle As RenderStyle = oOcur.RenderStyle
                'If Not oCurrentRenderStyle Is Nothing Then
                '    oOcur.RenderStyle = Nothing
                'Else
                '    Dim oNewRenderStyle As RenderStyle
                '    oNewRenderStyle = oOcur.RenderStyles.Item("Темно-красный")
                '    oOcur.RenderStyle = oNewRenderStyle
                'End If
                'MsgBox("перекрас кон333")
                ''=====================================================
            End If
        Next
        If MsgBox("Открыть переименованные?", vbYesNo, "Вопрос") = vbYes Then
            For iii = 0 To UBound(gl_OpenFiles)
                If Right(gl_OpenFiles(iii), 3) = "idw" Then Clipboard.SetText(gl_OpenFiles(iii - 1), TextDataFormat.Text)
                m_inventorApplication.Documents.Open(gl_OpenFiles(iii), True)
            Next iii
        End If
    End Sub

    Public Sub NameCopys(Optional ByVal DelStaroeVse As Boolean = False, Optional ByVal DelStaroe As Boolean = False)
        'If Not gl_OpenFiles Is Nothing Then gl_OpenFiles = Nothing
        ReDim gl_OpenFiles(0)
        Dim Udalit As Boolean = False
        If DelStaroeVse = True Then Udalit = True
        'If MsgBox("Удалять скопированые файлы?", vbYesNoCancel, "Вопрос") = vbYes Then DelStaroe = True
        If m_inventorApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then MsgBox("Правило необходимо запускать в сборке", 48, "Ошибка") : Exit Sub
        Dim FFName As String = m_inventorApplication.ActiveDocument.FullFileName 'Left(m_inventorApplication.ActiveDocument.FullFileName, InStrRev(m_inventorApplication.ActiveDocument.FullFileName, "\"))
        Dim OldFFName As String
        Dim OldDisplName As String
        'Dim NewFFName As String
        Dim PapkaCBOR As String = Left(FFName, FFName.LastIndexOf("\") + 1)
        Dim ImjaCBOR As String = Mid(FFName, FFName.LastIndexOf("\") + 1)
        For Each oOcur In m_inventorApplication.ActiveDocument.SelectSet
            'Try
            Dim ADocType As Object : ADocType = oOcur.Type
            If ADocType = 100669440 Or ADocType = 100669696 Then
                If oOcur.ParentComponents.Item(1).Type = 100669440 Or oOcur.ParentComponents.Item(1).Type = 100669696 Then 'oOcur = oOcur.ParentComponents.Item(i)
                    For i = 1 To oOcur.ParentComponents.Item(1).ParentComponents.Count
                        OldDisplName = oOcur.ParentComponents.Item(1).ParentComponents.Item(i).Name
                        OldFFName = oOcur.ParentComponents.Item(1).ParentComponents.Item(i).ReferencedFileDescriptor.FullFileName
                        If DelStaroeVse = False Then If DelStaroe = True Then If MsgBox("Удалять скопированые файл " & OldDisplName & "?", vbYesNoCancel, "Вопрос") = vbYes Then Udalit = True
                        ZamenaImeni(PapkaCBOR, OldFFName, OldDisplName, FFName, Udalit)
                    Next i
                Else
                    For i = 1 To oOcur.ParentComponents.Count
                        OldDisplName = oOcur.ParentComponents.Item(i).Name
                        OldFFName = oOcur.ParentComponents.Item(i).ReferencedFileDescriptor.FullFileName
                        If DelStaroeVse = False Then If DelStaroe = True Then If MsgBox("Удалять скопированые файл " & OldDisplName & "?", vbYesNoCancel, "Вопрос") = vbYes Then Udalit = True
                        ZamenaImeni(PapkaCBOR, OldFFName, OldDisplName, FFName, Udalit)
                    Next i
                End If
            End If
            If ADocType = ObjectTypeEnum.kComponentOccurrenceObject Then
                OldDisplName = oOcur.Name
                OldFFName = oOcur.Definition.Document.FullFileName
                If DelStaroeVse = False Then If DelStaroe = True Then If MsgBox("Удалять скопированые файл " & OldDisplName & "?", vbYesNoCancel, "Вопрос") = vbYes Then Udalit = True
                ZamenaImeni(PapkaCBOR, OldFFName, OldDisplName, FFName, Udalit)
            End If
        Next
        If MsgBox("Открыть переименованные?", vbYesNo, "Вопрос") = vbYes Then
            'MsgBox(UBound(gl_OpenFiles), , "UBound(gl_OpenFiles)")
            For iii = 0 To UBound(gl_OpenFiles)
                'MsgBox(iii, , "iii")
                If gl_OpenFiles(iii) <> "" Then
                    If Right(gl_OpenFiles(iii), 3) = "idw" Then
                        m_inventorApplication.Documents.Open(gl_OpenFiles(iii), False).Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedFileDescriptor.ReplaceReference(gl_OpenFiles(iii - 1))
                    Else
                        m_inventorApplication.Documents.Open(gl_OpenFiles(iii), True)
                    End If

                End If
            Next iii
        End If
    End Sub

    Private Function Zamena(ByVal PapkaCBOR As String, ByVal OldFFName As String, ByVal OldDisplName As String, ByVal sIndex As String, Optional ByVal DelStaroe As Boolean = False) As String
        Dim PapkaDET As String = Left(OldFFName, OldFFName.LastIndexOf("\") + 1)
        Dim ImjaDET As String = Mid(OldFFName, OldFFName.LastIndexOf("\") + 2, Len(Mid(OldFFName, OldFFName.LastIndexOf("\") + 2))) ' с расширением
        Dim ZamenImjaDET As String = NowoeImjaDET(ImjaDET, sIndex)
        Dim NewFFName As String = PapkaCBOR & ZamenImjaDET
        Dim NewFFNameOld As String

        'If PapkaDET <> PapkaCBOR Then
        '  проверка существования файла
        If System.IO.File.Exists(NewFFName) Then
            Select Case MsgBox("Файл  " & ZamenImjaDET & "  существует. Использовать имеющийся?", vbYesNoCancel, "Вопрос")
                Case vbYes : m_inventorApplication.ActiveDocument.ComponentDefinition.Occurrences.ItemByName(OldDisplName).Replace(NewFFName, False)
                Case vbNo
                    Do
                        'Dim DopIndex As String
                        sIndex = InputBox("Введите новый разделитель и новый индекс", "Ввод данных", "-11")
                        If sIndex = "" Then MsgBox("Индекс не выбран. Компонент  " & OldDisplName & "  не заменен", 48, "Отмена") : Zamena = "" : Exit Function
                        NewFFNameOld = NewFFName
                        NewFFName = PapkaCBOR & NowoeImjaDET(ImjaDET, sIndex)
                    Loop While NewFFName = NewFFNameOld
                    Kopirovanie(OldFFName, NewFFName, OldDisplName)
                Case vbCancel : MsgBox("Компонент  " & OldDisplName & "  не заменен", 48, "Отмена") : Zamena = "" : Exit Function
            End Select
        Else
            MsgBox(NewFFName, , "NewFFName")
            Kopirovanie(OldFFName, NewFFName, OldDisplName)
        End If
        'End If
        Zamena = NewFFName
    End Function 'Zamena

    Private Function ZamenaImeni(ByVal PapkaCBOR As String, ByVal OldFFName As String, ByVal OldDisplName As String, ByVal FFNameSB As String, Optional ByVal DelStaroe As Boolean = False) As String
        Dim NewFFName As String ' = PapkaCBOR & ZamenImjaDET
        Dim PartNumberFFNameSB As String = Mid(Split(FFNameSB, "\")(UBound(Split(FFNameSB, "\"))), 1, InStr(InStr(Split(FFNameSB, "\")(UBound(Split(FFNameSB, "\"))), " ") + 1, Split(FFNameSB, "\")(UBound(Split(FFNameSB, "\"))), " ") - 1)
        Dim DescriptionOldFFName As String = Mid(Split(OldFFName, "\")(UBound(Split(OldFFName, "\"))), InStr(InStr(Split(OldFFName, "\")(UBound(Split(OldFFName, "\"))), " ") + 1, Split(OldFFName, "\")(UBound(Split(OldFFName, "\"))), " ") + 1) 'Split(OldFFName, "\")(UBound(Split(OldFFName, "\")))
        'FFNPartNumber = Mid(FFNnameBR, 1, InStr(InStr(FFNnameBR, " ") + 1, FFNnameBR, " ") - 1)
        'FFNDescription = Mid(FFNnameBR, InStr(InStr(FFNnameBR, " ") + 1, FFNnameBR, " ") + 1)
        Dim sSFD As New SaveFileDialog()
        Do
            sSFD.InitialDirectory = PapkaCBOR ' Указываем начальную папку
            sSFD.Title = "Новое имя файла" ' Указываем заголовок
            sSFD.FileName = PartNumberFFNameSB & " " & DescriptionOldFFName 'ImjaSB ' Mid(ImjaSB, 1, InStr(InStr(ImjaSB, " ") + 1, ImjaSB, " ") - 1) 
            sSFD.Filter = "Сборки|*.iam|Детали|*.ipt|Все файлы (*.*)|*.*"
            sSFD.FilterIndex = 1 : If Right(OldFFName, 4) = ".ipt" Then sSFD.FilterIndex = 2
            If sSFD.ShowDialog() = DialogResult.OK Then
                NewFFName = sSFD.FileName
                Exit Do
            End If
            If MsgBox("Не задано имя файла. Отказаться от копирования?", vbYesNoCancel, "Вопрос") = vbYes Then ZamenaImeni = "" : Exit Function
        Loop
        Kopirovanie(OldFFName, NewFFName, OldDisplName, DelStaroe)
        ZamenaImeni = NewFFName
    End Function

    Private Function NowoeImjaDET(ByVal ImjaDET As String, ByVal sIndex As String) As String
        Dim VtorProbel As String = InStr(InStr(ImjaDET, " ") + 1, ImjaDET, " ")
        Dim sObozna4 As String = Left(ImjaDET, VtorProbel - 1)
        Dim sNaimen As String = Mid(ImjaDET, VtorProbel, Len(ImjaDET) - VtorProbel + 1)

        If InStr(sObozna4, "-") <> 0 Then
            Dim DopIndex As String = Mid(sObozna4, InStr(sObozna4, "-"))
            sIndex = InputBox("Индекс уже имеется. Введите индекс для детали:" & Chr(13) & sObozna4, "Уточнение данных", CStr(DopIndex))
            sObozna4 = Left(sObozna4, InStr(sObozna4, "-") - 1)
        End If
        NowoeImjaDET = sObozna4 & sIndex & sNaimen
        MsgBox(sObozna4 & "   - sObozna4" & Chr(13) & sNaimen & "   - sNaimen" & Chr(13) & NowoeImjaDET & "   - NowoeImjaDET", , "NowoeImjaDET -- Проверка значений")
    End Function 'NowoeImjaDET

    Private Sub Kopirovanie(ByVal OldFFName As String, ByVal NewFFName As String, ByVal OldDisplName As String, Optional ByVal DelStaroe As Boolean = False)
        'Dim NewPapka As String = Left(NewFFName, NewFFName.LastIndexOf("\"))
        'Dim OldPapka As String = Left(OldFFName, Len(NewPapka))
        'If DelStaroe Then If NewPapka <> OldPapka Then DelStaroe = False
        Dim IndexOpenFiles As Integer
        'MsgBox("147", , "222")
        Try
            System.IO.File.Copy(OldFFName, NewFFName) ' копирование в системе
            m_inventorApplication.ActiveDocument.ComponentDefinition.Occurrences.ItemByName(OldDisplName).Replace(NewFFName, False)
            'MsgBox("258 " & UBound(gl_OpenFiles), , "222")
            If UBound(gl_OpenFiles) = 1 Then IndexOpenFiles = 1 Else IndexOpenFiles = UBound(gl_OpenFiles) + 1
            'MsgBox(IndexOpenFiles, , "_IndexOpenFiles")
            ReDim Preserve gl_OpenFiles(IndexOpenFiles)
            gl_OpenFiles(IndexOpenFiles) = NewFFName
            'MsgBox(gl_OpenFiles(IndexOpenFiles), , "gl_OpenFiles(IndexOpenFiles)")
        Catch ex As Exception
            MsgBox("Ошибка в операции копирования.", , "Ошибка")
        End Try

        '=======   чертеж
        Dim DET4ertOldFulName As String = Left(OldFFName, InStrRev(OldFFName, ".")) & "idw"
        Dim NowPutDET4ert As String = Left(NewFFName, InStrRev(NewFFName, ".")) & "idw"
        If System.IO.File.Exists(DET4ertOldFulName) Then
            System.IO.File.Copy(DET4ertOldFulName, NowPutDET4ert) ' копирование в системе
            ReDim Preserve gl_OpenFiles(IndexOpenFiles + 1)
            gl_OpenFiles(IndexOpenFiles + 1) = NowPutDET4ert
            Try
                If DelStaroe Then System.IO.File.Delete(DET4ertOldFulName)
            Catch ex As Exception
                MsgBox("Чертеж " & Split(DET4ertOldFulName, "\")(UBound(Split(DET4ertOldFulName, "\"))) & " не удален", , "Ошибка")
            End Try
            'MsgBox("перед заменой чертежа")
            m_inventorApplication.Documents.Open(NowPutDET4ert, False).Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedFileDescriptor.ReplaceReference(NewFFName)
            'MsgBox("после заменой чертежа")
        End If 'System.IO.File.Exists
        '=======   чертеж
        Try
            If DelStaroe Then System.IO.File.Delete(OldFFName)
        Catch ex As Exception
            MsgBox("Файл " & Split(OldFFName, "\")(UBound(Split(OldFFName, "\"))) & " не удален", , "Ошибка")
        End Try

        Dim SpecOldFulName As String = Left(OldFFName, InStrRev(OldFFName, ".")) & "xls"
        Dim NowPutSpec As String = Left(NewFFName, InStrRev(NewFFName, ".")) & "xls"
        If System.IO.File.Exists(SpecOldFulName) Then
            System.IO.File.Copy(SpecOldFulName, NowPutSpec) ' копирование в системе
            Try
                If DelStaroe Then System.IO.File.Delete(SpecOldFulName)
            Catch ex As Exception
                MsgBox("Спецификация " & Split(SpecOldFulName, "\")(UBound(Split(SpecOldFulName, "\"))) & " не скопирована", , "Ошибка")
            End Try
        End If 'System.IO.File.Exists
    End Sub 'Kopirovanie

    Sub AvtoNameSB(Optional ByVal OldNameSB As String = "")
        'Dim OldNameSB As String
        Dim IndexNa4 As Integer
        Dim IndexKonec As Integer
        IndexNa4 = InStr(OldNameSB, " ")
        IndexKonec = InStr(OldNameSB, ".")
        Do
            IndexNa4 = IndexKonec
            IndexKonec = InStr(IndexNa4 + 1, OldNameSB, ".")
            If IndexKonec = 0 Then IndexKonec = Len(OldNameSB)
        Loop Until Val(Mid(OldNameSB, IndexNa4 + 1, IndexKonec - IndexNa4)) = 0

        MsgBox(Strings.Left(OldNameSB, IndexNa4) & "  IndexNa4" & Chr(13) & CStr(Mid(OldNameSB, IndexNa4 + 1, IndexKonec - IndexNa4)) & "  CStr" & Chr(13) & Mid(OldNameSB, IndexKonec) & "  IndexKonec")
    End Sub
    Sub Open_for_name()
        ReDim gl_OpenFiles(0)
        Dim IndexOpenFiles As Integer
        Dim Iskat As String = InputBox("имя файла")
        'Dim FFName As String = m_inventorApplication.ActiveDocument.FullFileName
        'Dim testFile As System.IO.FileInfo = My.Computer.FileSystem.GetFileInfo(m_inventorApplication.ActiveDocument.FullFileName)
        Dim folderPath As String = My.Computer.FileSystem.GetFileInfo(m_inventorApplication.ActiveDocument.FullFileName).DirectoryName
        'Dim folderPath As String = testFile.DirectoryName
        'MsgBox(folderPath)
        'Dim fileName As String = testFile.Name
        'MsgBox(fileName)

        'Dim Folder As IO.Directory ' Объявляем переменную Folder для работы с папками
        Dim Files() As String ' Объявляем массив Files для хранения строк. Если ты заметил, то в скобках количество индексов не указано, это потому что мы не знаем сколько файлов будет хранится, а если в скобках ни чего не указывать, то количество индексов не ограниченно
        Dim i As Integer

        Files = IO.Directory.GetFiles(folderPath) ' В массив Files записываются все файлы содержащиеся в folderPath
        Dim LText As String

        For i = 0 To Files.Length - 1 ' Массив длится столько, сколько файлов в folderPath
            If InStr(LCase(Files(i)), LCase(Trim(Iskat))) <> 0 Then
                IndexOpenFiles = UBound(gl_OpenFiles) + 1
                'If UBound(gl_OpenFiles) = 1 Then IndexOpenFiles = 1 Else IndexOpenFiles = UBound(gl_OpenFiles) + 1
                'MsgBox(IndexOpenFiles & "  _IndexOpenFiles" & Chr(13) & UBound(gl_OpenFiles) & "UBound(gl_OpenFiles)", , "111")
                ReDim Preserve gl_OpenFiles(IndexOpenFiles)
                gl_OpenFiles(IndexOpenFiles) = Files(i)
                LText = LText & Files(i) & vbCrLf ' Извлекаем из массива Files имена всех файлов, и записываем их в метку. Функция vbCrLf нужна для перевода строки
            End If
        Next i
        MsgBox(LText, , "Найдены следующие файлы")

        If MsgBox("Открыть найденные?", vbYesNo, "Вопрос") = vbYes Then
            'MsgBox(UBound(gl_OpenFiles) & "UBound(gl_OpenFiles)" & vbCrLf & gl_OpenFiles.Length - 1 & "gl_OpenFiles.Length-1")
            For iii = 0 To UBound(gl_OpenFiles)
                'MsgBox(iii & " iii" & Chr(13) & gl_OpenFiles(iii) & " gl_OpenFiles(iii)")
                If gl_OpenFiles(iii) <> "" Then
                    'If Right(gl_OpenFiles(iii), 3) = "idw" Then
                    '    m_inventorApplication.Documents.Open(gl_OpenFiles(iii), False).Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedFileDescriptor.ReplaceReference(gl_OpenFiles(iii - 1))
                    'Else
                    Try
                        m_inventorApplication.Documents.Open(gl_OpenFiles(iii), True)
                    Catch ex As Exception
                        MsgBox("Не удалось открыть документ:" & vbCrLf & gl_OpenFiles(iii), , "Ошибка")
                    End Try
                    'End If
                End If
            Next iii
        End If

    End Sub

End Module
