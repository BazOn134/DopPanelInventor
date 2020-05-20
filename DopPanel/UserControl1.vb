Imports System.Windows.Forms
Imports Inventor

Public Class UserControl1
    'Подготавливаем массив строк для хранения информации из файла:======================
    Dim Dannje(10) As String 'Размер берем с запасом

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Try
            If m_inventorApplication.Documents.VisibleDocuments.Count = 0 Then Exit Try
            If m_inventorApplication.ActiveEditDocument.FullFileName = "" Then m_inventorApplication.ActiveEditDocument.Save()
            m_Form.Obrabotka(m_inventorApplication.ActiveEditDocument.FullFileName, True, True)
        Catch ex As Exception
            MsgBox("Операция не выполнена.", 48, "Информация")
        End Try
    End Sub

    Private Sub ОбработкаToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        m_Form.Obrabotka(m_inventorApplication.ActiveEditDocument.FullFileName, True)
    End Sub


    Private Sub ReloadFNastrToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReloadFNastrToolStripMenuItem.Click
        StreamFile()
    End Sub

    Private Sub ИзменитьПутиКШаблонамЧертежейToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ИзменитьПутиКШаблонамЧертежейToolStripMenuItem.Click
        ИзменитьПутиКШаблонамЧертежей()
    End Sub

    Private Sub ОткрытьЧертежиToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОткрытьЧертежиToolStripMenuItem1.Click, ОткрытьЧертежиВыделенныхToolStripMenuItem.Click
        ОткрытьЧертежи()
    End Sub

    Private Sub СоздатьPDFToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles СоздатьPDFToolStripMenuItem.Click
        'If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
        SavePDF()
        'Else
        'MsgBox("Это не чертеж", 48, "Внимание")
        'End If
    End Sub

    Private Sub ОткрытьPDFИXLSToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОткрытьPDFИXLSToolStripMenuItem.Click
        OpenPDFXLS(True, True)
    End Sub

    Private Sub ОткрытьСпецификациюToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОткрытьСпецификациюToolStripMenuItem.Click
        OpenPDFXLS(False, True)
    End Sub

    Private Sub ОткрытьPDFToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОткрытьPDFToolStripMenuItem.Click
        OpenPDFXLS(True, False)
    End Sub

    Private Sub ОткрытьВПроводникеToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОткрытьВПроводникеToolStripMenuItem.Click
        Vydelennoe(True, False, False)
    End Sub

    Private Sub СкопироватьВБуферПутьToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles СкопироватьВБуферПутьToolStripMenuItem.Click
        Vydelennoe(False, False, True)
    End Sub

    Private Sub СкопироватьВБуферИмяToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles СкопироватьВБуферИмяToolStripMenuItem.Click
        Vydelennoe(False, True, False)
    End Sub

    Private Sub КопированиеСУдалениемToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles КопированиеСУдалениемToolStripMenuItem.Click
        NameCopys(True)
    End Sub

    Private Sub КопированиеToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles КопированиеToolStripMenuItem1.Click
        NameCopys()
    End Sub

    Private Sub КопированиеКаждойToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles КопированиеКаждойToolStripMenuItem.Click
        NameCopys(False, True)
    End Sub

    Private Sub ПроверкаСвойствОдноуровневыхЭлементовToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ПроверкаСвойствОдноуровневыхЭлементовToolStripMenuItem.Click
        ПроверкаСвойств()
    End Sub

    Private Sub ПроверкаСвойствВсехУровнейЭлементовToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ПроверкаСвойствВсехУровнейЭлементовToolStripMenuItem.Click
        ПроверкаСвойств(True)
    End Sub

    Private Sub ДобавитьПараметрыДхШхВToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ДобавитьПараметрыДхШхВToolStripMenuItem.Click
        Dim pDoc As Document = m_inventorApplication.ActiveDocument
        Dim oUserPars As UserParameters = pDoc.ComponentDefinition.Parameters.UserParameters
        Dim oPar As UserParameter = Nothing
        Try
            oPar = oUserPars.Item("Длина")
        Catch ex As Exception
            oPar = oUserPars.AddByExpression("Длина", 250, "mm")
            oPar.ExposedAsProperty = True
            oPar.IsKey = True
            oPar.Comment = "Добавлено"
        End Try
        oPar = Nothing
        Try
            oPar = oUserPars.Item("Ширина")
        Catch ex As Exception
            oPar = oUserPars.AddByExpression("Ширина", 100, "mm")
            oPar.ExposedAsProperty = True
            oPar.IsKey = True
            oPar.Comment = "Добавлено"
        End Try
        oPar = Nothing
        Try
            oPar = oUserPars.Item("Высота")
        Catch ex As Exception
            oPar = oUserPars.AddByExpression("Высота", 50, "mm")
            oPar.ExposedAsProperty = True
            oPar.IsKey = True
            oPar.Comment = "Добавлено"
        End Try
        oPar = Nothing
        m_inventorApplication.CommandManager.ControlDefinitions("AppParametersCmd").Execute()
    End Sub

    Private Sub ДобавитьПараметрыDxLToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ДобавитьПараметрыDxLToolStripMenuItem.Click
        Dim pDoc As Document = m_inventorApplication.ActiveDocument
        Dim oUserPars As UserParameters = pDoc.ComponentDefinition.Parameters.UserParameters
        Dim oPar As UserParameter = Nothing
        Try
            oPar = oUserPars.Item("DD")
        Catch ex As Exception
            oPar = oUserPars.AddByExpression("DD", 100, "mm")
            oPar.ExposedAsProperty = True
            oPar.IsKey = True
            oPar.Comment = "Диаметр"
        End Try
        oPar = Nothing
        Try
            oPar = oUserPars.Item("LL")
        Catch ex As Exception
            oPar = oUserPars.AddByExpression("LL", 250, "mm")
            oPar.ExposedAsProperty = True
            oPar.IsKey = True
            oPar.Comment = "Длина"
        End Try
        oPar = Nothing
        m_inventorApplication.CommandManager.ControlDefinitions("AppParametersCmd").Execute()
    End Sub

    Private Sub ДобавитьВПРОЧИЕИмяФайлаToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ДобавитьВПРОЧИЕИмяФайлаToolStripMenuItem.Click
        'Dim pDoc As PartDocument = m_inventorApplication.ActiveEditDocument
        'Dim oPropSet As PropertySet = m_inventorApplication.ActiveDocument.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}")
        Try
            m_inventorApplication.ActiveEditDocument.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}").Item("Имя файла").Value = "Не определено"
        Catch ex As Exception
            Dim oProp As [Property] = m_inventorApplication.ActiveDocument.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}").Add("Не определено", "Имя файла")
        End Try
    End Sub

    Private Sub УдалитьПрочиеСвойстваToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles УдалитьПрочиеСвойстваToolStripMenuItem1.Click
        УдалитьПрочиеСвойства()
    End Sub

    Private Sub УдалитьПрочиеСвойствасВыборомToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles УдалитьПрочиеСвойствасВыборомToolStripMenuItem.Click
        УдалитьПрочиеСвойства(True)
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        'FormSvojstva================================================================================================================================================
        Dim sv_Form As Form = New FormSvojstva
        sv_Form.Visible = True
        'sv_Form.Activate = True
    End Sub

    Private Sub ЗаменитьМодельToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ЗаменитьМодельToolStripMenuItem.Click
        Dim FNameOld As String = m_inventorApplication.ActiveDocument.Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedFileDescriptor.FullFileName
        Dim Ras6irenie As String = Strings.Right(FNameOld, 4)
        Dim Zamena As String = Strings.Left(m_inventorApplication.ActiveDocument.FullFileName, Len(m_inventorApplication.ActiveDocument.FullFileName) - 4) & Ras6irenie
        'MsgBox(Zamena, , "Zamena")
        If Zamena <> FNameOld Then m_inventorApplication.ActiveDocument.Sheets.Item(1).DrawingViews.Item(1).ReferencedDocumentDescriptor.ReferencedFileDescriptor.ReplaceReference(Zamena)
    End Sub

    Public Function Obrabotka(ByVal FFName As String, Optional ByVal Zamena As Boolean = False, Optional ByVal Prinudit As Boolean = False) As Boolean
        'MsgBox(m_inventorApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath & "_    WorkspacePath")
        'Strings.Left(FFName, Len(m_inventorApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath))
        m_Form.ToolStrip1.BackColor = Drawing.Color.White
        ReadFile()
        Try
            If m_inventorApplication.Documents.VisibleDocuments.Count > 0 Then
                If Strings.Left(FFName, Len(m_inventorApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath)) <> m_inventorApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath Then m_Form.ToolStrip1.BackColor = Drawing.Color.Aqua : Exit Function
                If Prinudit = True Then Zamena = True
                Dim Nesovpadenie As Boolean
                ' читаем свойства файла
                Dim ActiveEditDoc As PropertySets = m_inventorApplication.ActiveEditDocument.PropertySets
                Dim detAvtor As String = ActiveEditDoc.Item("Inventor Summary Information").Item("Author").Value
                Dim detRazrabotal As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Designer").Value
                Dim detProveril As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Checked By").Value
                Dim detNormokontrol As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Engr Approved By").Value
                Dim detUtverdil As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Mfg Approved By").Value
                Dim detTitle As String = ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value
                Dim detDescription As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value
                Dim detPartNumber As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value
                Dim detMaterial As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Material").Value

                ' обработка количества пробелов
                Dim FFNnameBR As String = Mid(FFName, InStrRev(FFName, "\") + 1, Len(FFName) - InStrRev(FFName, "\") - 4)
                Dim FFNPartNumber As String = ""
                Dim FFNDescription As String = ""
                Dim Pro4ie As Boolean = False
                Dim CustExist As Boolean = False
                Dim CustomName As String = "Имя файла"
                Dim CustomValue As String = ""
                Dim PodmenaDannje1 As Integer
                'Dim FFNDisplayName As String = m_inventorApplication.ActiveEditDocument.DisplayName

                ProverkaPro4ih(m_inventorApplication.ActiveEditDocument, CustomName, CustomValue, CustExist)

                Dim Koli4estvoProbelov As Integer = 0
                For Each Simvol In FFNnameBR
                    If Simvol = " " Then Koli4estvoProbelov = Koli4estvoProbelov + 1
                Next

                If Not FFNnameBR = detPartNumber & " " & detDescription Then 'detPartNumber + detDescription
                    If Dannje(1) = "-" Then PodmenaDannje1 = 0 Else PodmenaDannje1 = Dannje(1)
                    If Koli4estvoProbelov > PodmenaDannje1 Then
                        Select Case Dannje(1)
                            Case "-"
                                FFNPartNumber = FFNnameBR
                                FFNDescription = ""
                            Case "0"
                                FFNPartNumber = Mid(FFNnameBR, 1, InStr(FFNnameBR, " ") - 1)
                                FFNDescription = Mid(FFNnameBR, InStr(FFNnameBR, " ") + 1)
                            Case "1"
                                FFNPartNumber = Mid(FFNnameBR, 1, InStr(InStr(FFNnameBR, " ") + 1, FFNnameBR, " ") - 1)
                                FFNDescription = Mid(FFNnameBR, InStr(InStr(FFNnameBR, " ") + 1, FFNnameBR, " ") + 1)
                            Case Else
                                Pro4ie = True
                        End Select
                    Else
                        If Zamena = True Then
                            Dim msg As String = "Невозможно автоматически разделить имя файла." & Chr(13) & Chr(13) & Chr(34) & FFNnameBR & Chr(34) & Chr(13) & Chr(13) & "Выберите варианты разделения:" & Chr(13) & """Да"" - Имя файла в ОБОЗНАЧЕНИЕ;" & Chr(13) & """Нет"" - Имя файла в НАИМЕНОВАНИЕ;" & Chr(13) & """Отмена"" - оставляем свойства как есть."
                            Dim title As String = "Выбор пользователя"
                            Dim style As MsgBoxStyle = vbYesNoCancel
                            Try
                                m_inventorApplication.ActiveEditDocument.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}").Item("Имя файла").Value = "Не определено"
                            Catch ex As Exception
                                Dim oProp As [Property] = m_inventorApplication.ActiveDocument.PropertySets.Item("{D5CDD505-2E9C-101B-9397-08002B2CF9AE}").Add("Не определено", "Имя файла")
                            End Try

                            Select Case MsgBox(msg, style, title)
                                Case vbYes
                                    FFNPartNumber = FFNnameBR
                                    FFNDescription = ""
                                    ActiveEditDoc.Item("Inventor User Defined Properties").Item("Имя файла").Value = "В обозначение"
                                Case vbNo
                                    FFNPartNumber = ""
                                    FFNDescription = FFNnameBR
                                    ActiveEditDoc.Item("Inventor User Defined Properties").Item("Имя файла").Value = "В наименование"
                                Case vbCancel
                                    m_Form.ToolStrip1.BackColor = Drawing.Color.Red
                            End Select
                        Else
                            Nesovpadenie = True
                        End If
                    End If
                    'Else
                    'm_Form.ToolStrip1.BackColor = Drawing.Color.
                    '----------------------------------------------- проверка прочих -------------------------------------
                    If CustExist Then
                        Select Case CustomValue
                            Case "В обозначение"
                                ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value = FFNnameBR
                                ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value = ""
                                ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value = ""
                            Case "В наименование"
                                ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value = ""
                                ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value = FFNnameBR
                                ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value = FFNnameBR
                            Case Else
                                Dim msg As String = "Невозможно автоматически разделить имя файла." & Chr(13) & Chr(13) & Chr(34) & FFNnameBR & Chr(34) & Chr(13) & Chr(13) & "Выберите варианты разделения:" & Chr(13) & """Да"" - Имя файла в ОБОЗНАЧЕНИЕ;" & Chr(13) & """Нет"" - Имя файла в НАИМЕНОВАНИЕ;" & Chr(13) & """Отмена"" - оставляем свойства как есть."
                                Dim title As String = "Выбор пользователя"
                                Dim style As MsgBoxStyle = vbYesNoCancel
                                Select Case MsgBox(msg, style, title)
                                    Case vbYes
                                        ActiveEditDoc.Item("Inventor User Defined Properties").Item("Имя файла").Value = "В обозначение"
                                    Case vbNo
                                        ActiveEditDoc.Item("Inventor User Defined Properties").Item("Имя файла").Value = "В наименование"
                                    Case vbCancel
                                        m_Form.ToolStrip1.BackColor = Drawing.Color.Red
                                End Select
                        End Select
                        Exit Function
                        '----------------------------------------------- проверка прочих -------------------------------------
                    Else 'CustExist
                        If FFNPartNumber <> detPartNumber Or FFNDescription <> detDescription Or FFNDescription <> detTitle Then Nesovpadenie = True
                        If Nesovpadenie = True Or Prinudit = True Then
                            If Zamena = False Then
                                m_Form.ToolStrip1.BackColor = Drawing.Color.Red
                                'ReadFile()
                                'MessageBox.Show(Dannje(1) & "_   Dannje(1) -- Zamena False =  Obrabotka")
                                'MessageBox.Show(Dannje(2) & "_   Dannje(2)" & Chr(13) & Dannje(3) & "_   Dannje(3)" & Chr(13) & Dannje(4) & "_   Dannje(4)" & Chr(13) & Dannje(5) & "_   Dannje(5)" & Chr(13) & Dannje(6) & "_   Dannje(6)")
                            Else
                                m_inventorApplication.ActiveEditDocument.DisplayName = FFNnameBR
                                If Dannje(2) <> "" Then ActiveEditDoc.Item("Inventor Summary Information").Item("Author").Value = Dannje(2)
                                If Dannje(3) <> "" Then ActiveEditDoc.Item("Design Tracking Properties").Item("Designer").Value = Dannje(3)
                                If Dannje(4) <> "" Then ActiveEditDoc.Item("Design Tracking Properties").Item("Checked By").Value = Dannje(4)
                                If Dannje(5) <> "" Then ActiveEditDoc.Item("Design Tracking Properties").Item("Engr Approved By").Value = Dannje(5)
                                If Dannje(6) <> "" Then ActiveEditDoc.Item("Design Tracking Properties").Item("Mfg Approved By").Value = Dannje(6)
                                '------------- обработка прочих
                                If CustExist Then
                                    MessageBox.Show(CustomValue & "_ CustomValue" & Chr(13) & FFNPartNumber & "_ FFNPartNumber" & Chr(13) & FFNDescription & "_ FFNDescription", "CustExist")
                                    Select Case CustomValue
                                        Case "В обозначение"
                                            ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value = FFNnameBR
                                            ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value = ""
                                            ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value = ""
                                        Case "В наименование"
                                            ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value = ""
                                            ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value = FFNnameBR
                                            ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value = FFNnameBR
                                    End Select
                                Else
                                    ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value = FFNDescription
                                    ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value = FFNDescription
                                    ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value = FFNPartNumber
                                End If
                                m_Form.ToolStrip1.BackColor = Drawing.Color.White
                                Obrabotka = True
                                'MessageBox.Show("Zamena true =  конец")
                            End If 'Zamena = False
                        Else
                            If m_Form.ToolStrip1.BackColor = Drawing.Color.Red Then m_Form.ToolStrip1.BackColor = Drawing.Color.White
                        End If 'Nesovpadenie
                    End If 'CustExist
                End If 'detPartNumber + detDescription
            Else 'm_inventorApplication.Documents.VisibleDocuments.Count > 0
                m_Form.ToolStrip1.BackColor = Drawing.Color.White
                MessageBox.Show("Нет открытых документов", "Информация")
            End If 'm_inventorApplication.Documents.VisibleDocuments.Count > 0
            'm_inventorApplication.ActiveEditDocument.Rebuild()
        Catch ex As Exception
            m_Form.ToolStrip1.BackColor = Drawing.Color.White
            MsgBox("Операция опять не выполнена.", 36, "Информация")
        End Try
        Obrabotka = True
    End Function 'Obrabotka

    Public Sub SavePDF()
        If m_inventorApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kDrawingDocumentObject Then ОткрытьЧертежи() : Exit Sub
        ''MsgBox("SavesPDF " & SavesPDF & OpenPDF & OpenXLS, 48, "SavesPDF") '======================================================
        Dim oDoc As Document = m_inventorApplication.ActiveDocument
        Dim sFn As String = oDoc.FullFileName
        Dim sPDF As String = Replace(sFn, ".idw", ".pdf")
        Dim sXLS As String = Replace(sFn, ".idw", ".xls")

        Dim oPDFAddIn As TranslatorAddIn = m_inventorApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
        Dim oDocument As Document = m_inventorApplication.ActiveDocument
        Dim oContext = m_inventorApplication.TransientObjects.CreateTranslationContext
        oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        Dim oOptions = m_inventorApplication.TransientObjects.CreateNameValueMap
        Dim oDataMedium = m_inventorApplication.TransientObjects.CreateDataMedium

        If oPDFAddIn.HasSaveCopyAsOptions(oDataMedium, oContext, oOptions) Then
            oOptions.Value("All_Color_AS_Black") = 1
            'oOptions.Value("Remove_Line_Weights") = 1
            oOptions.Value("Vector_Resolution") = 600
            oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
            'oOptions.Value("Custom_Begin_Sheet") = 2
            'oOptions.Value("Custom_End_Sheet") = 4
        End If
        'Dim oFolder As String = Strings.Left(oPath, InStrRev(oPath, "\")) & "PDF" 'создаем путь к папке PDF
        'If Not System.IO.Directory.Exists(oFolder) Then  System.IO.Directory.CreateDirectory(oFolder) 'проверяем наличие папки PDF, создаем пари отсутствии

        oDataMedium.FileName = sPDF 'oFolder & "\" & oFileName & " Rev" & oRevNum & ".pdf"
        oPDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium) 'сохраняем document

        OpenPDFXLS(True, True)
        oDoc.Save2()
        oDoc.Close(True)
        '====================================================================================================================================================
        ' часть кода AddIn с сайта
        'If oDoc.DocumentType = 12292 Then
        '    Dim sFn As String = oDoc.FullFileName

        '    sFn = Replace(sFn, ".idw", ".pdf")
        '    sFn = Replace(sFn, ".IDW", ".PDF")

        '    ' Get the PDF translator Add-In.  
        '    Dim oPDFTrans As TranslatorAddIn = InvApp.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}")
        '    If oPDFTrans Is Nothing Then
        '        MsgBox("Не удалось создать PDF файл.")
        '        Exit Sub
        '    End If

        '    ' Create some objects that are used to pass information to  the translator Add-In.   
        '    Dim oContext As TranslationContext = InvApp.TransientObjects.CreateTranslationContext
        '    oContext.Type = IOMechanismEnum.kFileBrowseIOMechanism
        '    Dim oOptions As NameValueMap = InvApp.TransientObjects.CreateNameValueMap

        '    If oPDFTrans.HasSaveCopyAsOptions(oDoc, oContext, oOptions) Then
        '            oOptions.Value("Sheet_Range") = Inventor.PrintRangeEnum.kPrintAllSheets
        '            oOptions.Value("All_Color_AS_Black") = True
        '            oOptions.Value("Vector_Resolution") = 400
        '            oOptions.Value("Remove_Line_Weights") = 0
        '        ' Define various settings and input to provide the translator.  
        '        Dim oData As DataMedium
        '        oData = InvApp.TransientObjects.CreateDataMedium
        '        oData.FileName = sFn
        '        ' Call the translator.  
        '        Call oPDFTrans.SaveCopyAs(InvApp.ActiveDocument, oContext, oOptions, oData)
        '    End If

        'End If
        '====================================================================================================================================================
    End Sub

    Public Sub OpenPDFXLS(Optional ByVal OpenPDF As Boolean = False, Optional ByVal OpenXLS As Boolean = False)
        Dim oDoc As Document = m_inventorApplication.ActiveEditDocument
        Dim sFn As String = oDoc.FullFileName
        Dim sPDF As String = Strings.Left(sFn, Len(sFn) - 4) & ".pdf" 'Replace(sFn, ".idw", ".pdf")
        Dim sXLS As String = Strings.Left(sFn, Len(sFn) - 4) & ".xls" 'Replace(sFn, ".idw", ".xls")
        If OpenPDF = True Then If System.IO.File.Exists(sPDF) Then Process.Start(sPDF) Else MessageBox.Show("Файл PDF не найден", "Информация")
        ' проверка на сборку
        If OpenXLS = True Then
            OpenXLS = False
            If oDoc.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then OpenXLS = True
            If oDoc.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then If oDoc.Sheets(1).DrawingViews(1).ReferencedDocumentDescriptor.ReferencedDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then OpenXLS = True
        End If

        If OpenXLS = True Then
            If System.IO.File.Exists(sXLS) Then
                Process.Start(sXLS)
            Else
                MsgBox("Спецификация не найдена." & Chr(13) & "Имя скопировано в буфер.", 48, "Информация")
                Clipboard.Clear()
                Clipboard.SetText(Strings.Left(Split(sFn, "\")(UBound(Split(sFn, "\"))), Len(Split(sFn, "\")(UBound(Split(sFn, "\")))) - 4), TextDataFormat.Text) 'Split(sFn, "\")(UBound(Split(sFn, "\")))
            End If
        End If
    End Sub

    Public Function Vydelennoe(Optional ByVal Otkrytje As Boolean = True, Optional ByVal KopirImja As Boolean = False, Optional ByVal KopirPut As Boolean = False) As Boolean
        Dim FFName As String = ""
        Dim ClipboardName As String = ""
        If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Or m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            FFName = m_inventorApplication.ActiveDocument.FullFileName
            If Otkrytje = True Then Shell("explorer.exe /select, " & FFName, vbNormalFocus) '& FFName
            If KopirImja = True Then Clipboard.SetText(Split(FFName, "\")(UBound(Split(FFName, "\"))), TextDataFormat.Text) ' MsgBox(Split(FFName, "\")(UBound(Split(FFName, "\"))))
            If KopirPut = True Then Clipboard.SetText(FFName, TextDataFormat.Text)
        End If
        If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            If m_inventorApplication.ActiveDocument.SelectSet.Count = 0 Then 'MsgBox("Сначала выделите компоненты", MessageBoxIcon.Information, "Внимание") : Exit Sub
                FFName = m_inventorApplication.ActiveDocument.FullFileName
                If Otkrytje = True Then Shell("explorer.exe /select, " & FFName, vbNormalFocus) '& FFName
                If KopirImja = True Then Clipboard.SetText(Split(FFName, "\")(UBound(Split(FFName, "\"))), TextDataFormat.Text) ' MsgBox(Split(FFName, "\")(UBound(Split(FFName, "\"))))
                If KopirPut = True Then Clipboard.SetText(FFName, TextDataFormat.Text)
            Else
                'Dim FFName As String
                'If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                'Dim oOcur As ComponentOccurrences
                Dim oAssyDoc As AssemblyDocument = m_inventorApplication.ActiveDocument
                'If m_inventorApplication.ActiveDocument.SelectSet.Count = 0 Then MsgBox("Сначала выделите компоненты", MessageBoxIcon.Information, "Внимание") : Exit Sub
                Dim oSelSet As SelectSet = m_inventorApplication.ActiveDocument.SelectSet
                Dim oDoc As Document
                Dim oOcur As ComponentOccurrence
                For Each oOcur In oSelSet
                    oDoc = oOcur.Definition.Document
                    FFName = oDoc.FullFileName
                    If Otkrytje = True Then
                        Dim Focus As AppWinStyle
                        If oSelSet.Count = 1 Then Focus = vbNormalFocus Else Focus = vbNormalNoFocus
                        Shell("explorer.exe /select, " & FFName, vbNormalNoFocus) '& FFName
                    End If
                    If KopirPut = True Then
                        If ClipboardName = "" Then
                            ClipboardName = FFName
                        Else
                            ClipboardName = ClipboardName & Chr(13) & Chr(10) & FFName
                        End If
                        FFName = ClipboardName
                    End If
                    If KopirImja = True Then
                        If ClipboardName = "" Then
                            ClipboardName = Split(FFName, "\")(UBound(Split(FFName, "\")))
                        Else
                            ClipboardName = ClipboardName & Chr(13) & Chr(10) & Split(FFName, "\")(UBound(Split(FFName, "\")))
                        End If
                        FFName = ClipboardName
                        Clipboard.SetText(FFName, TextDataFormat.Text) ' MsgBox(Split(FFName, "\")(UBound(Split(FFName, "\"))))
                    End If
                Next  ' oOcur
                'Clipboard.SetText(FFName, TextDataFormat.Text)
                'End If
            End If
        Else

        End If
        Vydelennoe = True
    End Function

    Sub StreamFile()
        My.Settings.Reload()
        My.Settings.Koli4estvoProbelov = InputBox("Количество пробелов", "Ввод данных", My.Settings.Koli4estvoProbelov)
        My.Settings.AutorFamilie = InputBox("Фамилия автора", "Ввод данных", My.Settings.AutorFamilie)
        My.Settings.Razrabot4ikFamilie = InputBox("Фамилия разработчика", "Ввод данных", My.Settings.Razrabot4ikFamilie)
        My.Settings.ProverilFamilie = InputBox("Фамилия проверяющего", "Ввод данных", My.Settings.ProverilFamilie)
        My.Settings.NormokontrolFamilie = InputBox("Фамилия нормоконтролера", "Ввод данных", My.Settings.NormokontrolFamilie)
        My.Settings.UtverdilFamilie = InputBox("Фамилия утверждающего", "Ввод данных", My.Settings.UtverdilFamilie)
        My.Settings.Save()


        ''Создаем объект для записи информации в текстовый файл E:\VB\Filimon.txt:
        'Dim Запись As New System.IO.StreamWriter("C:\ProgramData\Autodesk\Inventor 2014\Addins\option.ini")
        ''Записываем в файл строки:
        'Запись.WriteLine("Файл настроек AddIn ""DopPanel"" for Autodesk Inventor")
        'Запись.WriteLine("Для корректной работы файлы .dll и .addin должны распологаться по адресу C:\ProgramData\Autodesk\Inventor 2014\Addins")
        'Запись.WriteLine("Для исключения любого параметра необходимо перед знаком""<"" вместо параметра поставить знак ""-""")
        'Запись.WriteLine("1 <== Количество пробелов в обозначении. Например ""ПХИ 01.01.13 Труба"" пробел один: ""ПХИ 01.01.13"" - обозначение, ""Труба"" - наименование")
        'Запись.WriteLine("Автор <== Фамилия автора")
        'Запись.WriteLine("Разработчик <== Фамилия разработчика")
        'Запись.WriteLine("Проверяющий <== Фамилия проверяющего")
        'Запись.WriteLine("Нормоконтролер <== Фамилия нормоконтролера")
        'Запись.Write("Утверждающий <== Фамилия утверждающего")
        'Запись.Close() 'Закрываем файл

        'If MsgBox("Открыть файл для настройки?", vbYesNo, "Ожидание информации") = vbYes Then Process.Start("C:\ProgramData\Autodesk\Inventor 2014\Addins\option.ini")
    End Sub

    Public Sub ReadFile()
        My.Settings.Reload()
        Dannje(1) = My.Settings.Koli4estvoProbelov
        Dannje(2) = My.Settings.AutorFamilie
        Dannje(3) = My.Settings.Razrabot4ikFamilie
        Dannje(4) = My.Settings.ProverilFamilie
        Dannje(5) = My.Settings.NormokontrolFamilie
        Dannje(6) = My.Settings.UtverdilFamilie

        If My.Settings.Put4ertegDet = "" Or My.Settings.Put4ertegSB = "" Then ИзменитьПутиКШаблонамЧертежей()
        Dannje(9) = My.Settings.Put4ertegDet
        Dannje(10) = My.Settings.Put4ertegSB

        'Try
        '    'Подготавливаем массив строк для хранения информации из файла:======================
        '    If Not System.IO.File.Exists("C:\ProgramData\Autodesk\Inventor 2014\Addins\option.ini") Then StreamFile()
        '    Dim Строки(10) As String 'Размер берем с запасом
        '    'СЧИТЫВАЕМ ИНФОРМАЦИЮ ИЗ ФАЙЛА:
        '    Dim Чтение As New System.IO.StreamReader("C:\ProgramData\Autodesk\Inventor 2014\Addins\option.ini")
        '    Dim i As Integer = 1 'Счетчик строк файла
        '    'Считываем все строки файла:
        '    Do While Чтение.Peek() <> -1
        '        Строки(i) = Чтение.ReadLine
        '        i = i + 1
        '    Loop
        '    Чтение.Close() 'Закрываем файл
        '    'ОБРАБАТЫВАЕМ ИНФОРМАЦИЮ В ОПЕРАТИВНОЙ ПАМЯТИ:
        '    Dim Число_строк_в_файле As Integer = i - 1
        '    'Dim sZna4enie As String
        '    For i = 1 To Число_строк_в_файле
        '        If InStr(Строки(i), "<==") <> 0 Then
        '            Dim Rezult As String = Trim(Split(Строки(i), "<")(0))
        '            'MsgBox(Строки(i) & "  - Строки(i)" & Chr(13) & "_" & Rezult & "_   Rezult" & Chr(13) & "_" & Trim(Rezult) & "_   Trim Rezult" & Chr(13) & "_" & Asc("-") & "_           Asc -")
        '            If Trim(Rezult) <> "-" Then ' Or Rezult <> "" 
        '                If Rezult <> "" Then
        '                    If InStr(Строки(i), "пробел") <> 0 Then Dannje(1) = Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)) 'MsgBox("искать " & Val(Строки(i)) + 1 & " пробела")
        '                    If InStr(Строки(i), "автор") <> 0 Then Dannje(2) = Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)) 'MsgBox("автор " & Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)))
        '                    If InStr(Строки(i), "разраб") <> 0 Then Dannje(3) = Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)) 'MsgBox("разработал " & Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)))
        '                    If InStr(Строки(i), "провер") <> 0 Then Dannje(4) = Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)) 'MsgBox("проверил " & Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)))
        '                    If InStr(Строки(i), "нормо") <> 0 Then Dannje(5) = Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)) 'MsgBox("нормоконтролер " & Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)))
        '                    If InStr(Строки(i), "утвер") <> 0 Then Dannje(6) = Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)) 'MsgBox("утвердил " & Trim(Mid(Строки(i), 1, InStr(Строки(i), "<") - 1)))
        '                Else
        '                    'MsgBox("пропустим пусто")
        '                End If
        '            Else
        '                MsgBox("пропустим ---")
        '            End If
        '        End If
        '    Next i
        'Catch ex As Exception
        '    MsgBox("Ошибка")
        'End Try
    End Sub 'ReadFile

    Private Function УдалитьПрочиеСвойства(Optional ByVal Wopros As Boolean = False) As Boolean
        Dim AssDoc As Document
        If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Or m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kDrawingDocumentObject Then
            AssDoc = m_inventorApplication.ActiveDocument
        ElseIf m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
            If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Or m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                AssDoc = m_inventorApplication.ActiveEditDocument
            ElseIf m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Or m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    AssDoc = m_inventorApplication.ActiveEditDocument
                ElseIf m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    If m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject Or m_inventorApplication.ActiveDocument.DocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        AssDoc = m_inventorApplication.ActiveEditDocument
                    End If
                End If
            End If
        End If
        'Dim AssDoc As AssemblyDocument = m_inventorApplication.ActiveEditDocument
        'Dim oOccur As ComponentOccurrences = AssDoc.ComponentDefinition.Occurrences
        'Dim oOcc As ComponentOccurrence
        Dim oCustomPropSet As PropertySet
        oCustomPropSet = AssDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim oCustProp As [Property]
        For Each oCustProp In oCustomPropSet
            If Wopros = True Then
                Select Case MsgBox("Удалить свойство -  " & oCustProp.Name & " = " & oCustProp.Value & Chr(13) & "Отмена - выход из программы", vbYesNoCancel, "Вопрос")
                    Case vbYes : oCustProp.Delete()
                    Case vbNo
                    Case vbCancel : Exit Function
                End Select
            Else
                If Not oCustProp.Name = "Имя файла" Then oCustProp.Delete()
            End If
        Next 'oCustProp
    End Function

    Private Function ОткрытьЧертежи() As Boolean
        ReadFile()
        Dim sDoc As Document
        Dim Otkr As String
        Dim txtt, Templatka As String
        If m_inventorApplication.ActiveDocument.SelectSet.Count <> 0 Then '===================================================================================
            'Dim oOcur As SelectSet = m_inventorApplication.ActiveDocument.SelectSet
            For Each oOcur In m_inventorApplication.ActiveDocument.SelectSet
                sDoc = oOcur.Definition.Document
                Otkr = Strings.Left(sDoc.FullFileName, Len(sDoc.FullFileName) - 4) + ".idw"  ' путь к чертежу
                Try
                    Dim oDoc As Inventor.Document = m_inventorApplication.Documents.Open(Otkr, True) 'True False    ' открываю чертеж
                Catch

                    If oOcur.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        txtt = "Деталь"
                        Templatka = Dannje(9) '"E:\§ Inventor\Темплатки\!14 Дет.idw"
                    ElseIf oOcur.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        txtt = "Сборка"
                        Templatka = Dannje(10) '"E:\§ Inventor\Темплатки\!14 СБ.idw"
                    End If

                    Dim wopr As Object = MsgBox(txtt & Chr(13) & oOcur.Name & Chr(13) & Chr(13) & "Чертеж не создан" & Chr(13) & Chr(13) & "Создать чертеж?", vbYesNo, "Выберите действие")

                    Dim DrawingViewScale As Double
                    If oOcur.DefinitionDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                        DrawingViewScale = 1
                    ElseIf oOcur.DefinitionDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                        DrawingViewScale = 1 / 10
                    End If

                    If wopr = vbYes Then
                        Dim oPartDoc As Document = m_inventorApplication.Documents.Open(sDoc.FullFileName, False) ' False True
                        Dim oDrawingDoc = m_inventorApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, Templatka, True)
                        Dim oSheet As Sheet = oDrawingDoc.Sheets.Item(1)
                        Dim oBaseViewOptions As NameValueMap
                        oBaseViewOptions = m_inventorApplication.TransientObjects.CreateNameValueMap
                        Dim oPoint1 As Point2d = m_inventorApplication.TransientGeometry.CreatePoint2d(12.0#, 20.0#) 'front view
                        Dim oView1 As Inventor.DrawingView = oSheet.DrawingViews.AddBaseView(oPartDoc, oPoint1, DrawingViewScale, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                        oPartDoc.Close(True)

                        For Each prop In oDrawingDoc.PropertySets.Item("Inventor User Defined Properties")
                            prop.Delete() ' удаление всех пользовательских свойств в создаваемом чертеже
                        Next
                    End If
                End Try

            Next
        Else 'SelectSet.Count <> 0     =====================================================================================================================================
            Dim ThisDoc As Document = m_inventorApplication.ActiveEditDocument
            Dim oPath As String = Strings.Left(ThisDoc.FullFileName, Len(ThisDoc.FullFileName) - 4) & ".idw"
            'MsgBox(ThisDoc.FullFileName, , "ThisDoc.FullFileName")
            'MsgBox(oPath, , "oPath")
            Try '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                m_inventorApplication.Documents.Open(oPath, True)
            Catch '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                If m_inventorApplication.ActiveDocumentType = DocumentTypeEnum.kPartDocumentObject Then
                    txtt = "Деталь"
                    Templatka = Dannje(9) '"E:\§ Inventor\Темплатки\!14 Дет.idw"
                ElseIf m_inventorApplication.ActiveDocumentType = DocumentTypeEnum.kAssemblyDocumentObject Then
                    txtt = "Сборка"
                    Templatka = Dannje(10) '"E:\§ Inventor\Темплатки\!14 СБ.idw"
                End If
                'MsgBox(System.IO.Path.GetFileName(ThisDoc.FullFileName), , "GetFileName")
                Dim sFileName As String = System.IO.Path.GetFileName(ThisDoc.FullFileName)
                Dim wopr As MsgBoxResult = MsgBox(txtt & Chr(13) & sFileName & Chr(13) & Chr(13) & "Чертеж не создан" & Chr(13) & Chr(13) & "Создать чертеж?", vbYesNo, "Выберите действие")
                If wopr = vbYes Then
                    Dim DrawingViewScale As Double
                    If txtt = "Деталь" Then
                        DrawingViewScale = 1
                    ElseIf txtt = "Сборка" Then
                        DrawingViewScale = 1 / 10
                    End If

                    ''				oPartDoc = ThisApplication.Documents.Open(sDoc.FullFileName, False)' False True
                    Dim oDrawingDoc = m_inventorApplication.Documents.Add(DocumentTypeEnum.kDrawingDocumentObject, Templatka, True)
                    Dim oSheet = oDrawingDoc.Sheets.Item(1)
                    Dim oBaseViewOptions As NameValueMap = m_inventorApplication.TransientObjects.CreateNameValueMap
                    Dim oPoint1 As Inventor.Point2d = m_inventorApplication.TransientGeometry.CreatePoint2d(12.0#, 20.0#) 'front view
                    Dim oView1 = oSheet.DrawingViews.AddBaseView(ThisDoc, oPoint1, DrawingViewScale, ViewOrientationTypeEnum.kFrontViewOrientation, DrawingViewStyleEnum.kHiddenLineRemovedDrawingViewStyle)
                    ''		oPartDoc.Close(True)
                    For Each prop In oDrawingDoc.PropertySets.Item("Inventor User Defined Properties")
                        prop.Delete() ' удаление всех пользовательских свойств в создаваемом чертеже
                    Next
                End If

            End Try '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        End If

        MsgBox("Все чертежи открыты", MessageBoxIcon.Information, "Правило")

        'Dim SH As IWshRuntimeLibrary.WshShell
        'Dim Res As Long
        'SH = New IWshRuntimeLibrary.WshShell
        'Res = SH.Popup("Все чертежи открыты", 1, "Выполнено", MessageBoxIcon.Information)
        'SH = Nothing

vihod:
        '======================================= 
        '======================================= 
        '======================================= 


        ' ''If m_inventorApplication.ActiveDocument.SelectSet.Count <> 0 Then
        ' ''    For Each oOcur In m_inventorApplication.ActiveDocument.SelectSet
        ' ''        'sDoc = oOcur.Definition.Document
        ' ''        'Otkr = String.Left(oOcur.Definition.Document.FullFileName, Len(oOcur.Definition.Document.FullFileName) - 4) + ".idw"  ' путь к чертежу
        ' ''        Try
        ' ''            m_inventorApplication.Documents.Open(Strings.Left(oOcur.Definition.Document.FullFileName, Len(oOcur.Definition.Document.FullFileName) - 4) + ".idw", True)
        ' ''        Catch ex As Exception
        ' ''            MsgBox("Чертеж отсутствует", , "Ошибка")
        ' ''        End Try
        ' ''    Next
        ' ''Else
        ' ''    Try
        ' ''        m_inventorApplication.Documents.Open(Strings.Left(m_inventorApplication.ActiveDocument.FullFileName, Len(m_inventorApplication.ActiveDocument.FullFileName) - 4) & ".idw", True)
        ' ''    Catch ex As Exception
        ' ''        MsgBox("Чертеж отсутствует", , "Ошибка")
        ' ''    End Try
        ' ''End If
    End Function

    Private Sub ОткрытьЭлементыБезЧертежейToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОткрытьЭлементыБезЧертежейToolStripMenuItem.Click
        If m_inventorApplication.ActiveDocument.DocumentType <> DocumentTypeEnum.kAssemblyDocumentObject Then
            MessageBox.Show("Правило необходимо запускать в сборке", "iLogic")
            Exit Sub
        End If

        Dim oAsmDoc As AssemblyDocument
        oAsmDoc = m_inventorApplication.ActiveDocument
        Dim oAsmName As String = Strings.Left(oAsmDoc.DisplayName, Len(oAsmDoc.DisplayName) - 4)

        Dim oRefDocs As DocumentsEnumerator = oAsmDoc.AllReferencedDocuments
        Dim oRefDoc As Document
        '===============================================================================================
        Dim PapkaCBOR As String = Strings.Left(oAsmDoc.FullFileName, oAsmDoc.FullFileName.LastIndexOf("\") + 1) 'm_inventorApplication.ActiveDocument.FullFileName
        'MsgBox(PapkaCBOR & "PapkaCBOR", , oAsmDoc.DisplayName)
        Dim iii As Integer = 0
        Dim idwPathName As String
        Dim dwgPathName As String
        For Each oRefDoc In oRefDocs
            Dim PapkaDET As String = Strings.Left(oRefDoc.FullDocumentName, Len(PapkaCBOR))
            If PapkaDET = PapkaCBOR Then
                idwPathName = Strings.Left(oRefDoc.FullFileName, Len(oRefDoc.FullFileName) - 3) & "idw"
                dwgPathName = Strings.Left(oRefDoc.FullFileName, Len(oRefDoc.FullFileName) - 3) & "dwg"
                If System.IO.File.Exists(idwPathName) = False And System.IO.File.Exists(dwgPathName) = False Then
                    'MsgBox(PapkaDET & "  PapkaDET" & Chr(13) & PapkaCBOR & " PapkaCBOR" & Chr(13) & idwPathName & " idwPathName", , oRefDoc.DisplayName)

                    Dim oOtkrDoc As Document = m_inventorApplication.Documents.Open(oRefDoc.FullDocumentName, True)
                    '		MsgBox(PapkaDET & "  PapkaDET" & Chr(13) & PapkaCBOR & "PapkaCBOR")
                    'oOtkrDoc = m_inventorApplication.Documents.Open(oRefDoc.FullDocumentName, True)
                    iii = iii + 1
                    If iii = 5 Then If MsgBox("Открыто уже много элементов. Продолжить?", vbYesNo, "Вопрос") = vbYes Then iii = 0 Else Return
                    '		Else
                End If
            End If
        Next
        '- - - - - - - - - - - - -Top Level Drawing - - - - - - - - - - - -
        MessageBox.Show("Все файлы открыты", "iLogic")
    End Sub

    Private Sub AutoNAMEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AutoNAMEToolStripMenuItem.Click
        AvtoNameSB(m_inventorApplication.ActiveDocument.DisplayName)
    End Sub

    Private Sub ProverkaPro4ih(ByVal AssDoc As Document, Optional ByVal CustomName As String = "", Optional ByRef CustomValue As String = "", Optional ByRef CustExist As Boolean = False)
        Dim oCustomPropSet As PropertySet
        oCustomPropSet = AssDoc.PropertySets.Item("Inventor User Defined Properties")
        Dim oCustProp As [Property]
        For Each oCustProp In oCustomPropSet
            If CustomName = oCustProp.Name Then
                CustomValue = oCustProp.Value
                CustExist = True
                Exit Sub
                'Else
            End If
        Next 'oCustProp
        CustExist = False
    End Sub

    Private Sub ПросмотретьНастройкиToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ПросмотретьНастройкиToolStripMenuItem.Click
        My.Settings.Reload()
        MsgBox("Количество пробелов - " & My.Settings.Koli4estvoProbelov & Chr(13) & _
               "Автор - " & My.Settings.AutorFamilie & Chr(13) & _
               "Разработчик - " & My.Settings.Razrabot4ikFamilie & Chr(13) & _
               "Проверяющий - " & My.Settings.ProverilFamilie & Chr(13) & _
               "Нормоконтролер - " & My.Settings.NormokontrolFamilie & Chr(13) & _
               "Утверждающий - " & My.Settings.UtverdilFamilie, , "Настройки")
    End Sub

    Private Sub ПроверкаСвойств(Optional ByRef ProverkaVSEH As Boolean = False)
        Dim Korjavye(0) As String
        Dim IndexOpenFiles As Integer = -1

        Dim oAsmDoc As AssemblyDocument = m_inventorApplication.ActiveDocument
        Dim oRefDocs As DocumentsEnumerator
        If ProverkaVSEH Then
            oRefDocs = oAsmDoc.AllReferencedDocuments
        Else
            oRefDocs = oAsmDoc.ReferencedDocuments
        End If

        'Dim FileName As String = "C:\YourFile.txt"
        Dim SBInfo As New System.IO.FileInfo(oAsmDoc.FullFileName)
        Dim DirSB As String = SBInfo.DirectoryName 'Узнали полное имя файла

        Dim oRefDoc As Document
        For Each oRefDoc In oRefDocs
            Dim ИмяФайлаСР As String = Split(oRefDoc.FullFileName, "\")(UBound(Split(oRefDoc.FullFileName, "\")))
            Dim ИмяФайла As String = Strings.Left(ИмяФайлаСР, Len(ИмяФайлаСР) - 4)
            Dim ВторойПробел As Integer = InStr(InStr(ИмяФайла, " ") + 1, ИмяФайла, " ")
            Dim Обозначение As String = Strings.Left(ИмяФайла, ВторойПробел - 1)
            Dim Наименование As String = Mid(ИмяФайла, ВторойПробел + 1)

            Dim PropDescription As String = oRefDoc.PropertySets.Item("Design Tracking Properties").Item("Description").Value
            Dim PropPartNumber As String = oRefDoc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value

            Dim DirFile As String = Strings.Left(oRefDoc.FullFileName, Len(DirSB))
            'MsgBox("DirSB -   " & DirSB & Chr(13) & "DirFile - " & DirFile & Chr(13) & "oRefDoc - " & oRefDoc.FullFileName)
            'Dim FileInfo As New System.IO.FileInfo(oAsmDoc.FullFileName)
            'Dim DirSB As String = FileInfo.DirectoryName 'Узнали полное имя файла

            'MsgBox("oDoc - " & oRefDoc.DisplayName & Chr(13) & "Description - " & PropDescription & Chr(13) & "Part Number - " & PropPartNumber & Chr(13) & _
            '"ИмяФайла - " & ИмяФайла & Chr(13) & "Обозначение - " & Обозначение & Chr(13) & "Наименование - " & Наименование, , "Настройки")

            If DirFile = DirSB Then
                'MsgBox("PropPartNumber -   " & PropPartNumber & Chr(13) & "Обозначение - " & Обозначение & Chr(13) & "PropDescription - " & PropDescription & Chr(13) & "Наименование - " & Наименование, , "DirFile = DirSB")

                If PropPartNumber <> Обозначение Or PropDescription <> Наименование Then
                    'If UBound(Korjavye) = 1 Then IndexOpenFiles = 1 Else IndexOpenFiles = UBound(Korjavye) + 1
                    'MsgBox("PropPartNumber <> Обозначение", , "PropPartNumber <> Обозначение")
                    IndexOpenFiles = IndexOpenFiles + 1
                    ReDim Preserve Korjavye(IndexOpenFiles)
                    Korjavye(IndexOpenFiles) = oRefDoc.FullFileName
                End If
            End If

        Next

        If UBound(Korjavye) > -1 Then
            If MsgBox("Открыть неисправные?", vbYesNo, "Вопрос") = vbYes Then
                For Each iii In Korjavye
                    If Not iii = "" Then m_inventorApplication.Documents.Open(iii, True)
                Next
            Else
                If MsgBox("Записать полные пути неисправных в файл?", vbYesNo, "Вопрос") = vbYes Then
                    Dim writePath As String = Replace(oAsmDoc.FullFileName, ".iam", ".txt")
                    Dim text As String = ""
                    For Each iii In Korjavye
                        'If  = "" Then

                        'End If
                    Next
                    Using writer As New IO.StreamWriter(writePath, False, System.Text.Encoding.UTF8)
                        writer.WriteLine(text)
                    End Using
                    If MsgBox("Открыть этот файл?", vbYesNo, "Вопрос") = vbYes Then m_inventorApplication.Documents.Open(writePath, True)
                End If
            End If
        Else
            MsgBox("Несоответствий не выялено.", , "Конец")
        End If
    End Sub

    Private Sub ИзменитьПутиКШаблонамЧертежей()
        My.Settings.Reload()
        My.Settings.Put4ertegDet = sOpenFileDialog("Путь к шаблону чертежа детали")
        My.Settings.Put4ertegSB = sOpenFileDialog("Путь к шаблону чертежа сборки")
        'My.Settings.Put4ertegDop1 = sOpenFileDialog("Путь к шаблону чертежа") 'InputBox("Фамилия разработчика", "Ввод данных", My.Settings.Razrabot4ikFamilie)
        'My.Settings.Put4ertegDop2 = sOpenFileDialog("Путь к шаблону чертежа") 'InputBox("Фамилия проверяющего", "Ввод данных", My.Settings.ProverilFamilie)
        My.Settings.Save()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Dim SH As IWshRuntimeLibrary.WshShell
        Dim Res As Long
        SH = New IWshRuntimeLibrary.WshShell
        Res = SH.Popup(Text:="Click Me", SecondsToWait:=1, Title:="Hello, World", Type:=vbOKOnly)
        SH = Nothing
        'IWshRuntimeLibrary.WshShell.Popup("Процесс не найден") ', vbOKCancel + vbInformation)
    End Sub

    Private Sub ОткрытьФайлПоИмениToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ОткрытьФайлПоИмениToolStripMenuItem.Click
        Open_for_name()
    End Sub

    Private Sub КопированиесиндексомToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles КопированиесиндексомToolStripMenuItem.Click
        InsexCopys()
    End Sub
End Class

