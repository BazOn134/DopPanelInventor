Imports System.Windows.Forms
Imports Inventor

Public Class FormSvojstva

    'Private Property sv_Form As Form

    Private Sub ButonZakryt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButonZakryt.Click
        Me.Close()
    End Sub

    Private Sub FormSvojstva_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If m_inventorApplication.Documents.VisibleDocuments.Count = 0 Then Exit Sub
        'Dim oAssyDoc As AssemblyDocument = m_inventorApplication.ActiveEditDocument
        Dim FFName As String
        If m_inventorApplication.ActiveDocument.SelectSet.Count = 0 Then
            Me.B_Sledujushij.Visible = False
            Dim ActiveEditDoc As PropertySets = m_inventorApplication.ActiveEditDocument.PropertySets
            FFName = m_inventorApplication.ActiveEditDocument.FullFileName
            Dim FFNnameBR As String = Mid(FFName, InStrRev(FFName, "\") + 1, Len(FFName) - InStrRev(FFName, "\") - 4)
            Dim detAvtor As String = ActiveEditDoc.Item("Inventor Summary Information").Item("Author").Value
            Dim detRazrabotal As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Designer").Value
            Dim detProveril As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Checked By").Value
            Dim detNormokontrol As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Engr Approved By").Value
            Dim detUtverdil As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Mfg Approved By").Value
            Dim detTitle As String = ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value
            Dim detDescription As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value
            Dim detPartNumber As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value
            Dim detMaterial As String
            If m_inventorApplication.ActiveEditDocument.DocumentType = DocumentTypeEnum.kPartDocumentObject _
                Then detMaterial = m_inventorApplication.ActiveEditDocument.ComponentDefinition.Material.Name _
                Else detMaterial = ""

            TB_Name.Text = FFNnameBR
            TB_PartNumber.Text = ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value
            TB_Title.Text = ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value
            TB_Description.Text = ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value
            TB_Material.Text = ActiveEditDoc.Item("Design Tracking Properties").Item("Material").Value

            ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Designer").Value)
            ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Checked By").Value)
            ListBox1.Items.Add("")
            ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Engr Approved By").Value)
            ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Mfg Approved By").Value)

        Else
            Me.B_Sledujushij.Visible = True
            'If m_inventorApplication.ActiveDocument.SelectSet.Count = 0 Then MsgBox("Сначала выделите компоненты", MessageBoxIcon.Information, "Внимание") : Exit Sub
            Dim oSelSet As SelectSet = m_inventorApplication.ActiveEditDocument.SelectSet
            Dim oDoc As Document
            Dim oOcur As ComponentOccurrence
            For Each oOcur In oSelSet
                oDoc = oOcur.Definition.Document
                FFName = oDoc.FullFileName
                Otobragenie(FFName)
            Next
        End If

    End Sub

    Private Sub Otobragenie(Optional ByVal FFName As String = "")
        Dim ActiveEditDoc As PropertySets = m_inventorApplication.ActiveDocument.PropertySets
        Dim FFNnameBR As String = Mid(FFName, InStrRev(FFName, "\") + 1, Len(FFName) - InStrRev(FFName, "\") - 4)
        Dim detAvtor As String = ActiveEditDoc.Item("Inventor Summary Information").Item("Author").Value
        Dim detRazrabotal As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Designer").Value
        Dim detProveril As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Checked By").Value
        Dim detNormokontrol As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Engr Approved By").Value
        Dim detUtverdil As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Mfg Approved By").Value
        Dim detTitle As String = ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value
        Dim detDescription As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value
        Dim detPartNumber As String = ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value
        Dim detMaterial As String = m_inventorApplication.ActiveEditDocument.ComponentDefinition.Material.Name

        TB_Name.Text = FFNnameBR
        TB_PartNumber.Text = ActiveEditDoc.Item("Design Tracking Properties").Item("Part Number").Value
        TB_Title.Text = ActiveEditDoc.Item("Inventor Summary Information").Item("Title").Value
        TB_Description.Text = ActiveEditDoc.Item("Design Tracking Properties").Item("Description").Value
        TB_Material.Text = ActiveEditDoc.Item("Design Tracking Properties").Item("Material").Value

        ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Designer").Value)
        ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Checked By").Value)
        ListBox1.Items.Add("")
        ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Engr Approved By").Value)
        ListBox1.Items.Add(ActiveEditDoc.Item("Design Tracking Properties").Item("Mfg Approved By").Value)
    End Sub

End Class