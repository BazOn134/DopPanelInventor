Module Globals
    'собственно сам Инвентор
    Public m_inventorApplication As Inventor.Application
    'окно
    Public m_Window As Inventor.DockableWindow
    'форма
    Public m_Form As UserControl1

    Public Function sOpenFileDialog(Optional ByVal sTitle As String = "") As String ' окно "Открыть файл"
        Dim OFD1 As New Windows.Forms.OpenFileDialog()
        OFD1.InitialDirectory = m_inventorApplication.DesignProjectManager.ActiveDesignProject.TemplatesPath '.FileOptions.TemplatesPath '' As String     Member of Inventor.DesignProject'"C:" ' Указываем начальную папку
        If sTitle = "" Then OFD1.Title = "Выбор файла" Else OFD1.Title = sTitle ' Указываем заголовок
        OFD1.Filter = "Чертежи|*.idw|Все файлы (*.*)|*.*" '"HTML файлы|*.html; *.htm|Текстовые файлы (*.txt)|*.txt|Все файлы (*.*)|*.*" ' При помощи фильтра можно отбросить ненужные типы файлов(в нашем случае делаем выбор из списка(HTML файлы, Текстовые файлы, Все файлы)).
        OFD1.FilterIndex = 1 ' Если есть список выбора типов, то можно указать какой тип будет выбран при загрузке диалога
        If OFD1.ShowDialog = Windows.Forms.DialogResult.OK Then sOpenFileDialog = OFD1.FileName Else sOpenFileDialog = "" 'MsgBox(OFD1.FileName) ' Открываем диалог выбора файлов(OFD1.ShowDialog), если был выбран файл и нажата кнопка 'Открыть'(= DialogResult.OK), то показываем полный путь выбранного файла(MsgBox(OFD1.FileName))(зная путь, можно открыть файл(если диалог SaveFileDialog - то сохранить))
    End Function


End Module

