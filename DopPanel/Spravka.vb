Module Spravka


    'MsgBox(IO.Path.GetExtension(sFn)) ' расширение файла
    'MsgBox(IO.Path.GetFileName(Application.ExecutablePath)) 'Как получить имя файла без полного пути
    'MsgBox(IO.Path.GetFileNameWithoutExtension(Application.ExecutablePath)) 'имя файла без расширения
    'MsgBox(IO.Path.GetPathRoot(Application.ExecutablePath)) 'имя корневого каталога для файла
    'IO.Directory.Delete("C:\Dir", True) 'удалить папку, даже если она не пуста
    'Dim Folders() As String = IO.Directory.GetDirectories("C:\Dir", "*", IO.SearchOption.TopDirectoryOnly) 'Получить из папки все директории
    '       2 последних параметра в функции являются необязательными, первый из них задает маску поиска.По умолчанию возвращаются все папки (любое имя).
    '       Второй параметр - глубина поиска, с IO.SearchOption.TopDirectoryOnly мы ищем папки только в текущей директории
    '       тогда как с IO.SearchOption.AllDirectories поиск будет осуществлен и во всех вложенных подкаталогах.По умолчанию поиск производится только в текущей папке.
    'Dim Files() As String = IO.Directory.GetFiles("C:\Dir", "*.txt", IO.SearchOption.TopDirectoryOnly) 'Получить все файлы из директории 
    '       Здесь все по аналогии с предыдущим примером, в том числе и работа с масками, с тем лишь различием, что у файлов имя включает также еще и расширение, 
    '       поэтому можно легко и удобно осуществлять поиск по типу файлов.

    'Как получить размер папки?
    'Public Function РазмерПапки(ByVal ПутьКПапке As String) As Integer
    '    Dim ИтоговыйРазмер As Integer
    '    Dim ИнфоПапки As IO.DirectoryInfo = New IO.DirectoryInfo(ПутьКПапке)
    '    Try
    '        For Each Файл In ИнфоПапки.GetFiles()
    '            ИтоговыйРазмер += Файл.Length
    '        Next
    '        For Each Папка In ИнфоПапки.GetDirectories()
    '            ИтоговыйРазмер += РазмерПапки(Папка.FullName)
    '        Next
    '    Catch Ex As Exception
    '    End Try
    '    Return (ИтоговыйРазмер)
    'End Function


    '===================POPUP окно ----------НАЧАЛО
    'Dim SH As IWshRuntimeLibrary.WshShell
    'Dim Res As Long
    '    SH = New IWshRuntimeLibrary.WshShell
    '    Res = SH.Popup(Text:="Click Me", SecondsToWait:=1, Title:="Hello, World", Type:=vbOKOnly)
    '    SH = Nothing
    '===================POPUP окно ----------КОНЕЦ

End Module
