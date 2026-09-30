Attribute VB_Name = "EmbedFonts"
' Пакетное встраивание шрифтов: сохраняет копии всех .pptx из папки с внедрёнными шрифтами.
' Запускать в PowerPoint (Alt+F11 → Файл → Импорт файла → F5) на машине с установленным Manrope.
Sub EmbedFontsInFolder()
    Dim folder As String, f As String, pres As Presentation
    folder = InputBox("Папка с презентациями:", "Встраивание шрифтов")
    If folder = "" Then Exit Sub
    If Right(folder, 1) <> "\" Then folder = folder & "\"
    MkDir folder & "embedded"
    f = Dir(folder & "*.pptx")
    Do While f <> ""
        Set pres = Presentations.Open(folder & f, WithWindow:=msoFalse)
        pres.SaveCopyAs folder & "embedded\" & f, ppSaveAsOpenXMLPresentation, msoTrue   ' msoTrue = EmbedTrueTypeFonts
        pres.Close
        f = Dir
    Loop
    MsgBox "Готово: " & folder & "embedded"
End Sub
