' Convert a Whatsapp log to Html
'
' Args = Text files to convert
' Output = Html files with the same names but with ".html" appended
'
' Author: Michel DESSAINTES
' Last modification: 2024-04-15

Option Explicit
Dim oArgs, oFso, oReader, oWriter, fileText, fileName, Regex, ThumbnailWidth, LargeWidth
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set oArgs = WScript.Arguments
Set oFso = CreateObject("Scripting.FileSystemObject")

ThumbnailWidth = "100px"
LargeWidth = "400px"

If oArgs.Count >= 1 Then
    for each fileName in oArgs
        Set oReader = oFso.OpenTextFile(fileName, ForReading)
        fileText = oReader.ReadAll
        oReader.Close()
        Set oReader = Nothing

        Set oWriter = oFso.OpenTextFile(fileName & ".html", ForWriting, True)
        oWriter.Write(" <!DOCTYPE html>" & _
            "<html>" & _
            "<head>" & _
            "<style>" & _
            "   .stamp  { background-color: Aqua;   color: blue; }" & _
            "   .author { background-color: Silver; color: blue; }" & _
            "   .small { width: " & ThumbnailWidth & "; }" & _
            "   .big   { width: " & LargeWidth & "; }" & _
            "   a, a:visited { position: relative; vertical-align: top; }" & _
            "   a .big { display: none; }" & _
            "   a:hover .big {" & _
            "       display: block;" & _
            "       position: absolute;" & _
            "       left: " & ThumbnailWidth & ";" & _
            "       border: 1px solid #666;" & _
            "   }" & _
			" .flex { display: flex; } " & _
            "</style>" & _
            "</head>" & _
            "<body>")

        Set Regex = New RegExp
        Regex.Global = True
        Regex.Ignorecase = True

        'Regex.Pattern : "\u200E" does not work!?
        Regex.Pattern = " ...([^ ]*?(jpg|jpeg|png|gif)) \(fichier joint\)": fileText = Regex.Replace(fileText, " <a class=flex href=""$1""><img class=small src=""$1""><img class=big src=""$1""></a> ")  ' Image
        Regex.Pattern = " ...([^ ]*?) \(fichier joint\)":   fileText = Regex.Replace(fileText, " <a href=""$1"">$1</a> ")   ' Other attachement
        Regex.Pattern = "(https?.*?)( |\n)":                fileText = Regex.Replace(fileText, "<a href=""$1"">$1</a>$2")   ' Url
        Regex.Pattern = "\n([0-9/, :]*?) - (.*?):":         fileText = Regex.Replace(fileText, "<br><span class=stamp>$1</span> - <span class=author>$2</span> :")
        Regex.Pattern = "_(.*?)_":                          fileText = Regex.Replace(fileText, "<i>$1</i>")                 ' Italic
        Regex.Pattern = "~(.*?)~":                          fileText = Regex.Replace(fileText, "<strike>$1</strike>")       ' Barré
        Regex.Pattern = "```(.*?)```":                      fileText = Regex.Replace(fileText, "<tt>$1</tt>")               ' Fonte fixe
        Regex.Pattern = "\*(.*?)\*":                        fileText = Regex.Replace(fileText, "<b>$1</b>")                 ' Gras
        ' The following lines should be placed after the above Replaces
        Regex.Pattern = "\n":                               fileText = Regex.Replace(fileText, "<br>")                      ' Saut de lignes
            
        oWriter.Write(fileText)
        oWriter.Write("</body></html>")
        oWriter.Close()
        Set oWriter = Nothing
    next
    MsgBox "End : " & oArgs.Count & " file(s)."
Else
    MsgBox "Usage: """ & WScript.ScriptName & """ <FileNames...>"
End If

