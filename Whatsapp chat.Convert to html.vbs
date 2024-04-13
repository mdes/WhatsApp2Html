' Convert a Whatsapp log to Html
'
' Args = Text files to convert
' Output = Html files with the same names but with ".html" appended
'
' !! This vbs file is to be saved in ASCII !!
'
' Evolutions: http://www.wickham43.net/hoverpopups.php
'
' Author: Michel DESSAINTES
' Last modification: 2024-04-13

Option Explicit
Dim oArgs, oFso, oReader, oWriter, fileText, fileName, Regex
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Set oArgs = WScript.Arguments
Set oFso = CreateObject("Scripting.FileSystemObject")

If oArgs.Count >= 1 Then
	for each fileName in oArgs
		Set oReader = oFso.OpenTextFile(fileName, ForReading)
		fileText = oReader.ReadAll
		oReader.Close()
		Set oReader = Nothing

		Set oWriter = oFso.OpenTextFile(fileName & ".html", ForWriting, True)
		oWriter.Write("	<!DOCTYPE html>" & _
			"<html>" & _
			"<head>" & _
			"<style>" & _
			"	.stamp  { background-color: Aqua;   color: blue; }" & _
			"	.author { background-color: Silver; color: blue; }" & _
			"	.small { width: 100px; }" & _
			"	.big   { width: 400px; }" & _
			"	#popup a, #popup a:visited {" & _
			"		position: relative;" & _
			"		vertical-align: top;" & _
			"	}" & _
			"	#popup a span {" & _
			"		display: none;" & _
			"	}" & _
			"	#popup a:hover span {" & _
			"		display: block;" & _
			"		position: absolute;" & _
			"		left: 100px;" & _
			"		width: 100;" & _
			"		border: 1px solid #666;" & _
			"		background: #e5e5e5;" & _
			"	}" & _
			"</style>" & _
			"</head>" & _
			"<body><div id=popup>")

		Set Regex = New RegExp
		Regex.Global = True
		Regex.Ignorecase = True

		'Regex.Pattern : "\u200E" does not work!?
		Regex.Pattern = " ...([^ ]*?(jpg|jpeg|png|gif)) \(fichier joint\)":	fileText = Regex.Replace(fileText, " <a href=""$1""><img class=small src=""$1""><span><img class=big src=""$1""></span></a> ")	' Image
		Regex.Pattern = " ...([^ ]*?) \(fichier joint\)":	fileText = Regex.Replace(fileText, " <a href=""$1"">$1</a> ")	' Other attachement
		Regex.Pattern = "(https?.*?)( |\n)":				fileText = Regex.Replace(fileText, "<a href=""$1"">$1</a>$2")	' Url
		Regex.Pattern = "\n([0-9/, :]*?) - (.*?):":			fileText = Regex.Replace(fileText, "<br><span class=stamp>$1</span> - <span class=author>$2</span> :")
		Regex.Pattern = "_(.*?)_":							fileText = Regex.Replace(fileText, "<i>$1</i>")					' Italic
		Regex.Pattern = "~(.*?)~":							fileText = Regex.Replace(fileText, "<strike>$1</strike>")		' Barré
		Regex.Pattern = "```(.*?)```":						fileText = Regex.Replace(fileText, "<tt>$1</tt>")				' Fonte fixe
		Regex.Pattern = "\*(.*?)\*":						fileText = Regex.Replace(fileText, "<b>$1</b>")					' Gras
		' À placer après les Replaces ci-dessus
		Regex.Pattern = "\n":								fileText = Regex.Replace(fileText, "<br>")						' Saut de lignes
			
		oWriter.Write(fileText)
		oWriter.Write("</div></body></html>")
		oWriter.Close()
		Set oWriter = Nothing
	next
    MsgBox "Fini : " & oArgs.Count & " fichier(s) traité(s)."
Else
    MsgBox "Usage: """ & WScript.ScriptName & """ <FileNames...>"
End If

