rem --------------------------
rem find the working directory 
rem --------------------------

Dim WshShell, strCurDir
Set WshShell = CreateObject("WScript.Shell")
strCurDir    = WshShell.CurrentDirectory


rem ---------------
rem parse file name
rem ---------------

set fso = createobject("scripting.filesystemobject")
basename = fso.getbasename(WScript.Arguments.Item(0))
texname = basename & ".tex"

rem -------------------
rem extract tex content
rem -------------------

set WD = CreateObject("Word.Application")
WD.ChangeFileOpenDirectory strCurDir
WD.Visible = False
set doc = WD.Documents.Open(WScript.Arguments.Item(0),False,True)
doc.SaveAs2 texname, 2
doc.Close()
WD.Quit()

rem ------------------------
rem compile the tex into pdf
rem ------------------------
err = WshShell.Run("pdflatex " & texname, , True)
If err <> 0 Then
	Wscript.Echo "Error in tex file, please check " + basename + ".log for details"
	Wscript.Quit
End If
err = WshShell.Run("bibtex " & basename, , True)
err = WshShell.Run("pdflatex " & texname, , True)
If err <> 0 Then
	Wscript.Echo "Error in tex file, please check " + basename + ".log for details"
	Wscript.Quit
End If
err = WshShell.Run("pdflatex " & texname, , True)
If err <> 0 Then
	Wscript.Echo "Error in tex file, please check " + basename + ".log for details"
	Wscript.Quit
End If


rem -------
rem cleanup
rem -------
Set WshShell = Nothing

