Option Explicit

Function GetCurrentDir()
	Dim wshell

	Set wshell = CreateObject("WScript.Shell")
	GetCurrentDir = wshell.CurrentDirectory
end Function

Sub Print(str)
	WScript.Echo str
End Sub
	
Sub Main()
	Dim xlApp
	Dim xlBook

	Set xlApp = CreateObject("Excel.Application")
	Set xlBook = xlApp.Workbooks.Open(GetCurrentDir() & "\\group_gen.xlsm", 0, False)
	xlApp.Run("ReadGroupOutputAndUpdateToStageEnemy")
	xlApp.DisplayAlerts = False
	xlBook.Save()
	xlBook.Close()
	xlApp.Quit()
	
	Set xlBook = Nothing
	Set xlApp = Nothing

	Print("finished!")
End Sub

Main
