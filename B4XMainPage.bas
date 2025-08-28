B4A=true
Group=Default Group
ModulesStructureVersion=1
Type=Class
Version=9.85
@EndOfDesignText@
#Region Shared Files
#Macro: Title, Export, ide://run?File=%B4X%\Zipper.jar&Args=%PROJECT_NAME%.zip
#Macro: Title, GitHub, ide://run?file=%WINDIR%\System32\cmd.exe&Args=/c&Args=github&Args=..\..\
#Macro: Title, Sync Files, ide://run?file=%WINDIR%\System32\Robocopy.exe&args=..\..\Shared+Files&args=..\Files&FilesSync=True
'#Macro: Title, JsonLayouts folder, ide://run?File=%WINDIR%\explorer.exe&Args=%PROJECT%\JsonLayouts
'#Macro: After Save, Sync Layouts, ide://run?File=%ADDITIONAL%\..\B4X\JsonLayouts.jar&Args=%PROJECT%&Args=%PROJECT_NAME%
'#CustomBuildAction: folders ready, %WINDIR%\System32\Robocopy.exe,"..\..\Shared Files" "..\Files"
#End Region
'B4J project: DB to XLSX Converter
'Requirements: jPOI library (for Excel export) + jSQL library (for SQLite)
Sub Class_Globals
	Private Root As B4XView
	Private xui As XUI
	Private sql As SQL
End Sub

Public Sub Initialize
'	B4XPages.GetManager.LogEvents = True
End Sub

Private Sub B4XPage_Created (Root1 As B4XView)
	Root = Root1
	Root.LoadLayout("MainPage")
	InitData
End Sub

Sub InitData
    ' Connect to database
    sql.InitializeSQLite(File.DirApp, "strings.db", True)
End Sub

Sub DB2XLSX
	' Path to output Excel file
	Dim outPath As String = File.Combine(File.DirApp, "strings.xlsx")
	
	' Load table data
	Dim rs As ResultSet = sql.ExecQuery("SELECT key, lang, value FROM data")
    
	' Store in a Map of Maps: key -> (lang -> value)
	Dim allData As Map
	allData.Initialize
    
	Do While rs.NextRow
		Dim k As String = rs.GetString("key")
		Dim lang As String = rs.GetString("lang")
		Dim val As String = rs.GetString("value")
        
		If allData.ContainsKey(k) = False Then
			Dim inner As Map
			inner.Initialize
			allData.Put(k, inner)
		End If
		Dim innerMap As Map = allData.Get(k)
		innerMap.Put(lang, val)
	Loop
	rs.Close
    
	' Find all unique langs
	Dim langs As List
	langs.Initialize
	For Each inner As Map In allData.Values
		For Each l As String In inner.Keys
			If langs.IndexOf(l) = -1 Then langs.Add(l)
		Next
	Next
    
	' Create Excel workbook
	Dim XL As XLUtils : XL.Initialize
	Dim Workbook As XLWorkbookWriter = XL.CreateWriterBlank
	Dim sheet1 As XLSheetWriter = Workbook.CreateSheetWriterByName("Sheet1")

	' Write header
	sheet1.PutString(XL.AddressZero(0, 0), "key")
	For i = 0 To langs.Size - 1
	    sheet1.PutString(XL.AddressZero(i + 1, 0), langs.Get(i))
	Next

	' Write rows
	Dim rowIndex As Int = 1
	For Each key As String In allData.Keys
		sheet1.PutString(XL.AddressZero(0, rowIndex), key)
		Dim inner As Map = allData.Get(key)
		For i = 0 To langs.Size - 1
			Dim lang As String = langs.Get(i)
			If inner.ContainsKey(lang) Then
				sheet1.PutString(XL.AddressZero(i + 1, rowIndex), inner.Get(lang))
			End If
		Next
		rowIndex = rowIndex + 1
	Next
	
	' Autosize columns
    For i = 0 To langs.Size
		sheet1.AutoSizeColumn(i)
	Next
	
	' Delete existing file
	File.Delete(File.DirApp, "strings.xlsx")
	
	' Save Excel file
	Dim f As String = Workbook.SaveAs(File.DirApp, "strings.xlsx", True)
	Wait For (XL.OpenExcel(f)) Complete (Success As Boolean)

	Log("Excel file saved: " & outPath)
	xui.MsgboxAsync("Conversion Completed!" & CRLF & outPath, "Done")
	
	sql.Close
End Sub

Private Sub BtnDB2XLSX_Click
	DB2XLSX
End Sub