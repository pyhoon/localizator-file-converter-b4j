B4A=true
Group=Default Group
ModulesStructureVersion=1
Type=Class
Version=9.85
@EndOfDesignText@
#Region Shared Files
#Macro: Title, Export, ide://run?File=%B4X%\Zipper.jar&Args=%PROJECT_NAME%.zip
#Macro: Title, GitHub, ide://run?file=%WINDIR%\System32\cmd.exe&Args=/c&Args=github&Args=..\..\
#End Region

Sub Class_Globals
	Private xui              As XUI
	Private fx               As JFX
	Private fc               As FileChooser
	Private Root             As B4XView
	Private LblMessage       As B4XView
	Private txtExcel         As TextField
	Private txtSQLite        As TextField
	Private settings         As Map
	Private BtnBrowseExcel   As Button
	Private BtnBrowseSQLite  As Button
	Private BtnExcelToSQLite As Button
	Private BtnSQLiteToExcel As Button
End Sub

Public Sub Initialize
'	B4XPages.GetManager.LogEvents = True
End Sub

Private Sub B4XPage_Created (Root1 As B4XView)
	Root = Root1
	Root.LoadLayout("MainPage")
	B4XPages.SetTitle(Me, "Localizator File Converter v2.00")
	fc.Initialize
	settings.Initialize
	BtnBrowseExcel.MouseCursor = fx.Cursors.HAND
	BtnBrowseSQLite.MouseCursor = fx.Cursors.HAND
	BtnExcelToSQLite.MouseCursor = fx.Cursors.HAND
	BtnSQLiteToExcel.MouseCursor = fx.Cursors.HAND
	If File.Exists(File.DirApp, "settings.json") Then
		Dim str As String = File.ReadString(File.DirApp, "settings.json")
		settings = str.As(JSON).ToMap
		txtExcel.Text = settings.Get("xlsx")
		txtSQLite.Text = settings.Get("sqlite")
	Else
		txtExcel.Text = File.Combine(File.DirApp, "strings.xlsx")
		txtSQLite.Text = File.Combine(File.DirApp, "strings.db")
	End If
	'InitData
End Sub

Private Sub Browse (tf As TextField, extension As String)
	If File.Exists(File.GetFileParent(tf.Text), "") Then fc.InitialDirectory = File.GetFileParent(tf.Text)
	fc.SetExtensionFilter(extension, Array($"*.${extension}"$))
	Dim form As Form = B4XPages.GetNativeParent(Me)
	Dim path As String = fc.ShowSave(form)
	If path <> "" Then tf.Text = path
	If extension = "xlsx" Then
		settings.Put("xlsx", path)
	End If
	If extension = "db" Then
		settings.Put("sqlite", path)
	End If
	File.WriteString(File.DirApp, "settings.json", settings.As(JSON).ToString)
End Sub

Private Sub BtnBrowseExcel_Click
	Browse(txtExcel, "xlsx")
End Sub

Private Sub BtnBrowseSQLite_Click
	Browse(txtSQLite, "db")
End Sub

Private Sub BtnExcelToSQLite_Click
	Main.SaveSqlite(txtExcel.Text, txtSQLite.Text)
End Sub

Private Sub BtnSQLiteToExcel_Click
	Main.SaveExcel(txtSQLite.Text, txtExcel.Text)
End Sub