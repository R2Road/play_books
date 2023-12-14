REM  *****  LibreOffice VBA  *****

Option Explicit


Const StartX = 0
Const StartY = 2



Type Size
    w as Integer
    h as Integer
End Type



Function CalculateSheetActiveArea( sheet as Variant ) as Size
	
	Dim ret as Size
	Dim i, j as Long
	
	'
	' W
	'
	j = sheet.Rows.Count - 1
	For i = StartX to j
		If sheet.getCellByPosition( i, 0 ).String = "" Then
			Exit For
		EndIf
	Next i
	ret.w = i - 1
		
	'
	' H
	'
	j = sheet.Rows.Count - 1
	For i = StartY to j
		If sheet.getCellByPosition( 0, i ).String = "" Then
			Exit For
		EndIf
	Next i
	ret.h = i - 1
	
	'
	' Return
	'
	CalculateSheetActiveArea = ret
	
End Function



Function LoadFile( load_fine_name as String, out_file as Variant )

	'
	' File Open
	'
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/" & load_fine_name )
	'MsgBox( file_path )
	
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	Dim header_pf As Variant
	Set header_pf = file_system.OpenTextFile(file_path, file_system.ForReading)
	
	
	'
	'
	'
	out_file.WriteLine( header_pf.ReadAll() )
	
	
	'
	' File Close
	'
	header_pf.CloseFile()
	header_pf = header_pf.Dispose()	

End Function



Function ExportList( sheet as Variant, active_area_h as Integer, key_index as Integer, sub_index as Integer, out_file as Variant )
	
	'
	'
	'
	Dim title, company, result as String
	Dim i, j as Integer
	For i = StartY to active_area_h
    
		'
		' Check Export Flag
		'
		If sheet.getCellByPosition( StartX, i ).String = "x" Then
			GoTo Continue
		EndIf
    	
    	
    	
		'
		' Empty is End
		'
		If sheet.getCellByPosition( key_index, i ).String = "" Then
			Exit For
		EndIf
    	
    	
    	
		'
		' Build Info
		'
		result = _
				"####" _
			_
			_
			& 	" " _
			&	"[ " _
				& 			sheet.getCellByPosition( 1, i ).String _
				& 	" ~ " & sheet.getCellByPosition( 2, i ).String _
			& " ]" _
			_
			_
			& 	" " _
			&  	sheet.getCellByPosition( 3, i ).String _
    	
    	
    	
		'
		'
		'
		out_file.WriteLine( result )
		'MsgBox( result )
    	
	Continue:
	Next i
	
End Function



Sub Main

	'
	'
	'
	GlobalScope.BasicLibraries.LoadLibrary("Tools") ' for Tools
	GlobalScope.BasicLibraries.LoadLibrary("ScriptForge") ' for FileSystem
	
	
	
	'
	' File Open
	'
	Dim file_path as String
	file_path = ( Tools.Strings.DirectoryNameoutofPath(ThisComponent.getURL(),"/" ) & "/" & "README.md" )
	MsgBox( file_path )
	
	Dim file_system As Variant
	file_system = CreateScriptService("FileSystem")
	
	Dim pf As Variant
	Set pf = file_system.CreateTextFile(file_path, Overwrite := true)
	
	
	
	'
	'
	'
	LoadFile( "header.txt", pf )
	
	
	'
	' Sheet
	'
	Dim sheet as Object
	sheet = ThisComponent.Sheets.getByName( "list" )
    
    
	'
	' Max X, Y
	'
	Dim active_area as Size
	active_area = CalculateSheetActiveArea( sheet )
	MsgBox( "Active Area : " & StartX & " : " & StartY & " ~ " & active_area.w & " : " & active_area.h )
	
	
	'
	' 게임 수 출력
	'
	pf.WriteLine( "* " & active_area.h - StartY & " 개" )
    
    
	'
	' Export List
	'
	On Error GoTo ERROR_END 'Error 발생시 File 해제 용도
		
		'
		'
		'
		pf.WriteLine( Chr( 10 ) & Chr( 10 ) )
		pf.WriteLine( "<br/><br/>" )
		
		
		'
		' Write : Korean List
		'
		pf.WriteLine( "## 목록" & Chr( 10 ) )
		ExportList( sheet, active_area.h, 1, 2, pf )
	
	ERROR_END:
    
    
	'
	' File Close
	'
	pf.CloseFile()
	pf = pf.Dispose()
    
End Sub











