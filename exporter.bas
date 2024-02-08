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
	j = sheet.Rows.Count - 1 ' 0 부터 시작이라 1 빼준다.
	For i = StartX to j
		If sheet.getCellByPosition( i, 0 ).String = "" Then
			Exit For
		EndIf
	Next i
	ret.w = i - 1
		
	'
	' H
	'
	j = sheet.Rows.Count - 1 ' 0 부터 시작이라 1 빼준다.
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



Function CalculateActiveDataCount( sheet as Variant ) as Integer
	
	Dim ret as Integer : ret = 0
		
	'
	' H
	'
	Dim cur_y as Long
	Dim end_y as Long : end_y = sheet.Rows.Count - 1 ' 0 부터 시작이라 1 빼준다.
	
	For cur_y = StartY to end_y
		If sheet.getCellByPosition( 0, cur_y ).String = "" Then
			Exit For
		ElseIf sheet.getCellByPosition( 0, cur_y ).String = "o" Then
			ret = ret + 1
		EndIf
	Next cur_y
	
	'
	' Return
	'
	CalculateActiveDataCount = ret
	
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
	Dim last_year as String
	Dim current_year as String
	Dim last_month as String
	Dim current_month as String
	
	
	'
	'
	'
	Dim title, company, result as String
	
	Dim start_y as Integer : start_y = active_area_h
	Dim end_y as Integer : end_y = StartY
	Dim current_y as Integer
	For current_y = start_y to end_y step -1
    
		'
		' Check Export Flag
		'
		If sheet.getCellByPosition( StartX, current_y ).String = "x" Then
			GoTo Continue
		EndIf
    	
    	
    	
		'
		' Empty is End
		'
		If sheet.getCellByPosition( key_index, current_y ).String = "" Then
			Exit For
		EndIf
		
		
		
		'
		' REF
		'
		' Latex Color : https://www.overleaf.com/learn/latex/Using_colors_in_LaTeX
		
		'
		' Year
		'
		current_year = Year( sheet.getCellByPosition( 2, current_y ).Value )
		If current_year <> last_year Then
		
			out_file.WriteLine( "####" & " " & "${\sf\color{RubineRed} {" & current_year & "}}$" )
			
			last_year = current_year
			
			last_month = ""
			
		EndIf
		
		
		
		'
		' Month
		'
		current_month = Month( sheet.getCellByPosition( 2, current_y ).Value )
		If current_month <> last_month Then
		
			out_file.WriteLine( "####" & " " & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & " " & "${\sf\color{Cyan} {" & current_month & "}}$" )
			
			last_month = current_month
			
		EndIf
    	
    	
    	
		'
		' Build Info
		'
		result = _
					"[ " _
				& 			sheet.getCellByPosition( 1, current_y ).String _
				& 	" ~ " & sheet.getCellByPosition( 2, current_y ).String _
			& " ]" _
			_
			_
			& 	" " _
			&  	sheet.getCellByPosition( 3, current_y ).String _
			&	Chr( 10 )
    	
    	
    	
		'
		'
		'
		out_file.WriteLine( result )
		'MsgBox( result )
    	
	Continue:
	Next current_y
	
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
	Dim active_data_count as Integer
	active_data_count = CalculateActiveDataCount( sheet )
	pf.WriteLine( "* " & active_data_count & " 개" )
    
    
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
		' Sort
		'
			'
			' Range
			'
			Dim range4sort
				range4sort = sheet.getCellRangeByPosition( StartX, StartY, active_area.w, active_area.h )
				
			'
		    ' Sort Field
		    '
		    Dim sort_field(0) as new com.sun.star.util.SortField
				sort_field(0).Field = 2
			    sort_field(0).SortAscending = TRUE
			    sort_field(0).FieldType = com.sun.star.util.SortFieldType.ALPHANUMERIC 'com.sun.star.util.SortFieldTypeNUMERIC
		    
		    '
		    ' Description
		    '
		    Dim sort_description(0) as new com.sun.star.beans.PropertyValue
			    sort_description(0).Name = "SortFields"
			    sort_description(0).Value = sort_field()
		    
		    '
		    ' Do
		    '
		    range4sort.Sort( sort_description() )
		
		
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











