Imports System.IO
Imports System.Data
Imports System.Data.OleDb

Module excel_to_csv

    Sub Main()
		'// CLEAR THE CONSOLE WINDOW
		Console.Clear
	
		
		'// SETUP SOME LOCAL VARIABLES THAT ARE NEEDED TO STORE PARAMETER VALUES
		Dim parameters() As String = Environment.GetCommandLineArgs
		Dim input_file As String = ""
		Dim output_folder As String = Environment.CurrentDirectory
		Dim worksheets_to_convert As String = "[Sheet1$]"
		Dim delimiter As String = ","
		Dim header_inclusion As Boolean = True
		Dim parameter_count As Integer = 0
		
		
		'// WALK THROUGH THE PARAMETERS AND ASSIGN TO LOCAL VARIABLES
		For Each item As String In parameters
			If item.ToUpper.StartsWith("/I:") Then
				input_file = Right(item, item.Length - 3).Trim
			End If
			
			If item.ToUpper.StartsWith("/O:") Then
				output_folder = Right(item, item.Length - 3).Trim
			End If
			
			If item.ToUpper.StartsWith("/W:") Then
				worksheets_to_convert = Right(item, item.Length - 3).Trim
			End If
			
			If item.ToUpper.StartsWith("/D:") Then
				delimiter = Right(item, item.Length - 3).Trim
			End If
			
			If item.ToUpper.ToUpper.StartsWith("/H:") Then
				Try
					header_inclusion = CBool(Right(item, item.Length - 3).Trim)
				Catch ex As Exception
					Console.Write("WARNING: Attribute Value Must Be Boolean" + vbCrLf)
				End Try
			End If
			
			If item.ToUpper.ToUpper.StartsWith("/?") Then
				input_file = ""
				Exit For
			End If
			'Console.Write(item.ToString + vbCrLf)
			parameter_count = parameter_count + 1
		Next
		'Exit Sub
		
		'// IF NO PARAMETERS SPECIFIED, SHOW USAGE NOTES
		If parameter_count <= 1 Or input_file.Trim.Length = 0 Then
			Dim application_build As FileVersionInfo = FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location)

			Dim str_usage As String = ""
			str_usage += vbCrLf + "Excel To CSV Converter v" + application_build.FileMajorPart.ToString + "." + application_build.FileMinorPart.ToString + vbCrLf
			str_usage += "Convert an Excel (.xls) file to one or more CSV files." + vbCrLf
			str_usage += vbCrLf
			str_usage += "USAGE:"  + vbTab + "excel_to_csv  /I:input file  /O:output folder  /W:""worksheet list""" + vbCrLf + vbTab + "/D:""delimiter"" /H:header inclusion" + vbCrLf
			str_usage += vbCrLf + vbCrLf
			str_usage += " /I:" + vbTab + "Specifies the input Excel file." + vbCrLf
			str_usage += " /O:" + vbTab + "Specifies the output folder." + vbCrLf
			str_usage += " /W:" + vbTab + "Specified the list of worksheets to be converted within the source file." + vbCrLf
			str_usage += " /D:" + vbTab + "Specifies the CSV file delimiter (default is comma)." + vbCrLf
			str_usage += " /H:" + vbTab + "Specifies if first row should be included in output (default is True)." + vbCrLf
			Console.Write(str_usage + vbCrLf)
			Exit Sub
		End If
		
		
		'// SET THE CONNECTION STRING
		Dim connection_string As String = ""
		connection_string = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + input_file + ";Extended Properties=Excel 8.0;"
		
		
		'// SEE IF WE CAN OPEN THE CONNECTION
		Dim data_connection As New OleDbConnection(connection_string)
		Try
			data_connection.Open()
			If data_connection.State = ConnectionState.Closed Then
				Console.Write(vbCrLf + "ERROR: Source File Closed Prematurely" + vbCrLf)
			Else
				'//
			End If
		Catch ex As Exception
		  Console.Write(vbCrLf + "ERROR: Invalid Source File" + vbCrLf)
		  Exit Sub
		End Try
		
		
		'// LOOP THROUGH THE SPECIFIED WORKSHEETS AND EXPORT THE DATA
		Dim str_sheet_array As Array = Split(worksheets_to_convert, ";")
		
		For Each str_sheet As String In str_sheet_array
			Dim data_adapter As New OleDbDataAdapter
			Dim data_set As New DataSet
			Dim temp_string As String = ""
			Dim count_sheets As Integer = 0
			Dim count_rows As Integer = 0
			Dim count_columns As Integer = 0
			Dim count_inner As Integer = 0
			Dim count_outer As Integer = 0
		
			If str_sheet.Replace(";","").Trim.Length > 0 Then
				
				'// TRY TO SELECT THE DATA AND FILL A DATASET
				Dim data_command As New OleDbCommand("SELECT * FROM [" + str_sheet.Replace("[","").Replace("]", "").Replace("$","") + "$]", data_connection)
				data_command.CommandType = CommandType.Text
			 
			 	Try
					data_adapter.SelectCommand = data_command
					data_adapter.Fill(data_set, "XLData")
				Catch ex As Exception
					Console.Write(vbCrLf + ex.Message + vbCrLf)
					Console.Write(vbCrLf + "ERROR: Invalid Worksheet Specified (" + str_sheet + ")" + vbCrLf)
					Exit Sub
				End Try
			
				
				'// COUNT HOW MANY ROWS AND COLUMNS EXIST IN THE DATASET
				count_rows = data_set.Tables(0).Rows.Count
				count_columns = data_set.Tables(0).Columns.Count
				
				
				'// DATA CONNECTION IS ALL GOOD, LET US CREATE THE OUTPUT FILE
				Dim data_stream As StreamWriter
				Try
					data_stream = New StreamWriter((output_folder + "\").Replace("\\", "\") + str_sheet.Replace("[","").Replace("]", "").Replace("$","").ToLower.Replace(" ","_") + ".csv")
				Catch
					data_stream = New StreamWriter((output_folder + "\").Replace("\\", "\") + "output_" + count_sheets.ToString.ToLower.Replace(" ","_") + ".csv")
				End Try
			 
				
				'// WRITE OUT COLUMN HEADERS
				If header_inclusion = True Then
					Dim str_header As String = ""
					For i As Integer = 0 To count_columns - 1 
						str_header += data_set.Tables(0).Columns(i).ColumnName.Replace(delimiter, " ").Trim + delimiter
					Next 
					data_stream.WriteLine(str_header)
				End If

				
				Try
					'// WRITE OUT THE DATA
					For count_outer = 0 To count_rows - 1
						temp_string = ""
						For count_inner = 0 To count_columns - 1
							temp_string += data_set.Tables(0).Rows(count_outer).Item(count_inner).ToString.Replace(delimiter, " ").Trim + delimiter
						Next
						'temp_string &= ","
						data_stream.WriteLine(temp_string)
					Next
					data_stream.Close()
				Catch ex As Exception
					'// SOMETHING GOT SCREWED UP, TRY TO WRITE THE ERROR LOG
					Try 
						Console.Write(ex.ToString)
						Dim log_stream As New StreamWriter(Environment.CurrentDirectory + "\error.log")
						log_stream.WriteLine(ex.ToString)
						log_stream.Close()
					Catch
						'// WHAT THE @#%$???  THAT FAILED TOO...
						Console.Write(vbCrLf + "ERROR: Log Could Not Be Written" + vbCrLf)
						Exit Sub
					End Try
				End Try
				
				
				'// RESET THE DATA COMMAND
				data_command.Dispose()
				data_command = Nothing
				
				count_sheets = count_sheets + 1
			End If
		Next
	 
		'// CLEAN UP AND GET OUT...
		data_connection.Close()
		data_connection.Dispose()
		data_connection = Nothing
 
	End Sub
	

End Module
