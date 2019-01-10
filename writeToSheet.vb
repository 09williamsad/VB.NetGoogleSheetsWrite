Imports Google.Apis.Sheets.v4
Imports Google.Apis.Sheets.v4.Data
    Public Sub insertToSheet()       	
	Static Dim Scopes As String() = {SheetsService.Scope.Spreadsheets} 'If changing the scope then delete  App_Data\MyGoogleStorage\.credentials\sheets.googleapis.com-dotnet-quickstart.json
	Dim ApplicationName As String = "Google Sheets API .NET Quickstart" 'Tutorial name for .net api
	Dim location = Server.MapPath("client_secret.json") 'Designate file with sheets api key and use it to setup authentication
	Using stream = New FileStream(location, FileMode.Open, FileAccess.Read) 'Read file and setup credentials for sheets api
		Dim credPath As String = System.Web.HttpContext.Current.Server.MapPath("/App_Data/MyGoogleStorage")
		credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json")
		Dim credential As UserCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None, New FileDataStore(credPath, True)).Result
	End Using
		
        Dim service = New SheetsService(New BaseClientService.Initializer() With {.HttpClientInitializer = credential, .ApplicationName = ApplicationName}) 'Create Google Sheets API service for connecting to the API.
        Dim spreadsheetId As String = "Spread sheet ID here"
		Dim oblist = New List(Of Object)() From {"Data to write to cell"}
		Dim valueRange As New ValueRange()
		valueRange.MajorDimension = "COLUMNS"
		valueRange.Values = New List(Of IList(Of Object))() From {oblist}
		Dim subSheetID as String = "Subsheet ID here"
		Dim range as String = subSheetID & "!A5"
		Dim writeCellRowRequest As SpreadsheetsResource.ValuesResource.AppendRequest = service.Spreadsheets.Values.Append(valueRange, spreadsheetId, range)
		If newRow = 1 Then 'For adding a new row if needed, sheets can only add a new row if the row it is adding at is an empty row, if not empty then it will add the new row to the next available empty row 
			writeCellRowRequest.InsertDataOption = SpreadsheetsResource.ValuesResource.AppendRequest.InsertDataOptionEnum.INSERTROWS
		End If 'If not adding new row then will overwrite cell
		writeCellRowRequest.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW
		writeCellRowRequest.Execute()
End Sub
