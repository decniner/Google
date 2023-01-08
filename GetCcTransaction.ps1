# Install required modules
Install-Module -Name Google.Apis.Sheets.v4 -Scope CurrentUser
Install-Module -Name Google.Apis.Auth -Scope CurrentUser
Install-Module -Name Newtonsoft.Json -Scope CurrentUser

# Authenticate using the OAuth client ID
$credentialPath = "path\to\client_id.json"
$credential = New-Object Google.Apis.Auth.OAuth2.GoogleCredential(
    (New-Object Google.Apis.Auth.OAuth2.ServiceAccountCredential).FromJson($credentialPath).ToString()
)

# Send a request to https://www.eposcard.co.jp/memberservice/pc/webservicetop/web_service_top_preload.do to retrieve credit card transactions
$transactions = Invoke-WebRequest -Uri "https://www.eposcard.co.jp/memberservice/pc/webservicetop/web_service_top_preload.do" -Method Get -Headers @{
    Authorization = "Bearer $($credential.Token.AccessToken)"
}

# Write the retrieved data to a Google Sheets document
$service = New-Object Google.Apis.Sheets.v4.SheetsService
$service.HttpClient.DefaultRequestHeaders.Authorization =
    New-Object System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", $credential.Token.AccessToken)

$spreadsheetId = "spreadsheet_id"
$range = "Sheet1!A1"
$requestBody = New-Object Google.Apis.Sheets.v4.Data.ValueRange
$requestBody.Values = $transactions

$request = $service.Spreadsheets.Values.Update($requestBody, $spreadsheetId, $range)
$request.ValueInputOption = "USER_ENTERED"
$response = $request.Execute()
