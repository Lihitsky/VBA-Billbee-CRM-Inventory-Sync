Option Explicit

' Constants for API credentials
Private Const API_KEY As String = "YOUR API_KEY"
Private Const BILLBEE_USERNAME As String = "YOUR BILLBEE USERNAME"
Private Const BILLBEE_API_PASSWORD As String = "YOUR API PASSWORD"

' Function to fetch products from Billbee CRM
Public Function GetProducts(page As Long, pageSize As Long) As Collection
    ' Collection to hold all fetched data
    Dim allData As Collection
    Set allData = New Collection
    
    ' Loop to fetch all pages of data
    Do While True
        ' Construct the URL for the API request
        Dim url As String
        url = "https://api.billbee.io/api/v1/products?pageSize=" & pageSize & "&page=" & page
        
        ' Fetch the JSON response from the API
        Dim jsonResponse As String
        jsonResponse = GetAPIData(url, BILLBEE_USERNAME, BILLBEE_API_PASSWORD, API_KEY)

        ' Parse the JSON response into a dictionary
        Dim data As Dictionary
        Set data = JsonConverter.ParseJson(jsonResponse)
        
        ' Add the parsed data to the collection
        allData.Add data
        
        ' Check if there are more pages to fetch
        Dim totalPages As Long
        totalPages = CLng(data("Paging")("TotalPages"))
        
        If page >= totalPages Then
            Exit Do
        End If
        
        ' Move to the next page
        page = page + 1
    Loop
    
    ' Return the collected data
    Set GetProducts = allData
End Function
