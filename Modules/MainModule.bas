Option Explicit

' Main subroutine to fetch data from Billbee CRM and insert it into the Excel worksheet
Sub Main()
    ' Create instances of CRMClient and DataProcessor classes
    Dim CRMClient As New CRMClient
    Dim DataProcessor As New DataProcessor
    
    ' Set initial page and page size for data fetching
    Dim page As Long
    Dim pageSize As Long
    page = 1
    pageSize = 250
    
    ' Fetch all data from Billbee CRM
    Dim allData As Collection
    Set allData = CRMClient.GetProducts(page, pageSize)
    
    ' Insert the fetched data into the worksheet
    DataProcessor.InsertData allData
End Sub
