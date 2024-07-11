Option Explicit

' Function to perform API request and return the response as a string
Public Function GetAPIData(url As String, username As String, password As String, apiKey As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    ' Create basic authentication header
    Dim auth As String
    auth = "Basic " & EncodeBase64(username & ":" & password)
    
    ' Configure the HTTP request
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", auth
    http.setRequestHeader "X-Billbee-Api-Key", apiKey
    http.setRequestHeader "Content-Type", "application/json"
    
    ' Send the HTTP request
    http.Send
    
    ' Check if the response is successful
    If http.Status = 200 Then
        GetAPIData = http.ResponseText
    Else
        Err.Raise vbObjectError + 1, "GetAPIData", "Something wrong with API request: " & http.Status
    End If
End Function

' Function to encode text in Base64 format
Public Function EncodeBase64(text As String) As String
    ' Convert the text to a byte array
    Dim arrData() As Byte
    arrData = StrConv(text, vbFromUnicode)
    
    ' Create XML objects to encode the byte array
    Dim objXML As Object
    Dim objNode As Object
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arrData
    
    ' Get the Base64 encoded text
    EncodeBase64 = objNode.text
    
    ' Clean up objects
    Set objNode = Nothing
    Set objXML = Nothing
End Function
