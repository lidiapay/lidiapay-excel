# lidiapay-excel
Integrating the LidiaPay app with MS-Excel

![LidiaPay Excel](https://i.ibb.co/7tmCqSZ/excel-lidia.png)

The code can be found in VBA Macro editor.
![LidiaPay Excel Developer Tab](https://i.ibb.co/hf1cn3V/excel-lidia-developer-tab.png)

All actions require API authentication. Enter your account and password before any action.
![LidiaPay Excel API Authentication](https://i.ibb.co/YkWftwN/excel-lidia-authentication.png)

Authentication calls the LidiaPay API authentication method.
[LidiaPay Authentication API Details](https://api.lidia.co.in/#section_authentication)

```
Private Function GetAuthToken() As String
    Worksheets("Transactions").Range(CELL_STATUS).Value = "Authenticating..."
    Worksheets("Transactions").Range(CELL_RESULT_TEST).Value = ""
    
    Dim oHttp As Object
    Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    Dim account As String
    account = Worksheets("Transactions").Range(CELL_ACCOUNT).Value
    account = Trim(Replace(account, "-", ""))

    Dim Password As String
    Password = Worksheets("Transactions").txtPassword.Value

    Dim Body As String
    Body = "{""account"": """ & account & """,""password"": """ & Password & """}"

    oHttp.Open "POST", AUTH_URL, False
    oHttp.setRequestHeader "Content-type", "application/json"
    oHttp.setRequestHeader "Cache-Control", "no-cache"
    oHttp.setRequestHeader "cache-control", "no-cache"
    oHttp.send (Body)

    Dim token As String
    
    If oHttp.Status = 200 Then
        sJSON = oHttp.responseText
        Set oHttp = Nothing
    
        Dim Json As Object
        Set Json = JsonConverter.ParseJson(sJSON)
        
    
        If Not IsNull(Json("token")) Then
            token = Json("token")
        End If
    
        Worksheets("Transactions").Range(CELL_STATUS).Value = "Ready"
            
    End If
        
    If Len(token) = 0 Then
        Worksheets("Transactions").Range(CELL_RESULT_TEST).Value = "Authentication Failed"
    End If
        
    GetAuthToken = token
    
End Function
```

Load/Reload Transaction List action show the list of transactions for authenticated account. The Status cell informs the current status of the connection.
![LidiaPay Excel Transaction List](https://i.ibb.co/4d98J04/excel-lidia-load-tx.png)

```
Sub btnLoadTransactions_Click()
    Dim token As String
    token = GetAuthToken()
    
    If Len(token) > 0 Then
        ListTx token
    End If
    
End Sub
```
Transaction list is retrieved by API call
[LidiaPay Transactions API Details](https://api.lidia.co.in/#section_transactions)

```
Sub ListTx(token As String)
    Worksheets("Transactions").Range(CELL_STATUS).Value = "Loading list..."
    Worksheets("Transactions").Range(CELL_RESULT_TEST).Value = ""
    
    Dim oHttp As Object
    Set oHttp = CreateObject("WinHttp.WinHttpRequest.5.1")

    oHttp.Open "GET", LIST_TX_URL, False
    oHttp.setRequestHeader "Authorization", "Bearer " & token
    oHttp.setRequestHeader "Content-type", "application/json"
    oHttp.setRequestHeader "Cache-Control", "no-cache"
    oHttp.setRequestHeader "cache-control", "no-cache"
    oHttp.send (Body)

    sJSON = oHttp.responseText
    Set oHttp = Nothing

    Dim Json As Object
    Set Json = JsonConverter.ParseJson(sJSON)
    
    Dim Transactions As New Dictionary
    Set Transactions = Json("transactions")

    Dim Values As Variant
    Dim TotalFieldsInJSON As Integer
    TotalFieldsInJSON = 17
    
    ReDim Values(Transactions("result").Count, TotalFieldsInJSON)
    
    Dim Value As Dictionary
    Dim i As Long
    
    i = 0
    For Each Value In Transactions("result")
      Values(i, 0) = Value("txid")
      Values(i, 1) = Value("phoneid")
      Values(i, 2) = Value("coin")
      Values(i, 3) = Value("totalincoin")
      Values(i, 4) = Value("totalinfiat")
      Values(i, 5) = Value("fiatcurrency")
      Values(i, 6) = Value("coinvalueinfiat")
      Values(i, 7) = Value("minerfee")
      Values(i, 8) = Value("txidblockchain")
      Values(i, 9) = Value("status")
      Values(i, 10) = Value("paymentdate")
      Values(i, 11) = Value("transactiondate")
      Values(i, 12) = Value("demolive")
      Values(i, 13) = Value("accountpayer")
      Values(i, 14) = Value("administration_fee")
      Values(i, 15) = Value("administration_fee_paid")
      Values(i, 16) = Value("phoneidpayer")
      
      i = i + 1
    Next Value

    Worksheets("Transactions").Range(Cells(8, 1), Cells(Transactions("result").Count, TotalFieldsInJSON)) = Values

    'Worksheets("Transactions").Range(CELL_RESULT_TEST).Value = ""
    Worksheets("Transactions").Range(CELL_STATUS).Value = "Ready"
End Sub
```

To pay for a transaction, you need to tell the cashier (receiver's account) and the transaction ID.
![LidiaPay Excel Pay Transaction](https://i.ibb.co/wNqbTNY/excel-lidia-pay-typing.png)

After filling in the payment fields, click the Pay button and wait for Blockchain's payment ID to be shown.
![LidiaPay Excel Paid Transaction](https://i.ibb.co/sFB9dMD/excel-lidia-paid.png)

The payment code can be found inside the click of the Pay button.
```
Sub btnPay_Click()
    Dim token As String
    token = GetAuthToken()
    
    If Len(token) > 0 Then
        Pay token
    End If
End Sub
```
The Pay function calls the Send Payment method of the LidiaPay API.
[LidiaPay Send Payment API Details](https://api.lidia.co.in/#section_send_payment)

