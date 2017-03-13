<%
' ****************************************************
' *                  showinv.asp                     *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 26/11/2013                 *
' *             Version : 1.1.0                      *
' *                                                  *
' ****************************************************
%>
<!--#include file="Connections/InventoryConnection.asp" -->
<%
Function  AssetConvert(a)
	AssetConvert = "A"
	For b = 1 to (6 - len(a))
		AssetConvert = AssetConvert + "0"
	Next
	AssetConvert = AssetConvert + cstr(a)
End Function

Dim arrUser, Username
arrUser = Split(Request.ServerVariables("LOGON_USER"), "\")
Username= arrUser(1)

Dim AdminQuery
Set AdminQuery = Server.CreateObject("ADODB.Recordset")
AdminQuery.ActiveConnection = MM_InventoryConnection_STRING
AdminQuery.Source = "SELECT * FROM dbo.AdminTable"
AdminQuery.CursorType = 0
AdminQuery.CursorLocation = 2
AdminQuery.LockType = 1
AdminQuery.Open()

AdminUser = false
While (NOT AdminQuery.EOF)
	UName = AdminQuery.Fields.Item("AdminUsername").Value
	if lcase(UName) = lcase(Username) then AdminUser = true
	AdminQuery.MoveNext()
Wend

DataRoom = Request.QueryString("Room")
DataAsset = Request.Form("Asset")
DataAsset = cint(DataAsset)
DataDescription = Request.Form("Description")
DataMake = Request.Form("Make")
DataModel = Request.Form("Model")
DataSupplier = Request.Form("Supplier")
DataOrderNumber = Request.Form("OrderNumber")
DataLocation = Request.Form("Location")
DataPurchaseDate = Request.Form("PurchaseDate")
DataStatus = Request.Form("Status")
DataSerialNumber = Request.Form("SerialNumber")

Dim LocationQuery
Dim Location(5000)
Set LocationQuery = Server.CreateObject("ADODB.Recordset")
LocationQuery.ActiveConnection = MM_InventoryConnection_STRING
LocationQuery.Source = "SELECT * FROM dbo.LocationTable"
LocationQuery.CursorType = 0
LocationQuery.CursorLocation = 2
LocationQuery.LockType = 1
LocationQuery.Open()

While (NOT LocationQuery.EOF)
	Location(LocationQuery.Fields.Item("ID").Value) = LocationQuery.Fields.Item("Location").Value
	if lcase(DataRoom) = lcase(LocationQuery.Fields.Item("Location").Value) then
		DataLocation = LocationQuery.Fields.Item("ID").Value
	end if
	LocationQuery.MoveNext()
Wend

Dim MakeQuery
Dim Make(5000)
Set MakeQuery = Server.CreateObject("ADODB.Recordset")
MakeQuery.ActiveConnection = MM_InventoryConnection_STRING
MakeQuery.Source = "SELECT * FROM dbo.MakeTable"
MakeQuery.CursorType = 0
MakeQuery.CursorLocation = 2
MakeQuery.LockType = 1
MakeQuery.Open()

While (NOT MakeQuery.EOF)
	Make(MakeQuery.Fields.Item("ID").Value) = MakeQuery.Fields.Item("Make").Value
	MakeQuery.MoveNext()
Wend

Dim SupplierQuery
Dim Supplier(5000)
Set SupplierQuery = Server.CreateObject("ADODB.Recordset")
SupplierQuery.ActiveConnection = MM_InventoryConnection_STRING
SupplierQuery.Source = "SELECT * FROM dbo.SupplierTable"
SupplierQuery.CursorType = 0
SupplierQuery.CursorLocation = 2
SupplierQuery.LockType = 1
SupplierQuery.Open()

While (NOT SupplierQuery.EOF)
	Supplier(SupplierQuery.Fields.Item("ID").Value) = SupplierQuery.Fields.Item("Supplier").Value
	SupplierQuery.MoveNext()
Wend

Dim StatusQuery
Dim StatusInfo(10)
Set StatusQuery = Server.CreateObject("ADODB.Recordset")
StatusQuery.ActiveConnection = MM_InventoryConnection_STRING
StatusQuery.Source = "SELECT * FROM dbo.StatusTable"
StatusQuery.CursorType = 0
StatusQuery.CursorLocation = 2
StatusQuery.LockType = 1
StatusQuery.Open()

While (NOT StatusQuery.EOF)
	StatusInfo(StatusQuery.Fields.Item("ID").Value) = StatusQuery.Fields.Item("Status").Value
	StatusQuery.MoveNext()
Wend

Dim RoomQuery
Dim RoomQuery_numRows

Set RoomQuery = Server.CreateObject("ADODB.Recordset")
RoomQuery.ActiveConnection = MM_InventoryConnection_STRING
if DataAsset <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Asset LIKE '"&DataAsset&"'": Query = "Asset = "& DataAsset
if DataDescription <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Description LIKE '"&DataDescription&"'": Query = "Description = "& DataDescription
if DataMake <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Make LIKE '"&DataMake&"'": Query = "Make = "& Make(DataMake)
if DataModel <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Model LIKE '"&DataModel&"'": Query = "Model = "& DataModel
if DataSupplier <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Supplier LIKE '"&DataSupplier&"'": Query = "Supplier = "& Supplier(DataSupplier)
if DataOrderNumber <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE OrderNumber LIKE '"&DataOrderNumber&"'": Query = "Order Number = "& DataOrderNumber
if DataLocation <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Location LIKE '"&DataLocation&"'": Query = "Location = "& Location(DataLocation)
if DataPurchaseDate <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE PurchaseDate LIKE '"&DataPurchaseDate&"'": Query = "Purchase Date = "& DataPurchaseDate
if DataStatus <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Status LIKE '"&DataStatus&"'": Query = "Status = "& StatusInfo(DataStatus)
if DataSerialNumber <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE SerialNumber LIKE '%"&DataSerialNumber&"%'": Query = "Serial Number = "& DataSerialNumber
if RoomQuery.Source = "" then RoomQuery.Source  = "SELECT * FROM dbo.MainTable":Query = "Show all records"
QueryStrings = RoomQuery.Source 
'RoomQuery.Source = "SELECT * FROM dbo.MainTable"
RoomQuery.CursorType = 0
RoomQuery.CursorLocation = 2
RoomQuery.LockType = 1
RoomQuery.Open()

RoomQuery_numRows = 0

%>

<style type="text/css">
<!--
body,td,th {
	font-family: Arial, Helvetica, sans-serif;
	font-size: 14px;
}
body {
	margin-left: 0px;
	margin-top: 5px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.style1 {font-size: x-large;
	color: #FFFFFF;
}
.style4 {font-size: x-large}
.style5 {font-size: large}
-->
</style>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backdefault.png" bgcolor="#192F68"><div align="center" class="style1 style4">Inventory</div></td>
  </tr>
</table>
<% If AdminUser = true then %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td><div align="center">
      <form action="editinv.asp?Action=Multiple#Top" method="post" name="AllRecords" id="AllRecords">
        <label>
          <input type="submit" name="Submit" value="Show all result details" />
        </label>
        <input type="hidden" name="QueryData" id="QueryData" value="<%response.write(QueryStrings)%>" size="100" />
      </form>
    </div></td>
  </tr>
  <tr>
    <td>
    </td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><form id="form1" name="form1" method="post" action="editinv.asp?Asset=0">
      <center><input type="submit" name="Submit2" value="New Record" /></center>
    </form></td>
  </tr>
</table>
<% end if %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backcaritas.png" bgcolor="#192F68"><div align="center" class="style1 style4"><span class="style5">Query : <%response.write(Query)%> </span></div></td>
  </tr>
</table>
<table width="715" border="1" cellspacing="0" cellpadding="0" <% if num = 1 then response.write("bgcolor='#999999'") else response.write("bgcolor='#CCCCCC'")%>>
  <tr>
    <td width="15">&nbsp;</td>
    <td width="60">Asset</td>
    <td width="125">Location</td>
    <td width="132">Make</td>
    <td width="132">Model</td>
  </tr>
</table>
  <%
num = 0
Amount = 0
While (NOT RoomQuery.EOF)
	if num = 0 then num = 1 else num = 0
	Amount = Amount + 1
%>
<table width="715" border="1" cellspacing="0" cellpadding="0" <% if num = 1 then response.write("bgcolor='#999999'") else response.write("bgcolor='#CCCCCC'")%>>
  <tr>
    <td width="15" height="18"><center><img src="
    	<% if RoomQuery.Fields.Item("Status").Value = 0 then response.write("Images/StatusYES.png") else response.write("Images/StatusNO.png")%>
        " width="12" height="12" />
    </center></td>
    <td width="60">
    	<% if AdminUser = true then %>
    <a href="editinv.asp?Asset=<%=(RoomQuery.Fields.Item("Asset").Value)%>#Top" target="_self"><%=(AssetConvert(RoomQuery.Fields.Item("Asset").Value))%></a></td>
    <% else
        	response.write(AssetConvert(RoomQuery.Fields.Item("Asset").Value)) %>
        <% end if %>
    <td width="125"><%=(Location(RoomQuery.Fields.Item("Location").Value))%></td>
    <td width="132"><%=(Make(RoomQuery.Fields.Item("Make").Value))%></td>
    <td width="132"><%=(RoomQuery.Fields.Item("Model").Value)%></td>
  </tr>
</table>
<%
RoomQuery.MoveNext()
wend

RoomQuery.Close()
Set RoomQuery = Nothing
%>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backdefault.png" bgcolor="#192F68"><div align="center" class="style1 style4"><span class="style5">Amount of records <%response.write(Amount)%></span></div></td>
  </tr>
</table>
<label for="Query"></label>
<%
LocationQuery.Close()
Set LocationQuery = Nothing

MakeQuery.Close()
Set MakeQuery = Nothing

SupplierQuery.Close()
Set SupplierQuery = Nothing

StatusQuery.Close()
Set StatusQuery = Nothing



%>