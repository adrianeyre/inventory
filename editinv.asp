<%
' ****************************************************
' *                  editinv.asp                     *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 27/11/2013                 *
' *             Version : 1.0.0                      *
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

Function TrueCostValue(InitalCost, YearlyDepreciation, TruePurcaseDate, InitialStatus)
	Dim YearsOld
	Dim TrueCost
	
	if InitalCost = 0 or InitalCost = "" then
		TrueCostValue = FormatNumber(0,2)
	else
		TrueCost = 0
		YearsOld = DateDiff("yyyy",TruePurcaseDate,date)
		TrueCost = InitalCost - (YearlyDepreciation * YearsOld)
		if TrueCost < 0 then TrueCost = 0
		if TrueCost = "" then TrueCost = 0
		if InitialStatus = 0 or InitialStatus = 4 then 
			TrueCostValue = FormatNumber(TrueCost,2)
		else
			TrueCostValue = FormatNumber(0,2)
		end if
	end if
End Function

Function YearlyDepreciationValue(InitalCost, IntialDepreciation)
	YearlyDepreciation = cint(InitalCost) / cint(IntialDepreciation)
	YearlyDepreciationValue = FormatNumber(YearlyDepreciation,2)
End Function

DataAsset = Request.QueryString("Asset")
if DataAsset = "" then DataAsset = 1
DataLoan = Request.QueryString("Loan")
DataForname = Request.QueryString("Forname")
DataSurname = Request.QueryString("Surname")
DataAction = Request.QueryString("Action")
DataQuery = Request.Form("QueryData")
if DataQuery = "" then DataQuery = Request.QueryString("Query")

Dim TempAccount
Set TempAccount = Server.CreateObject("ADODB.Recordset")
TempAccount.ActiveConnection = MM_InventoryConnection_STRING
TempAccount.Source = "SELECT * FROM dbo.MainTable"
TempAccount.CursorType = 0
TempAccount.CursorLocation = 2
TempAccount.LockType = 1
TempAccount.Open()

NewAsset = 0
While (NOT TempAccount.EOF)
	NewAsset = NewAsset + 1
	TempAccount.MoveNext()
Wend
TempAccount.Close()
Set TempAccount = Nothing

if lcase(DataAction) = "update" or lcase(DataAction) = "new" then
	DataAmountOfRecords = Request.Form("AmountOfNewRecords")
	DataStatus = Request.Form("Status")
	DataDescription = Request.Form("Description")
	DataMake = Request.Form("Make")
	DataModel = Request.Form("Model")
	DataSerialNumber = Request.Form("SerialNumber")
	DataSupplier = Request.Form("Supplier")
	DataPurchaseDate = Request.Form("PurchaseDate")
	DataOrderNumber = Request.Form("OrderNumber")
	DataWarranty = Request.Form("Warranty")
	DataPicture = Request.Form("Picture")
	DataLocation = Request.Form("Location")
	if DataPicture = "" then DataPicture = "NoImage.png"
	DataCost = Request.Form("Cost")
	DataDepreciation = Request.Form("Depreciation")
	
	set Command1 = Server.CreateObject("ADODB.Command")
	Command1.ActiveConnection = MM_InventoryConnection_STRING
	Command1.CommandType = 1
	Command1.CommandTimeout = 0
	Command1.Prepared = true
	
	if lcase(DataAction) = "update" then
		Command1.CommandText = "UPDATE dbo.MainTable SET Status='"&DataStatus&"',LastUpdated='"&Date&"',Description='"&DataDescription&"',Make='"&DataMake&"',Model='"&DataModel&"',Location='"&DataLocation&"',SerialNumber='"&DataSerialNumber&"',PurchaseDate='"&DataPurchaseDate&"',Warranty='"&DataWarranty&"',OrderNumber='"&DataOrderNumber&"',Supplier='"&DataSupplier&"',Picture='"&DataPicture&"', Cost='"&DataCost&"', Depreciation='"&DataDepreciation&"' WHERE Asset = '"&DataAsset&"'"
		Command1.Execute()
	else
		NewAsset=NewAsset+1
		Command1.CommandText = "INSERT INTO dbo.MainTable (Asset, Status, LastUpdated, Description, Make, Model, Location, SerialNumber, PurchaseDate, Warranty, OrderNumber, Supplier, Picture, Cost, Depreciation) VALUES ('"&NewAsset&"','"&DataStatus&"','"&Date&"','"&DataDescription&"','"&DataMake&"','"&DataModel&"','"&DataLocation&"','"&DataSerialNumber&"','"&DataPurchaseDate&"','"&DataWarranty&"','"&DataOrderNumber&"','"&DataSupplier&"','"&DataPicture&"','"&DataCost&"','"&DataDepreciation&"')"
		Command1.Execute()
		DataAsset = NewAsset
	end if
end if


Dim DescriptionQuery
Dim DescriptionInfo(100)
Set DescriptionQuery = Server.CreateObject("ADODB.Recordset")
DescriptionQuery.ActiveConnection = MM_InventoryConnection_STRING
DescriptionQuery.Source = "SELECT * FROM dbo.DescriptionTable"
DescriptionQuery.CursorType = 0
DescriptionQuery.CursorLocation = 2
DescriptionQuery.LockType = 1
DescriptionQuery.Open()

DescriptionAmount = -1
While (NOT DescriptionQuery.EOF)
	DescriptionAmount = DescriptionAmount + 1
	DescriptionInfo(DescriptionAmount) = DescriptionQuery.Fields.Item("Description").Value
	DescriptionQuery.MoveNext()
Wend

Dim MakeQuery
Dim MakeInfo(100)
Set MakeQuery = Server.CreateObject("ADODB.Recordset")
MakeQuery.ActiveConnection = MM_InventoryConnection_STRING
MakeQuery.Source = "SELECT * FROM dbo.MakeTable"
MakeQuery.CursorType = 0
MakeQuery.CursorLocation = 2
MakeQuery.LockType = 1
MakeQuery.Open()

MakeAmount = -1
While (NOT MakeQuery.EOF)
	MakeAmount = MakeAmount + 1
	MakeInfo(MakeAmount) = MakeQuery.Fields.Item("Make").Value
	MakeQuery.MoveNext()
Wend

Dim SupplierQuery
Dim SupplierInfo(100)
Set SupplierQuery = Server.CreateObject("ADODB.Recordset")
SupplierQuery.ActiveConnection = MM_InventoryConnection_STRING
SupplierQuery.Source = "SELECT * FROM dbo.SupplierTable"
SupplierQuery.CursorType = 0
SupplierQuery.CursorLocation = 2
SupplierQuery.LockType = 1
SupplierQuery.Open()

SupplierAmount = -1
While (NOT SupplierQuery.EOF)
	SupplierAmount = SupplierAmount + 1
	SupplierInfo(SupplierAmount) = SupplierQuery.Fields.Item("Supplier").Value
	SupplierQuery.MoveNext()
Wend

Dim LocationQuery
Dim LocationInfo(500)
Set LocationQuery = Server.CreateObject("ADODB.Recordset")
LocationQuery.ActiveConnection = MM_InventoryConnection_STRING
LocationQuery.Source = "SELECT * FROM dbo.LocationTable"
LocationQuery.CursorType = 0
LocationQuery.CursorLocation = 2
LocationQuery.LockType = 1
LocationQuery.Open()

LocationAmount = -1
While (NOT LocationQuery.EOF)
	LocationAmount = LocationAmount + 1
	LocationInfo(LocationAmount) = LocationQuery.Fields.Item("Location").Value
	LocationQuery.MoveNext()
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

StatusAmount = -1
While (NOT StatusQuery.EOF)
	StatusAmount = StatusAmount + 1
	StatusInfo(StatusAmount) = StatusQuery.Fields.Item("Status").Value
	StatusQuery.MoveNext()
Wend

Dim RoomQuery
Dim RoomQuery_numRows
If DataAsset <> 0 then
	Set RoomQuery = Server.CreateObject("ADODB.Recordset")
	RoomQuery.ActiveConnection = MM_InventoryConnection_STRING
	if DataQuery = "" then
		RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Asset LIKE '"&DataAsset&"'"
		Query = "Asset = "& AssetConvert(DataAsset)
	else
		RecordOn = Request.QueryString("Record")
		if RecordOn = "" then RecordOn = 1
		PreviousRecord = RecordOn - 1
		if PreviousRecord < 1 then PreviousRecord = 1
		NextRecord = RecordOn + 1
		RoomQuery.Source = DataQuery
	end if
	'if DataDescription <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Description LIKE '"&DataDescription&"'": Query = "Description = "& DataDescription
	'if DataMake <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Make LIKE '"&DataMake&"'": Query = "Make = "& Make(DataMake)
	'if DataModel <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Model LIKE '"&DataModel&"'": Query = "Model = "& DataModel
	'if DataSupplier <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Supplier LIKE '"&DataSupplier&"'": Query = "Supplier = "& Supplier(DataSupplier)
	'if DataOrderNumber <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE OrderNumber LIKE '"&DataOrderNumber&"'": Query = "Order Number = "& DataOrderNumber
	'if DataLocation <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Location LIKE '"&DataLocation&"'": Query = "Location = "& Location(DataLocation)
	'if DataPurchaseDate <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE PurchaseDate LIKE '"&DataPurchaseDate&"'": Query = "Purchase Date = "& DataPurchaseDate
	'if DataStatus <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Status LIKE '"&DataStatus&"'": Query = "Status = "& StatusInfo(DataStatus)
	'if DataSerialNumber <> "" then RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE SerialNumber LIKE '%"&DataSerialNumber&"%'": Query = "Serial Number = "& DataSerialNumber
	'if RoomQuery.Source = "" then RoomQuery.Source  = "SELECT * FROM dbo.MainTable":Query = "Show all records"
	'RoomQuery.Source = "SELECT * FROM dbo.MainTable"
	RoomQuery.CursorType = 0
	RoomQuery.CursorLocation = 2
	RoomQuery.LockType = 1
	RoomQuery.Open()

	RoomQuery_numRows = 0
	RecordAmont = 0
	TrueValeTotal = 0
	
	if DataQuery <> "" then
		for a = 1 to (RecordOn - 1)
			RecordAmont = RecordAmont + 1
			IStatus = RoomQuery.Fields.Item("Status").Value
			ICost = RoomQuery.Fields.Item("Cost").Value
			IDepreciation = RoomQuery.Fields.Item("Depreciation").Value
			IPurchaseDate = RoomQuery.Fields.Item("PurchaseDate").Value
			if IPurchaseDate <> "" then
				IYearlyDepreciation = YearlyDepreciationValue(ICost, IDepreciation)
				ITrueCost = TrueCostValue(ICost, IYearlyDepreciation, IPurchaseDate, IStatus)
				TrueValeTotal = TrueValeTotal + ITrueCost
			end if
			RoomQuery.MoveNext()
		next
	end if
	
	IntPicture = ""
	IntAsset = RoomQuery.Fields.Item("Asset").Value
	IntStatus = RoomQuery.Fields.Item("Status").Value
	IntLastUpdated = RoomQuery.Fields.Item("LastUpdated").Value
	IntDescription = RoomQuery.Fields.Item("Description").Value
	IntMake = RoomQuery.Fields.Item("Make").Value
	IntModel = RoomQuery.Fields.Item("Model").Value
	IntLocation = RoomQuery.Fields.Item("Location").Value
	IntSerialNumber = RoomQuery.Fields.Item("SerialNumber").Value
	IntSupplier = RoomQuery.Fields.Item("Supplier").Value
	IntPurchaseDate = RoomQuery.Fields.Item("PurchaseDate").Value
	IntOrderNumber = RoomQuery.Fields.Item("OrderNumber").Value
	IntWarranty = RoomQuery.Fields.Item("Warranty").Value
	IntPicture = RoomQuery.Fields.Item("Picture").Value
	IntCost = RoomQuery.Fields.Item("Cost").Value
	IntDepreciation = RoomQuery.Fields.Item("Depreciation").Value
	
	if IntPurchaseDate <> "" then
		IntYearlyDepreciation = YearlyDepreciationValue(IntCost, IntDepreciation)
		IntTrueCost = TrueCostValue(IntCost, IntYearlyDepreciation, IntPurchaseDate, IntStatus)
		TrueValeTotal = TrueValeTotal + IntTrueCost
	end if

	if DataQuery <> "" then
		While (NOT RoomQuery.EOF)
			RecordAmont = RecordAmont + 1
			IStatus = RoomQuery.Fields.Item("Status").Value
			ICost = RoomQuery.Fields.Item("Cost").Value
			IDepreciation = RoomQuery.Fields.Item("Depreciation").Value
			IPurchaseDate = RoomQuery.Fields.Item("PurchaseDate").Value
			if IPurchaseDate <> "" then
				IYearlyDepreciation = YearlyDepreciationValue(ICost, IDepreciation)
				ITrueCost = TrueCostValue(ICost, IYearlyDepreciation, IPurchaseDate, IStatus)
				TrueValeTotal = TrueValeTotal + ITrueCost
			end if
			RoomQuery.MoveNext()
		Wend
		Query = "Asset = "& AssetConvert(IntAsset)
		DataAsset = IntAsset
		if NextRecord > RecordAmont then NextRecord = RecordAmont
	end if	
else
	Set RoomQuery = Server.CreateObject("ADODB.Recordset")
	RoomQuery.ActiveConnection = MM_InventoryConnection_STRING
	RoomQuery.Source = "SELECT * FROM dbo.MainTable"
	RoomQuery.CursorType = 0
	RoomQuery.CursorLocation = 2
	RoomQuery.LockType = 1
	RoomQuery.Open()

	RoomQuery_numRows = 0
	While (NOT RoomQuery.EOF)
		Amount = RoomQuery.Fields.Item("Asset").Value + 1
		RoomQuery.MoveNext()
	Wend
end if

if IntStatus = "" then IntStatus = 0
if IntLastUpdated = "" then IntLastUpdated = "Unknown"
if IntDescription = "" then IntDescription = 0
if IntMake = "" then IntMake = 0
if IntLocation = "" then IntLocation = 0
if IntSupplier = "" then IntSupplier = 0
if IntPurchaseDate = "" then IntPurchaseDate = date
if IntPicture = "" then IntPicture = "NoImage.png"
if Query = "" then Query = "New Record : " & AssetConvert(NewAsset+1)
if IntWarranty = "" then IntWarranty = 1

If lcase(DataLoan) = "return" then
	set Command1 = Server.CreateObject("ADODB.Command")
	Command1.ActiveConnection = MM_InventoryConnection_STRING
	'Command1.CommandText = "UPDATE dbo.LoanTable SET ReturnDate '" & Date & "' WHERE Asset='" &DataAsset&"' AND Forname='" &DataForname&"' AND Surname='" &DataSurname&"'")
	Command1.CommandText = "UPDATE dbo.LoanTable SET ReturnDate='" + Replace(Date, "'", "''") + "' WHERE Asset='" + Replace(DataAsset, "'", "''") + "' AND Forname='" + Replace(DataForname, "'", "''") + "' AND Surname='"+ Replace(DataSurname, "'", "''") + "'"
	Command1.CommandType = 1
	Command1.CommandTimeout = 0
	Command1.Prepared = true
	Command1.Execute()
end if

If lcase(DataLoan) = "new" then
	DataForname = Request.Form("Forname")
	DataSurname = Request.Form("Surname")
	DataLoanDate = Request.Form("LoanDate")
	if DataLoanDate = "" then DataLoanDate = Date
	if DataForname <> "" then
		set Command1 = Server.CreateObject("ADODB.Command")
		Command1.ActiveConnection = MM_InventoryConnection_STRING
		'Command1.CommandText = "UPDATE dbo.LoanTable SET ReturnDate '" & Date & "' WHERE Asset='" &DataAsset&"' AND Forname='" &DataForname&"' AND Surname='" &DataSurname&"'")
		Command1.CommandText = "INSERT INTO dbo.LoanTable (Asset, Forname, Surname, LoanDate)  VALUES ('" & DataAsset & "' ,'" & DataForname & "' , '" & DataSurname & "' , '" & DataLoanDate & "' ) "
		Command1.CommandType = 1
		Command1.CommandTimeout = 0
		Command1.Prepared = true
		Command1.Execute()
	else
		Response.Write("<script>alert('Please provide a name')</script>")
	end if
end if

Dim LoanQuery
Set LoanQuery = Server.CreateObject("ADODB.Recordset")
LoanQuery.ActiveConnection = MM_InventoryConnection_STRING
LoanQuery.Source = "SELECT * FROM dbo.LoanTable WHERE Asset LIKE '"&DataAsset&"' ORDER BY ID DESC"
LoanQuery.CursorType = 0
LoanQuery.CursorLocation = 2
LoanQuery.LockType = 1
LoanQuery.Open()

If DataAsset <> 0 then
	IntYearlyDepreciation = YearlyDepreciationValue(IntCost, IntDepreciation)
	IntTrueCost = TrueCostValue(IntCost, IntYearlyDepreciation, IntPurchaseDate, IntStatus)
end if
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
.style2 {
	font-size: medium;
	color: #FFFFFF;
}
.style4 {font-size: x-large}
.style5 {font-size: large}
-->
</style>
<a name="Top"></a>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backdefault.png" bgcolor="#192F68"><div align="center" class="style1 style4">Inventory</div></td>
  </tr>
</table>
<% if DataQuery <> "" then %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td height="40" align="center"><a href="editinv.asp?Record=1&Query=<%response.write(DataQuery)%>#Top"><img src="../images/firstbutton.png" width="140" height="29" alt="First Record" border="0" /></a></td>
        <td height="40" align="center"><a href="editinv.asp?Record=<%response.write(PreviousRecord)%>&Query=<%response.write(DataQuery)%>#Top"><img src="../images/previousbutton.png" width="140" height="29" alt="Previous Record" border="0" /></a></td>
        <td height="40" align="center"><a href="editinv.asp?Record=<%response.write(NextRecord)%>&Query=<%response.write(DataQuery)%>#Top"><img src="../images/nextbutton.png" width="140" height="29" alt="Next Record" border="0" /></a></td>
        <td height="40" align="center"><a href="editinv.asp?Record=<%response.write(RecordAmont)%>&Query=<%response.write(DataQuery)%>#Top"><img src="../images/lastbutton.png" width="140" height="29" alt="Last Record" border="0" /></a></td>
      </tr>
    </table></td>
  </tr>
</table>
<% end if %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="40"><center>
      <a href="adminpage.asp" target="_self"><img src="Images/AdminPage.png" width="204" height="29" alt="Admin Page" border="0" /></a>
    </center></td>
  </tr>
</table>
<% if DataQuery <> "" then %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backindustria.png" bgcolor="#192F68"><div align="center" class="style1 style4">Record <%response.write(RecordOn)%> / <%response.write(RecordAmont)%>
    </div></td>
  </tr>
</table>
<% end if %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backcaritas.png" bgcolor="#192F68"><div align="center" class="style1 style4"><span class="style5"><%=(Query)%> </span></div></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><form id="MainForm" name="MainForm" method="post" action="
    <% if DataAsset <> 0 then 
		response.write("editinv.asp?Asset="&DataAsset&"&Query="&DataQuery&"&Record="&RecordOn&"&Action=Update")
	   else	
	   	response.write("editinv.asp?Asset="&RecordAmont+1&"&Action=New")
	   end if
	%>">
      <table width="715" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="10">&nbsp;</td>
          <td width="150">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="305">&nbsp;</td>
        </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Asset</td>
          <td width="250" height="25">
          <label for="NewAsset"></label>
            <input name="NewAsset" type="text" id="NewAsset" value="<%if Amount="" then response.write(DataAsset) else response.write("NEW")%>" readonly="readonly" />
            </td>
          <td width="305" rowspan="13"><center><img src="Images/<%response.write(IntPicture)%>" width="280" height="280" alt="" /></center></td>
        </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Status</td>
          <td width="250" height="25"><select name="Status" id="Status">
            <option value="<%response.write(IntStatus)%>" selected="selected"><%response.write(StatusInfo(IntStatus))%></option>
            <%
				For a = 0 to StatusAmount%>
            <option value="<%response.write(a)%>">
              <%response.write(StatusInfo(a))%>
              </option>
            <%Next%>
          </select>
            <% if DataQuery <> "" then %><input type="checkbox" name="checkbox" id="checkbox" /><% end if %>
            <label for="checkbox"></label></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Last Updated</td>
          <td width="250" height="25"><%response.write(IntLastUpdated)%></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Description</td>
          <td width="250" height="25"><select name="Description" id="Description">
            <option value="<%response.write(IntDescription)%>" selected="selected">
              <%response.write(DescriptionInfo(IntDescription))%>
              </option>
            <%
				For a = 0 to DescriptionAmount%>
            <option value="<%response.write(a)%>">
              <%response.write(DescriptionInfo(a))%>
              </option>
            <%Next%>
          </select>
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox2" id="checkbox2" />
            <% end if %>
            <label for="checkbox2"></label></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Make</td>
          <td width="250" height="25"><select name="Make" id="Make">
            <option value="<%response.write(IntMake)%>">
              <%response.write(MakeInfo(IntMake))%>
              </option>
            <%
				For a = 0 to MakeAmount%>
            <option value="<%response.write(a)%>">
              <%response.write(MakeInfo(a))%>
              </option>
            <%Next%>
          </select>
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox3" id="checkbox3" />
            <% end if %>
            <label for="checkbox3"></label></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Model</td>
          <td width="250" height="25"><label for="Model"></label>
            <input name="Model" type="text" id="Model" value="<%response.write(IntModel)%>" size="30" maxlength="50" />
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox4" id="checkbox4" />
            <% end if %>
            <label for="checkbox4"></label></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Location</td>
          <td width="250" height="25"><select name="Location" id="Location">
            <option value="<%response.write(IntLocation)%>">
              <%response.write(LocationInfo(IntLocation))%>
              </option>
            <%
				For a = 0 to LocationAmount%>
            <option value="<%response.write(a)%>">
              <%response.write(LocationInfo(a))%>
              </option>
            <%Next%>
          </select>
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox5" id="checkbox5" />
            <% end if %>
            <label for="checkbox5"></label></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Serial Number</td>
          <td width="250" height="25"><input name="SerialNumber" type="text" id="SerialNumber" value="<%response.write(IntSerialNumber)%>" size="30" maxlength="50" /></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Supplier</td>
          <td width="250" height="25"><select name="Supplier" id="Supplier">
            <option value="<%response.write(IntSupplier)%>">
              <%response.write(SupplierInfo(IntSupplier))%>
              </option>
            <%
				For a = 0 to SupplierAmount%>
            <option value="<%response.write(a)%>">
              <%response.write(SupplierInfo(a))%>
              </option>
            <%Next%>
          </select>
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox7" id="checkbox7" />
            <% end if %>
            <label for="checkbox7"></label></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Purcase Date</td>
          <td width="250" height="25"><input name="PurchaseDate" type="text" id="PurchaseDate" value="<%response.write(IntPurchaseDate)%>" size="10" maxlength="10" />
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox6" id="checkbox6" />
            <% end if %>
            <label for="checkbox6"></label></td>
          </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Order Number</td>
          <td height="25"><input name="OrderNumber" type="text" id="OrderNumber" value="<%response.write(IntOrderNumber)%>" size="30" maxlength="50" /></td>
          </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Years of Warranty</td>
          <td height="25" valign="middle"><input name="Warranty" type="text" id="Warranty" value="<%response.write(IntWarranty)%>" size="2" maxlength="2" />
            <% if DataQuery <> "" then %>
              <input type="checkbox" name="checkbox8" id="checkbox8" />
              <% end if %>
              End
              <label for="checkbox8"></label>
              <% response.write(DateAdd("yyyy",IntWarranty,IntPurchaseDate))
				if DateAdd("yyyy",IntWarranty,IntPurchaseDate) < Date then %>
					<img src="Images/NoWarranty.png" width="102" height="20" alt="No Warranty" />
                <% else %>
                	<img src="Images/InWarranty.png" width="102" height="20" alt="In Warranty" />
                <% end if %>
			 </td>
          </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Image Name</td>
          <td height="25"><input name="Picture" type="text" id="Picture" value="<%response.write(IntPicture)%>" size="30" maxlength="50" />
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox9" id="checkbox9" />
            <% end if %>
            <label for="checkbox9"></label></td>
          </tr>

        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Original Cost</td>
          <td height="25">£
            <input name="Cost" type="text" id="Cost" value="<%response.write(IntCost)%>" size="10" maxlength="10" />
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox10" id="checkbox10" />
            <% end if %>
            <label for="checkbox10"></label></td>
          <td height="25">Depreciation Years
            <input name="Depreciation" type="text" id="Depreciation" value="<%if IntDepreciation = "" then response.write("4") else response.write(IntDepreciation)%>" size="3" maxlength="3" />
            <% if DataQuery <> "" then %>
            <input type="checkbox" name="checkbox11" id="checkbox11" />
            <% end if %>
            <label for="checkbox11"></label></td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Yearly Deduction</td>
          <td height="25">£<%response.write(IntYearlyDepreciation) %></td>
          <td height="25"><%response.write(DateDiff("yyyy",IntPurchaseDate,date)) %> 
          year(s) old</td>
        </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">Net Book Value</td>
          <td height="25">£<font color="red"><%response.write(IntTrueCost)%></font></td>
          <td height="25">&nbsp;</td>
        </tr>
        <tr>
          <td height="25" colspan="4"><center><% if DataAsset <> 0 then %><input type="submit" name="Submit2" id="Submit2" value="Update Record" /><% end if %>
            <% if DataAsset = 0 then %>
            Amount of new records
            <input name="AmountOfNewRecords" type="text" id="AmountOfNewRecords" value="1" size="4" maxlength="4" />
            <input type="submit" name="Submit3" id="Submit3" value="New Record" />
            <% end if %>
          </center></td>
          </tr>
        <tr>
          <td height="25">&nbsp;</td>
          <td height="25">&nbsp;</td>
          <td height="25">&nbsp;</td>
          <td height="25">&nbsp;</td>
        </tr>
      </table>
    </form></td>
  </tr>
</table>
<% if DataQuery <> "" then %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backindustria.png" bgcolor="#192F68"><div align="center" class="style1 style4">Total Net Book Value £
        <%response.write(FormatNumber(TrueValeTotal))%></div></td>
  </tr>
</table>
<br /><br />
<% end if %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backcaritas.png" bgcolor="#192F68"><div align="center" class="style1 style4"><span class="style5">Loan Information</span></div></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
    <form id="NewLoan" name="NewLoan" method="post" action="editinv.asp?Asset=<%response.write(DataAsset)%>&Loan=New">
  <table width="715" border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td width="10" bgcolor="#999999">&nbsp;</td>
      <td width="150" bgcolor="#999999">&nbsp;</td>
      <td width="250" bgcolor="#999999">&nbsp;</td>
      <td width="150" bgcolor="#999999">&nbsp;</td>
      <td width="152" bgcolor="#999999">&nbsp;</td>
    </tr>
    <tr>
      <td width="10" bgcolor="#999999">&nbsp;</td>
      <td width="150" bgcolor="#999999">First Name</td>
      <td width="250" bgcolor="#999999"><label for="Forname"></label>
      <input name="Forname" type="text" id="Forname" value="" size="30" maxlength="30" /></td>
      <td width="150" bgcolor="#999999">Loan Date</td>
      <td width="152" bgcolor="#999999"><input name="LoanDate" type="text" id="LoanDate" value="<%response.write(Date)%>" size="10" maxlength="10" /></td>
    </tr>
    <tr>
      <td width="10" bgcolor="#999999">&nbsp;</td>
      <td width="150" bgcolor="#999999">Surname</td>
      <td width="250" bgcolor="#999999"><input name="Surname" type="text" id="Surname" value="" size="30" maxlength="30" /></td>
      <td width="150" bgcolor="#999999">&nbsp;</td>
      <td width="152" bgcolor="#999999">&nbsp;</td>
    </tr>
    <tr>
      <td height="30" colspan="5" bgcolor="#999999"><center><input type="submit" name="Submit" id="Submit" value="Add Loan Information" /></center></td>
      </tr>
  </table>
</form>
    </td>
  </tr>
</table>
<%
While (NOT LoanQuery.EOF)
%>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="20" background="../images/backcaritas.png" bgcolor="#192F68"><div align="center" class="style1 style4"><span class="style5">Existing Loans</span></div></td>
  </tr>
</table>
  <table width="715" border="0" cellspacing="0" cellpadding="0">
  <form id="ReturnForm" name="ReturnForm" method="post" action="editinv.asp?Asset=<%response.write(DataAsset)%>&Loan=Return&Forname=<%=(LoanQuery.Fields.Item("Forname").Value)%>&Surname=<%=(LoanQuery.Fields.Item("Surname").Value)%>">
    <tr>
      <td width="10">&nbsp;</td>
      <td width="150">&nbsp;</td>
      <td width="250">&nbsp;</td>
      <td width="150">&nbsp;</td>
      <td width="152">&nbsp;</td>
    </tr>
    <tr>
      <td width="10">&nbsp;</td>
      <td width="150">First Name</td>
      <td width="250"><%=(LoanQuery.Fields.Item("Forname").Value)%></td>
      <td width="150">Loan Date</td>
      <td width="152"><%=(LoanQuery.Fields.Item("LoanDate").Value)%></td>
    </tr>
    <tr>
      <td width="10">&nbsp;</td>
      <td width="150">Surname</td>
      <td width="250"><%=(LoanQuery.Fields.Item("Surname").Value)%></td>
      <td width="150">Return Date</td>
      <td width="152"><%=(LoanQuery.Fields.Item("ReturnDate").Value)%></td>
    </tr>
    <tr>
      <td height="30" colspan="5"><center><%
	  	if len(LoanQuery.Fields.Item("ReturnDate").Value) = 10 then 
		else%> 
        	<input type="submit" name="Submit" id="Submit" value="Return Asset" />
		<% end if %>
      </center></td>
    </tr>
    </form>
  </table>
<%
	LoanQuery.MoveNext()
	
Wend

DescriptionQuery.Close()
Set DescriptionQuery = Nothing

MakeQuery.Close()
Set MakeQuery = Nothing

SupplierQuery.Close()
Set SupplierQuery = Nothing

LocationQuery.Close()
Set LocationQuery = Nothing

'RoomQuery.Close()
'Set RoomQuery = Nothing

StatusQuery.Close()
Set StatusQuery = Nothing

LoanQuery.Close()
Set LoanQuery = Nothing

%>
