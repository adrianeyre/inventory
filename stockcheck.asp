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

DataAsset = Request.Form("AssetNumberText")
if DataAsset = "" then
	DataAsset = Request.QueryString("Asset")
end if

if DataAsset = "" then 
	ShowInput = true
else 
	ShowInput = false
	DataAsset = cint(DataAsset)
end if

DataAction = ""
DataAction = Request.QueryString("Query")

if lcase(DataAction) = "update" then
	ShowInput = true
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
	Command1.CommandText = "UPDATE dbo.MainTable SET Status='"&DataStatus&"',LastUpdated='"&Date&"',Description='"&DataDescription&"',Make='"&DataMake&"',Model='"&DataModel&"',Location='"&DataLocation&"',SerialNumber='"&DataSerialNumber&"',PurchaseDate='"&DataPurchaseDate&"',Warranty='"&DataWarranty&"',OrderNumber='"&DataOrderNumber&"',Supplier='"&DataSupplier&"',Picture='"&DataPicture&"', Cost='"&DataCost&"', Depreciation='"&DataDepreciation&"' WHERE Asset = '"&DataAsset&"'"
	Command1.Execute()
	
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
		RoomQuery.Source = "SELECT * FROM dbo.MainTable WHERE Asset LIKE '"&DataAsset&"'": Query = "Asset = "& AssetConvert(DataAsset)
	else
		RecordOn = Request.QueryString("Record")
		if RecordOn = "" then RecordOn = 1
		PreviousRecord = RecordOn - 1
		if PreviousRecord < 1 then PreviousRecord = 1
		NextRecord = RecordOn + 1
		RoomQuery.Source = DataQuery
	end if

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
if Query = "" then Query = "New Record : " & AssetConvert(Amount)
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
<%if lcase(DataAction) = "update" then %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backIndustria.png" bgcolor="#192F68"><div align="center" class="style1 style4">Record Updated!</div></td>
  </tr>
</table>
<% end if %>
<% if ShowInput = true then %>
<form id="MainForm" name="MainForm" method="post" action="?Query=">
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" align="center" valign="middle" background="../images/backcaritas.png" bgcolor="#192F68"><div align="center" class="style1 style4">Asset Number 
            <label for="AssetNumberText"></label>
            <input name="AssetNumberText" type="text" id="AssetNumberText" value="" autofocus/>
          </span>
               <input name="Submit" type="submit" class="style5" id="Submit" value="Submit" />
    </div></td>
  </tr>
</table>
<% else %>
<form id="MainForm" name="MainForm" method="post" action="<% response.write("?Query=update&asset="&DataAsset)%>">
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backcaritas.png" bgcolor="#192F68"><div align="center" class="style1 style4">Update?
      
<label for="AssetNumberText"></label>
          </span>
               <input name="Submit" type="submit" class="style5" id="Submit" value="Submit" />
    </div></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <table width="715" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="10">&nbsp;</td>
          <td width="150">&nbsp;</td>
          <td width="250">&nbsp;</td>
          <td width="305">&nbsp;</td>
        </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Asset Number</td>
          <td width="250" height="25" class="style5"><%response.write(AssetConvert(DataAsset))%></td>
          <td width="305" rowspan="13"><center><img src="Images/<%response.write(IntPicture)%>" width="280" height="280" alt="" /></center></td>
        </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Curretn Status</td>
          <td width="250" height="25"><%Response.write(StatusInfo(IntStatus))%> change to <select name="Status" id="Status">
            <option value="<%response.write("0")%>" selected="selected"><%response.write("Active")%></option>
            <%
				For a = 0 to StatusAmount%>
            <option value="<%response.write(a)%>">
              <%response.write(StatusInfo(a))%>
              </option>
            <%Next%>
          </select>
            </td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Location</td>
          <td width="250" height="25"><select name="Location" class="style5" id="Location">
            <option value="<%response.write(IntLocation)%>">
              <%response.write(LocationInfo(IntLocation))%>
              </option>
            <%
				For a = 0 to LocationAmount%>
            <option value="<%response.write(a)%>">
              <%response.write(LocationInfo(a))%>
              </option>
            <%Next%>
          </select></td>
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
          </select></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Make</td>
          <td width="250" height="25"><label for="Model"></label>
            <label for="checkbox4">
              <select name="Make" id="Make">
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
            </label></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Model</td>
          <td width="250" height="25"><input name="Model" type="text" id="Model" value="<%response.write(IntModel)%>" size="30" maxlength="50" /></td>
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
          </select></td>
          </tr>
        <tr>
          <td width="10" height="25">&nbsp;</td>
          <td width="150" height="25">Purcase Date</td>
          <td width="250" height="25"><input name="PurchaseDate" type="text" id="PurchaseDate" value="<%response.write(IntPurchaseDate)%>" size="10" maxlength="10" /></td>
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
					<img src="Images/NoWarranty.png" width="95" height="19" alt="No Warranty" />
                <% else %>
                	<img src="Images/InWarranty.png" width="95" height="20" alt="In Warranty" />
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
            <input name="Cost" type="text" id="Cost" value="<%response.write(IntCost)%>" size="10" maxlength="10" /></td>
          <td height="25">Depreciation Years
            <input name="Depreciation" type="text" id="Depreciation" value="<%if IntDepreciation = "" then response.write("4") else response.write(IntDepreciation)%>" size="3" maxlength="3" /></td>
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
          <td height="25" colspan="4"><center>
          </center></td>
          </tr>
        <tr>
          <td height="25" colspan="4" align="center"><label for="checkbox2"><a href="?" target="_self"><img src="../images/cancel.png" width="140" height="29" alt="Cancel" border="false"/></a></label></td>
          </tr>
      </table>
    </td>
  </tr>
</table>
<% end if %>
</form>
<% if DataQuery <> "" then %>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backindustria.png" bgcolor="#192F68"><div align="center" class="style1 style4">Total Net Book Value £
        <%response.write(FormatNumber(TrueValeTotal))%></div></td>
  </tr>
</table>
<br /><br />
<% end if 


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
