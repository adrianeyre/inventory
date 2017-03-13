<%
' ****************************************************
' *                  inventory.asp                   *
' *                                                  *
' *            Coded by : Adrian Eyre                *
' *                Date : 26/11/2013                 *
' *             Version : 1.1.0                      *
' *                                                  *
' ****************************************************
%>
<!--#include file="Connections/InventoryConnection.asp" -->
<%
Dim DescriptionQuery
Set DescriptionQuery = Server.CreateObject("ADODB.Recordset")
DescriptionQuery.ActiveConnection = MM_InventoryConnection_STRING
DescriptionQuery.Source = "SELECT * FROM dbo.DescriptionTable ORDER BY Description ASC"
DescriptionQuery.CursorType = 0
DescriptionQuery.CursorLocation = 2
DescriptionQuery.LockType = 1
DescriptionQuery.Open()

Dim MakeQuery
Set MakeQuery = Server.CreateObject("ADODB.Recordset")
MakeQuery.ActiveConnection = MM_InventoryConnection_STRING
MakeQuery.Source = "SELECT * FROM dbo.MakeTable ORDER BY Make ASC"
MakeQuery.CursorType = 0
MakeQuery.CursorLocation = 2
MakeQuery.LockType = 1
MakeQuery.Open()

Dim ModelQuery
Set ModelQuery = Server.CreateObject("ADODB.Recordset")
ModelQuery.ActiveConnection = MM_InventoryConnection_STRING
ModelQuery.Source = "SELECT DISTINCT Model FROM dbo.MainTable ORDER BY Model ASC"
ModelQuery.CursorType = 0
ModelQuery.CursorLocation = 2
ModelQuery.LockType = 1
ModelQuery.Open()

Dim SupplierQuery
Set SupplierQuery = Server.CreateObject("ADODB.Recordset")
SupplierQuery.ActiveConnection = MM_InventoryConnection_STRING
SupplierQuery.Source = "SELECT * FROM dbo.SupplierTable ORDER BY Supplier ASC"
SupplierQuery.CursorType = 0
SupplierQuery.CursorLocation = 2
SupplierQuery.LockType = 1
SupplierQuery.Open()

Dim OrderNumberQuery
Set OrderNumberQuery = Server.CreateObject("ADODB.Recordset")
OrderNumberQuery.ActiveConnection = MM_InventoryConnection_STRING
OrderNumberQuery.Source = "SELECT DISTINCT OrderNumber FROM dbo.MainTable ORDER BY OrderNumber ASC"
OrderNumberQuery.CursorType = 0
OrderNumberQuery.CursorLocation = 2
OrderNumberQuery.LockType = 1
OrderNumberQuery.Open()

Dim LocationQuery
Set LocationQuery = Server.CreateObject("ADODB.Recordset")
LocationQuery.ActiveConnection = MM_InventoryConnection_STRING
LocationQuery.Source = "SELECT * FROM dbo.LocationTable ORDER BY Location ASC"
LocationQuery.CursorType = 0
LocationQuery.CursorLocation = 2
LocationQuery.LockType = 1
LocationQuery.Open()

Dim PurchaseDateQuery
Set PurchaseDateQuery = Server.CreateObject("ADODB.Recordset")
PurchaseDateQuery.ActiveConnection = MM_InventoryConnection_STRING
PurchaseDateQuery.Source = "SELECT DISTINCT PurchaseDate FROM dbo.MainTable ORDER BY PurchaseDate ASC"
PurchaseDateQuery.CursorType = 0
PurchaseDateQuery.CursorLocation = 2
PurchaseDateQuery.LockType = 1
PurchaseDateQuery.Open()

Dim StatusQuery
Set StatusQuery = Server.CreateObject("ADODB.Recordset")
StatusQuery.ActiveConnection = MM_InventoryConnection_STRING
StatusQuery.Source = "SELECT * FROM dbo.StatusTable ORDER BY Status ASC"
StatusQuery.CursorType = 0
StatusQuery.CursorLocation = 2
StatusQuery.LockType = 1
StatusQuery.Open()


response.cookies("link")="index.asp>links.asp>staff\staff.asp>staff\inventory.asp"
response.cookies("linktext")="Home>Links>Staff Portal>Inventory"


%>

<style type="text/css">
<!--
.style1 {	font-size: x-large;
	color: #FFFFFF;
}
.style4 {font-size: x-large}
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
.style5 {font-size: large}
-->
</style>
<table width="715" height="45" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td width="716" height="45" background="../images/backdefault.png" bgcolor="#192F68"><div align="center" class="style1 style4">Inventory</div></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td bgcolor="#999999"><table width="715" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td colspan="3"><div align="center" class="style5">Search</div></td>
        </tr>
      <tr>
      <form id="form1" name="form1" method="post" action="showinv.asp">
        <td width="95" bgcolor="#FFFFFF">&nbsp;</td>
        <td width="134" bgcolor="#FFFFFF">Asset Tag</td>
        <td width="486" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><input type="text" name="Asset" id="Asset" /></td>
            <td width="50%"><label>
              <input name="Submit1" type="submit" id="Submit1" value="Submit" />
            </label>
            </td>
          </tr>
        </table></td>
        </form>
        </tr>
              <tr>
              <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">Description</td>
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><label>
              <select name="Description" id="Description">
                <option value="0" selected="selected">&lt; Select &gt;</option>
                <%
				While (NOT DescriptionQuery.EOF)%>
                	<option value="<%response.write(DescriptionQuery.Fields.Item("ID").Value)%>"><%response.write(DescriptionQuery.Fields.Item("Description").Value)%></option>
                	<%DescriptionQuery.MoveNext()
				Wend
				%>
              </select>
            </label></td>
            <td width="50%"><label>
              <input name="Submit2" type="submit" id="Submit2" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
        </tr>
              <tr>
              <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#FFFFFF">&nbsp;</td>
        <td bgcolor="#FFFFFF">Make</td>
        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><select name="Make" id="Make">
              <option value="0" selected="selected">&lt; Select &gt;</option>
              <%
				While (NOT MakeQuery.EOF)%>
              <option value="<%response.write(MakeQuery.Fields.Item("ID").Value)%>">
                <%response.write(MakeQuery.Fields.Item("Make").Value)%>
                </option>
              <%MakeQuery.MoveNext()
				Wend
				%>
            </select></td>
            <td width="50%"><label>
              <input name="Submit3" type="submit" id="Submit3" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
        </tr>
              <tr>
              <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">Model</td>
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><select name="Model" id="Model">
              <option value="0" selected="selected">&lt; Select &gt;</option>
              <%
				While (NOT ModelQuery.EOF)%>
              <option value="<%response.write(ModelQuery.Fields.Item("Model").Value)%>">
                <%response.write(ModelQuery.Fields.Item("Model").Value)%>
                </option>
              <%ModelQuery.MoveNext()
				Wend
				%>
            </select></td>
            <td width="50%"><label>
              <input name="Submit4" type="submit" id="Submit4" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
        </tr>
      <tr>
      	<form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#FFFFFF">&nbsp;</td>
        <td bgcolor="#FFFFFF">Supplier</td>
        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><select name="Supplier" id="Supplier">
              <option value="0" selected="selected">&lt; Select &gt;</option>
              <%
				While (NOT SupplierQuery.EOF)%>
              <option value="<%response.write(SupplierQuery.Fields.Item("ID").Value)%>">
                <%response.write(SupplierQuery.Fields.Item("Supplier").Value)%>
                </option>
              <%SupplierQuery.MoveNext()
				Wend
				%>
            </select></td>
            <td width="50%"><label>
              <input type="submit" name="Submit5" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
        </tr>
      <tr>
      <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">Order Number</td>
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><select name="OrderNumber" id="OrderNumber">
              <option value="0" selected="selected">&lt; Select &gt;</option>
              <%
				While (NOT OrderNumberQuery.EOF)%>
              <option value="<%response.write(OrderNumberQuery.Fields.Item("OrderNumber").Value)%>">
                <%response.write(OrderNumberQuery.Fields.Item("OrderNumber").Value)%>
                </option>
              <%OrderNumberQuery.MoveNext()
				Wend
				%>
            </select></td>
            <td width="50%"><label>
              <input type="submit" name="Submit6" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
      </tr>
      <tr>
      <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#FFFFFF">&nbsp;</td>
        <td bgcolor="#FFFFFF">Location</td>
        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><select name="Location" id="Location">
              <option value="0" selected="selected">&lt; Select &gt;</option>
              <%
				While (NOT LocationQuery.EOF)%>
              <option value="<%response.write(LocationQuery.Fields.Item("ID").Value)%>">
                <%response.write(LocationQuery.Fields.Item("Location").Value)%>
                </option>
              <%LocationQuery.MoveNext()
				Wend
				%>
            </select></td>
            <td width="50%"><label>
              <input type="submit" name="Submit7" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
      </tr>
      <tr>
      <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">Purchase Date</td>
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><select name="PurchaseDate" id="PurchaseDate">
              <option value="0" selected="selected">&lt; Select &gt;</option>
              <%
				While (NOT PurchaseDateQuery.EOF)%>
              <option value="<%response.write(PurchaseDateQuery.Fields.Item("PurchaseDate").Value)%>">
                <%response.write(PurchaseDateQuery.Fields.Item("PurchaseDate").Value)%>
                </option>
              <%PurchaseDateQuery.MoveNext()
				Wend
				%>
            </select></td>
            <td width="50%"><label>
              <input type="submit" name="Submit8" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
      </tr>
      <tr>
      <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#FFFFFF">&nbsp;</td>
        <td bgcolor="#FFFFFF">Status</td>
        <td bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><select name="Status" id="Status">
              <option value="0" selected="selected">&lt; Select &gt;</option>
              <%
				While (NOT StatusQuery.EOF)%>
              <option value="<%response.write(StatusQuery.Fields.Item("ID").Value)%>">
                <%response.write(StatusQuery.Fields.Item("Status").Value)%>
                </option>
              <%StatusQuery.MoveNext()
				Wend
				%>
            </select></td>
            <td width="50%"><label>
              <input type="submit" name="Submit9" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
      </tr>
      <tr>
      <form id="form1" name="form1" method="post" action="showinv.asp">
        <td bgcolor="#CCCCCC">&nbsp;</td>
        <td bgcolor="#CCCCCC">Serial Number</td>
        <td bgcolor="#CCCCCC"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%"><label for="SerialNumber"></label>
              <input type="text" name="SerialNumber" id="SerialNumber" /></td>
            <td width="50%"><label>
              <input type="submit" name="Submit10" value="Submit" />
            </label></td>
          </tr>
        </table></td>
        </form>
      </tr>
      <tr>
      <form id="form1" name="form1" method="post" action="showinv.asp">
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="50%">&nbsp;</td>
            <td width="50%"><label>
              <input type="submit" name="Submit11" value="Show All Records" />
            </label></td>
          </tr>
        </table></td>
        </form>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
      </tr>
    </table></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<table width="715" height="45" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="45" background="../images/backdefault.png" bgcolor="#192F68"><div align="center" class="style1 style4">School Map</div></td>
  </tr>
</table>
<table width="715" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><img src="../images/school.jpg" width="715" height="600" border="0" usemap="#Map"></td>
  </tr>
</table>


<map name="Map" id="Map">
  <area shape="rect" coords="107,49,125,77" href="showinv.asp?room=x1" target="_self" alt="Room: X1" />
<area shape="rect" coords="265,7,310,35" href="showinv.asp?room=x6" target="_self" alt="Room: X6" />
<area shape="rect" coords="310,6,356,37" href="showinv.asp?room=x7" target="_self" alt="Room: X7" />
<area shape="rect" coords="371,131,433,176" href="showinv.asp?room=pa1" target="_self" alt="Room: PA1" />
<area shape="rect" coords="371,79,431,120" href="showinv.asp?room=pa2" target="_self" alt="Room: PA2" />
<area shape="rect" coords="451,63,493,97" href="showinv.asp?room=pa2" target="_self" alt="Room: PA3" />
<area shape="rect" coords="449,96,492,133" href="showinv.asp?room=pa4" target="_self" alt="Room: PA4" />
<area shape="rect" coords="451,134,492,163" href="showinv.asp?room=pa5" target="_self" alt="Room: PA5" /><area shape="rect" coords="450,162,466,185" href="showinv.asp?room=pa6" target="_self" alt="Room: PA6" />
<area shape="rect" coords="382,10,429,37" href="showinv.asp?room=x8" target="_self" alt="Room: X8" />
<area shape="rect" coords="429,10,475,40" href="showinv.asp?room=x9" target="_self" alt="Room: X9" />
<area shape="rect" coords="499,10,542,35" href="showinv.asp?room=x10" target="_self" alt="Room: X10" />
<area shape="rect" coords="542,9,588,36" href="showinv.asp?room=x11" target="_self" alt="Room: X11" />
<area shape="rect" coords="612,9,660,37" href="showinv.asp?room=x12" target="_self" alt="Room: X12" />
<area shape="rect" coords="659,8,710,38" href="showinv.asp?room=x13" target="_self" alt="Room !3" />
<area shape="rect" coords="500,74,548,101" href="showinv.asp?room=x14" target="_self" alt="Room: X14" />
<area shape="rect" coords="548,73,592,102" href="showinv.asp?room=x15" target="_self" alt="Room: X15" />
<area shape="rect" coords="618,73,667,102" href="showinv.asp?room=x16" target="_self" alt="Room: X16" />
<area shape="rect" coords="667,73,711,102" href="showinv.asp?room=x17" target="_self" alt="Room: X17" />
<area shape="rect" coords="528,136,586,164" href="showinv.asp?room=x18" target="_self" alt="Room: X18" /><area shape="rect" coords="137,49,147,64" href="showinv.asp?room=x1 suite" target="_self" alt="Room: X1 Suite" />
<area shape="rect" coords="146,50,149,52" href="#" /><area shape="rect" coords="600,136,659,166" href="showinv.asp?room=x19" target="_self" alt="Room: X19" />
<area shape="rect" coords="147,49,175,76" href="showinv.asp?room=x2" target="_self" alt="Room: X2" />
<area shape="rect" coords="190,49,231,77" href="showinv.asp?room=x3" target="_self" alt="Room: X3" />
<area shape="rect" coords="243,59,289,87" href="showinv.asp?room=x4" target="_self" alt="Room: X4" />
<area shape="rect" coords="289,58,335,87" href="showinv.asp?room=x5" target="_self" alt="Room: X5" />
<area shape="rect" coords="175,49,184,76" href="showinv.asp?room=x2 office" target="_self" alt="Room: X2 Office" />
<area shape="rect" coords="124,49,139,68" href="showinv.asp?room=x1 office" target="_self" alt="Room: X1 Office" />
<area shape="rect" coords="276,120,303,147" href="showinv.asp?room=t1" target="_self" alt="Room: T1" />
<area shape="rect" coords="314,121,354,147" href="showinv.asp?room=a1" target="_self" alt="Room: A1" />
<area shape="rect" coords="313,147,351,173" href="showinv.asp?room=a2" target="_self" alt="Room: A2" />
<area shape="rect" coords="325,172,353,199" href="showinv.asp?room=n1" target="_self" alt="Room: N1" /><area shape="rect" coords="274,146,314,173" href="showinv.asp?room=a3" target="_self" alt="Room: A3" /><area shape="rect" coords="246,133,273,187" href="showinv.asp?room=t2" target="_self" alt="Room: T2" /><area shape="rect" coords="219,159,246,199" href="showinv.asp?room=t3" target="_self" alt="Room: T3" /><area shape="rect" coords="182,159,219,198" href="showinv.asp?room=t4" target="_self" alt="Room: T4" />
<area shape="rect" coords="102,291,182,332" href="showinv.asp?room=library" target="_self" alt="Room: Library" /><area shape="rect" coords="188,325,209,346" href="showinv.asp?room=sf1" target="_self" alt="Room: SF1" />
<area shape="rect" coords="209,326,221,346" href="showinv.asp?room=sf2" target="_self" alt="Room: SF2" />
<area shape="rect" coords="222,327,235,345" href="showinv.asp?room=sf3" target="_self" alt="Room: SF3" />
<area shape="rect" coords="236,327,257,345" href="showinv.asp?room=sf4" target="_self" alt="Room: SF4" />
<area shape="rect" coords="257,344,260,345" href="#" /><area shape="rect" coords="271,212,314,252" href="showinv.asp?room=d1" target="_self" alt="Room: D1" /><area shape="rect" coords="285,252,314,277" href="showinv.asp?room=c1" target="_self" alt="Room: C1" /><area shape="rect" coords="320,269,351,304" href="showinv.asp?room=r4" target="_self" alt="Room: R4" /><area shape="rect" coords="353,323,386,351" href="showinv.asp?room=r2" target="_self" alt="Room: R2" />
<area shape="rect" coords="321,318,352,350" href="showinv.asp?room=r3" target="_self" alt="Room: R3" /><area shape="rect" coords="351,370,385,402" href="showinv.asp?room=r1" target="_self" alt="Room: R1" />
<area shape="poly" coords="485,349,509,363,488,398,463,380" href="showinv.asp?room=s4" target="_self" alt="Room: S4" />
<area shape="poly" coords="527,393,506,379,485,413,506,425" href="showinv.asp?room=s3" target="_self" alt="Room: S3" />
<area shape="poly" coords="506,424,484,412,463,445,485,457" href="showinv.asp?room=s2" target="_self" alt="Room: S2" />
<area shape="poly" coords="486,458,464,493,442,476,463,446" href="showinv.asp?room=s1" target="_self" alt="Room: S1" />
<area shape="poly" coords="564,362,587,376,566,406,544,393" href="showinv.asp?room=s8" target="_self" alt="Room: S8" />
<area shape="poly" coords="534,409,557,423,534,456,513,441" href="showinv.asp?room=s9" target="_self" alt="Room: S9" />
<area shape="poly" coords="606,407,584,393,563,424,585,439" href="showinv.asp?room=s7" target="_self" alt="Room: S7" />
<area shape="circle" coords="349,442,24" href="showinv.asp?room=chapel" target="_self" alt="Room: Chapel" /><area shape="rect" coords="37,517,62,543" href="showinv.asp?room=m22" target="_self" alt="Room: M22" /><area shape="rect" coords="62,518,88,544" href="showinv.asp?room=m21" target="_self" alt="Room: M21" />
<area shape="rect" coords="38,498,51,517" href="showinv.asp?room=english office" target="_self" alt="Room: English Office" />
<area shape="rect" coords="37,475,63,499" href="showinv.asp?room=m23" target="_self" alt="Room: M23" />
<area shape="poly" coords="586,440,564,473,541,458,563,426" href="showinv.asp?room=s6" target="_self" alt="Room: S6" />
<area shape="poly" coords="564,473,542,505,523,490,542,460" href="showinv.asp?room=s5" target="_self" alt="Room: S5" />
<area shape="rect" coords="70,475,90,504" href="showinv.asp?room=m29" target="_self" alt="Room: M28" />
<area shape="rect" coords="38,450,63,476" href="showinv.asp?room=m24" target="_self" alt="Room: M24" />
<area shape="rect" coords="69,439,89,461" href="showinv.asp?room=m28" target="_self" alt="Room: M28" />
<area shape="rect" coords="73,411,91,441" href="showinv.asp?room=m27" target="_self" alt="Room: M27" />
<area shape="rect" coords="63,386,90,414" href="showinv.asp?room=m26" target="_self" alt="Room: M26" />
<area shape="rect" coords="39,392,63,426" href="showinv.asp?room=m25" target="_self" alt="Room: M25" />
<area shape="rect" coords="37,384,64,393" href="showinv.asp?room=language office" target="_self" alt="Room: Language Office" />
<area shape="rect" coords="163,551,210,580" href="showinv.asp?room=staff room" target="_self" alt="Room: Staff Room" />
<area shape="rect" coords="227,482,241,496" href="showinv.asp?room=meeting room" target="_self" alt="Room: Meeting Room" />
<area shape="poly" coords="10,297,9,354,54,354,52,312,31,314,30,298" href="showinv.asp?room=jpc" target="_self" alt="Room: JPC" />
<area shape="rect" coords="99,357,127,381" href="showinv.asp?room=m14" target="_self" alt="Room: M14" />
<area shape="rect" coords="101,381,127,404" href="showinv.asp?room=m13" target="_self" alt="Room: M13" />
<area shape="rect" coords="133,380,152,416" href="showinv.asp?room=m15" target="_self" alt="Room: M15" />
<area shape="rect" coords="125,427,152,450" href="showinv.asp?room=m11" target="_self" alt="Room: M11" /><area shape="rect" coords="100,427,125,450" href="showinv.asp?room=m12" target="_self" alt="Room: M12" />
<area shape="rect" coords="258,251,270,279" href="showinv.asp?room=resources" target="_self" alt="Room: Resources" />
<area shape="rect" coords="260,238,273,251" href="showinv.asp?room=lucas office" target="_self" alt="Room: Lucas Office" />
<area shape="rect" coords="164,538,181,552" href="showinv.asp?room=buet office" target="_self" alt="Room: Buet Office" /><area shape="rect" coords="166,523,181,538" href="showinv.asp?room=holts office" target="_self" alt="Room: Holts Office" />
<area shape="rect" coords="193,531,209,544" href="showinv.asp?room=data office" target="_self" alt="Room: Data Office" />
<area shape="rect" coords="195,519,212,531" href="showinv.asp?room=finance office" target="_self" alt="Room: Finance Office" /><area shape="rect" coords="166,481,186,507" href="showinv.asp?room=general office" target="_self" alt="Room: General Office" />
<area shape="rect" coords="166,507,186,521" href="showinv.asp?room=heads office" target="_self" alt="Room: Heads Office" />
<area shape="rect" coords="280,187,292,200" href="showinv.asp?room=exam office" target="_self" alt="Room: Exam Office" />
<area shape="rect" coords="164,468,212,483" href="showinv.asp?room=main corridor" target="_self" alt="Room: Main Corridor" />
<area shape="rect" coords="224,468,240,483" href="showinv.asp?room=reception" target="_self" alt="Room: Reception" />
<area shape="rect" coords="32,296,52,311" href="showinv.asp?room=heads of house office" target="_self" alt="Room: Heads of House Office" />
<area shape="rect" coords="258,326,270,346" href="showinv.asp?room=talmeys office" target="_self" alt="Room: Telmeys Office" />
<area shape="rect" coords="270,318,313,347" href="showinv.asp?room=6th form work room" target="_self" alt="Room: 6th Form Workroom" />
<area shape="rect" coords="92,47,108,62" href="showinv.asp?room=x1c" target="_self" alt="Room: X1C" />
<area shape="rect" coords="91,62,107,76" href="showinv.asp?room=x1b" target="_self" alt="Room: X1B" />
<area shape="rect" coords="449,198,471,211" href="showinv.asp?room=stantons office" target="_self" alt="Room: Stanton Office" />
<area shape="rect" coords="410,195,431,209" href="showinv.asp?room=brailsfords office" target="_self" alt="Room: Brailsfords Office" />
<area shape="poly" coords="446,407,463,381,486,398,469,424" href="showinv.asp?room=science prep room" target="_self" alt="Room: Science Prep Room" />
<area shape="rect" coords="138,345,155,363" href="showinv.asp?room=it support office" target="_self" alt="Room: IT Support Office" />
<area shape="rect" coords="141,338,154,345" href="showinv.asp?room=davie office" target="_self" alt="Room: Davie's Office" />
<area shape="rect" coords="167,338,182,345" href="showinv.asp?room=chaplin office" target="_self" alt="Room: Chaplin Office" />
<area shape="rect" coords="320,254,386,271" href="showinv.asp?room=tunnel" target="_self" alt="Room: Tunnel" />
<area shape="rect" coords="457,210,492,230" href="showinv.asp?room=pa6" target="_self" alt="Room: PA6" />
<area shape="rect" coords="91,96,181,147" href="showinv.asp?room=gym" target="_self" alt="Room: Gym" />
<area shape="rect" coords="128,223,175,278" href="showinv.asp?room=drama hall" target="_self" alt="Room: Drama Hall" />
<area shape="rect" coords="128,277,174,292" href="showinv.asp?room=av room" target="_self" alt="Room: AV Room" />
</map>
<%
DescriptionQuery.Close()
Set DescriptionQuery = Nothing

MakeQuery.Close()
Set MakeQuery = Nothing

ModelQuery.Close()
Set ModelQuery = Nothing

SupplierQuery.Close()
Set SupplierQuery = Nothing

OrderNumberQuery.Close()
Set OrderNumberQuery = Nothing

LocationQuery.Close()
Set LocationQuery = Nothing

PurchaseDateQuery.Close()
Set PurchaseDateQuery = Nothing

StatusQuery.Close()
Set StatusQuery = Nothing

%>
