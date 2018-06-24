<%@ language="javascript" %>
<!DOCTYPE html>
<html>
<body>

<%
	var conn, rs, sql, i=0;
	
	conn = Server.CreateObject("ADODB.Connection");
	conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+Server.MapPath("testdb.mdb"));
	rs = Server.CreateObject("ADODB.Recordset");
	
	sql = "SELECT * FROM PTj";
	rs.Open(sql, conn);


	if(Request.Form("view") == "View")
	{
		var id = Request.Form("ptj");
		Response.Write("<script> window.location='viewone.asp?id="+id+"'; </script>");
	}
%>


<center><h1>MENU</h1></center>


<table align="center" border="2" width="20%">
	<tr><td colspan="4" align="right"><a href="addUserLab4.asp">New Record</a></td></tr>
	<th>Pusat Tanggungjawab</th><th>Action</th>
	<form name="action" method="post" action="">
	<tr>
	   <td><select name="ptj">
	   			<% while(!rs.EOF){ %>
	   				<option value="<%= rs.Fields('id')%>"><%= rs.Fields('name')%></option>
	   			<% rs.MoveNext;}	
	   				rs.Close;
	   				conn.Close;	
				%>	
	   		</select>
	   	</td>
	   	<td>   		
	   		<input type="submit" name="view" value="View">   		
	   	</td>
	</tr>
	</form>
</table>
</body>
</html>