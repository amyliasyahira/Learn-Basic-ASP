<%@ language="javascript" %>
<!DOCTYPE html>
<html>
<head><title>View a record</title>
</head>
<body>

<%
	var id, conn, rs, sql, urlid;
	
	conn = Server.CreateObject("ADODB.Connection");
	conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+Server.MapPath("testdb.mdb"));
	rs = Server.CreateObject("ADODB.Recordset");
	
	urlid = Request.ServerVariables("QUERY_STRING")+"";
	id = urlid.substring(urlid.indexOf("=")+1);
	//id = Request.QueryString("id");
	
	sql = "SELECT * FROM PTj WHERE id = "+id;
	rs.Open(sql, conn);
	
	if(Request.Form("btnDel")=="Delete")
	{
		sql = "DELETE FROM PTj WHERE id = "+id;
		conn.Execute(sql);
		Response.Redirect("view.asp");
		conn.Close;
		rs.Close;
	}

%>

<center><h1>VIEW A RECORD</h1></center>
<form name="frm" action="" method="post">
	<table align="center" border="0" width="50%" bgcolor="#AED6F1">
		<td><b>Name</b></td><td><b>Edit</b></td>
		<% if(!rs.EOF){ %>
		<tr>
		   <td><% = rs.Fields("name") %></td>
		   <td><a href="update.asp?id=<% = rs.Fields('id')%>">Edit</a></td>
		   <td><input type="submit" name="btnDel" value="Delete" onclick="return confirm('Are you sure you want to Delete this record ?')"; /></td>
		</tr>
		<tr><td colspan="4" align="right"><a href="view.asp">Menu</a></td></tr>
		<% 
		   rs.Close;
		   conn.Close;
		} %>
	</table>
</form>

</body>
</html>