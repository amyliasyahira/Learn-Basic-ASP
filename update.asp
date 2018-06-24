<%@ language="javascript" %>
<!DOCTYPE html>
<html>
<head>
	<title>Update</title>
	
	<script>
		function setFocus(f)
		{
			f.elements["name"].focus();
		}
	</script>
	
	<%
		var id, conn, rs, sql, urlid, n, d;
	
		conn = Server.CreateObject("ADODB.Connection");
		conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+Server.MapPath("testdb.mdb"));
		rs = Server.CreateObject("ADODB.Recordset");
	
		urlid = Request.ServerVariables("QUERY_STRING")+"";
		id = urlid.substring(urlid.indexOf("=")+1);
	
		sql = "SELECT * FROM PTj WHERE id = "+id;
		rs.Open(sql, conn);
		
		if(Request.Form("btnEdit")=="Edit")
		{
			n = Request.Form("name");
			d = Request.Form("desc");
			
			//Response.Write(n+d);

			sql = "UPDATE PTj SET name ='"+n+"',descp='"+d+"' WHERE id = "+id;
			conn.Execute(sql);
			
			Response.Write("<script> alert('The record has been edited ...');window.location='view.asp'; </script>");
			
			conn.Close;
			rs.Close;
		}
	%>
	
</head>
<body>
	
	<center><h1>.::UPDATE::.</h1></center>
	<% if(!rs.EOF){ %>
	<form name="frm" action="" method="post">
		<table width="40%" border="1" align="center" cellpadding="10">
			<tr>
				<td style="width:20%">ID</td><td><input type="text" name="id" id="id" value="<% = rs.Fields('id') %>" maxlength="10" size="10" disabled /></td>
			</tr>
			<tr>
				<td>Name</td><td><input type="text" name="name" id="name" value="<% = rs.Fields('name') %>"  maxlength="10" size="10" autofocus /></td>
			</tr>
			<tr>
				<td>Description</td><td><input type="text" name="desc" id="desc" value="<% = rs.Fields('descp') %>" maxlength="50" size="50" /></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><input type="submit" name="btnEdit" value="Edit" /> &nbsp;&nbsp;&nbsp;&nbsp;
											   <!-- <input type="reset" value="Reset" onclick="setFocus(this.form);" /> -->
				</td>
			</tr>
		</table>
	</form>
	<% 
		rs.Close;
		conn.Close;
	}else
		Response.Write("Error");
	%>
</body>
</html>