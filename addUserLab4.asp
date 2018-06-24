<%@ language="javascript" %>
<!DOCTYPE html>
<html>
<head>
	<title>Lab 4</title>
	
	<script>
		function setFocus(f)
		{
			f.elements["name"].focus();
		}
	</script>
	
	<%
		if(Request.Form("btnAdd")=="Add")
		{
			var conn, rs, sql, n, desc;
			
			n = Request.Form("name");
			desc = Request.Form("desc");
			
			conn = Server.CreateObject("ADODB.Connection");
			conn.Open("Provider=Microsoft.Jet.OLEDB.4.0; Data Source="+Server.MapPath("testdb.mdb"));
			rs = Server.CreateObject("ADODB.Recordset");
			
			sql = "INSERT INTO PTj (name, descp) VALUES ('"+n+"','"+desc+"')";
			conn.Execute(sql);
			
			Response.Write("<script> alert('A new record has been inserted into the database'); window.location='view.asp'; </script>");			
		
			rs.Close;
			conn.Close;
		}
	%>
	
</head>
<body>
	
	<center><h1>::ADD::</h1></center>
	<form name="frm" action="" method="post">
		<table width="40%" border="1" align="center" cellpadding="10">
			<tr>
				<td>Name</td>
				<td><input type="text" name="name" id="name" maxlength="30" size="30" autofocus /></td>
			</tr>
			<tr>
				<td>Description</td>
				<td><input type="text" name="desc" id="desc" maxlength="50" size="50" /></td>
			</tr>
			<tr>
				<td colspan="2" align="center"><input type="submit" name="btnAdd" value="Add" /> &nbsp;&nbsp;&nbsp;&nbsp;
											   <input type="reset" value="Reset" onclick="setFocus(this.form);" />
				</td>
			</tr>
		</table>
	</form>
</body>
</html>