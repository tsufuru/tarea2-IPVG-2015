<%
if REQUEST.FORM <> "" then

	Set Conn= Server.CreateObject("ADODB.connection")
	Conn.open "DSN=dsnalumnos;UID=invitado;PWD=2015;DATABASE=mantencion"
	
	RUT = REQUEST.FORM("RUT") 		
	NOMBRES = REQUEST.FORM("NOMBRES")
	MAIL = REQUEST.FORM("MAIL")
	DIRECCION = REQUEST.FORM("DIRECCION")
	
	if (RUT<>"" and NOMBRES<>"" and MAIL<>"" and DIRECCION<>"") then
		
		SQL = "INSERT INTO dbo.alumnos " & _
			"(RUT, NOMBRE, MAIL, DIRECCION) " & _
			"VALUES " & _
	  		"('" & RUT & "', '" & NOMBRES & "', '" & MAIL & "', '" & DIRECCION & "')" 
			
		Conn.execute(SQL)
		
		RUT = ""
		NOMBRES = ""
		MAIL = ""
		DIRECCION = ""
		
	end if
	
end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script language="javascript">
function volver()
{
	document.location.href="O2-mantenedor1.asp"
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Documento sin t&iacute;tulo</title>
</head>

<body>

<form method="post" action="O2-mantenedor2.asp" name="form1">
	<table border="0" cellpadding="0" cellspacing="0" bordercolor="#c0c0c0">
		<tr> 
			<td height="25" nowrap bordercolor="#FFFFFF" bgcolor="#FFFFFF">
				<font face="Verdana, Arial, Helvetica, sans-serif">RUT</font>
			</td>
			<td>
				<input type="text" name="RUT" maxlength="20" size="20" class="texto"> 
			</td>
		</tr>
		<tr> 
			<td height="25" nowrap bordercolor="#FFFFFF" bgcolor="#FFFFFF">
				<font face="Verdana, Arial, Helvetica, sans-serif">NOMBRES</font> 
			</td>
			<td>
				<input type="text" name="NOMBRES" size="60" maxlength="50" class="texto">
			</td>
		</tr>
		<tr> 
			<td height="25" nowrap bordercolor="#FFFFFF" bgcolor="#FFFFFF">
				<font face="Verdana, Arial, Helvetica, sans-serif">MAIL</font> 
			</td>
			<td>
				<input type="text" name="MAIL" size="60" maxlength="50" class="texto">         
			</td>
		</tr>
	</table>
    <input type="submit" value="Insertar">
	<input type="button" value="Volver" onclick="volver()">
</form>
</body>
</html>
