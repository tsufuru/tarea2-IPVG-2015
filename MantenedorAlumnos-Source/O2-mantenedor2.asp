<%Set Conn= Server.CreateObject("ADODB.connection")
Conn.open "DSN=dsnalumnos;UID=invitado;PWD=2015;DATABASE=mantencion"
	RUT = REQUEST.QUERYSTRING("RUT") 		
	NOMBRES = REQUEST.QUERYSTRING("NOMBRES")
	MAIL = REQUEST.QUERYSTRING("MAIL")
	
	if (RUT<>"") then
	
	
	SQL = "INSERT INTO mantencion.dbo.alumnos " & _
				"(RUT, NOMBRES, MAIL) " & _
				"VALUES " & _
				"('" & RUT & "', '" & NOMBRES & "', '" & CORREO & "')" 
			Conn.execute(SQL)
			RUT = ""
			NOMBRES = ""
			MAIL = ""
	end if
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<script language="javascript">
function volver()
{
	document.form1.action="mantenedor1.asp"
	document.form1.submit();
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Documento sin t&iacute;tulo</title>
</head>

<body>

<form method="post" action="O2-mantenedor2.asp" name="form1">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#c0c0c0">
             
    <tr align="center" valign="middle"> 
        <td height="25" colspan="6" nowrap bordercolor="#FFFFFF" bgcolor="#FFFFFF"><b><font face="Verdana,Arial, Helvetica, sans-serif" size="1">RUT 
                <input type="text" name="RUT" maxlength="5" size="15" class="texto">
                <font face="Verdana, Arial, Helvetica, sans-serif">NOMBRES</font> 
                <input type="text" name="NOMBRES" size="60" maxlength="50" class="texto">
                </font></b></td>
            </tr>
            <tr align="center" valign="middle">
              <td height="25" colspan="6" nowrap bordercolor="#FFFFFF" bgcolor="#FFFFFF"><b><font face="Arial, Helvetica, sans-serif" size="1"><font face="Verdana, Arial, Helvetica, sans-serif">MAIL</font> 
                <input type="text" name="MAIL" size="60" maxlength="50" class="texto">              
                
                </font></b></td>
            </tr>
          </table>
		  <input type="submit" value="Insertar">
		  <input type="button" value="Volver" onclick="volver()">
		</form>
</body>
</html>
