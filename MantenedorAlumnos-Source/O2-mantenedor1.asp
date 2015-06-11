<%
Set Conn= Server.CreateObject("ADODB.connection")
Conn.open "DSN=dsnalumnos;UID=invitado;PWD=2015;DATABASE=mantencion"

SQL = "SELECT * FROM dbo.alumnos ORDER BY NOMBRE"
Set REG1 = Conn.execute(SQL)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Documento sin t&iacute;tulo</title>
</head>

<body>

 <table width="100%" border="0" cellpadding="0" cellspacing="0">
            <tr align="center" valign="middle"> 
              <td width="10%" height="10" bgcolor="#ffffff"></td>
            </tr>
          </table>
		</td>
      </tr>
      <tr> 
        <td bgcolor="#ffffff"> 
          <table width="100%" border=1 align="center" cellpadding=1  cellspacing=0 bordercolor="#00009C" bgcolor=#00009C >
            <tr bgcolor="#00009C"> 
              <td width="10%" valign=top nowrap> 
                <div align="center"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">RUT</font></b></div>
              </td>
              <td width="42%" valign=top> 
                <div align="center"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">NOMBRES</font></b></div>
              </td>
              <td width="34%" valign=top> 
                <div align="center"><b><font size="1" color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif">MAIL</font></b></div>
              </td>
             
            </tr>
            <% 
				do While not REG1.eof
			%>
            <tr bgcolor="#FFFFFF"> 
              <td width="10%" valign=top nowrap> 
                <div align="center"><font  face="Arial, Helvetica, sans-serif" size=1>
					<%=REG1("RUT")%>
				</font></div>
              </td>
              <td width="42%" valign=top> 
                <div align="left"><font face="Arial, Helvetica, sans-serif" size="1">
					<%=REG1("NOMBRE")%>
				</font></div>
              </td>
              <td width="34%" valign=top> 
                <div align="left"><font face="Arial, Helvetica, sans-serif" size="1">
					<%=REG1("MAIL")%>
				</font></div>
              </td>
            </tr>
            <% 
			   	REG1.movenext
			  	loop
				REG1.close
				Conn.close
			%>
          </table>
		  <a href="O2-mantenedor2.asp">
			<input type="button" value="Insertar">
		  </a>

</body>
</html>
