<HTML>
<HEAD>
<link rel="stylesheet" title="estilos.css" type="text/css" href="estilos.css">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title> PAQUETES A PASAR A PAQUETERIA</title>
</HEAD>
<body onload="maximizar()">
<H3 align= "right">Hoy es: <%=date%></H3>
<br>
<center><h1><p align="center"><u><b><font size="12"> INGRESOS DE PAQUETES EN EL CI (exclusivamente)</font size></b></u></p> </h1></center>

<center><h1><p align="center"><u><b><font size="9"> LEA EL NRO DE OEP (sin parentesis ni espacios)</font size></b></u></p> </h1></center>
<br>
<hr size= 6 color="black"></hr>
<br>
<br>
<FORM name="inicia" action="procesa.asp" method="post">

<TABLE WIDTH= "100%" border="3" bordercolor="black">

<TR>
<TD align="center" fontsize= "8"><H2><B>PAQUETE</B></H2></TD>
</TR>

<TR>
<TD align="center"><INPUT TYPE="TEXT" NAME="OEP" onblur="validar()" OnKeyPress="if (event.keyCode==32) event.returnValue= false;" maxlength="19"></TD>
</TR>
</TABLE>
</FORM>

<BR>
<BR>
<BR>
<BR>
<BR>
<table width="100%">
<tr align="center">
<td align="center">
<a href="CARGADOenCI.asp" target="_self"><input type="button" name="cargadoENCI" value="Ver e imprimir lo cargado" style="FONT-SIZE: 20pt; width:600px;border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a></td>
<td align="center"><a href="cierreFORZADO.asp" target="_self"><input type="button" name="cierreFORZADO" value="Paquetes forzado el cierre hasta ayer" style="FONT-SIZE: 20pt;width:600px; border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a></td>
</tr>
<tr align="center">
<td colspan=2>
<br>
<a href="index.asp" target="_self"><input type="button" name="cargaCI" value="INICIO" style="FONT-SIZE: 20pt;width:150px;border: 5px solid; [b]FONT-FAMILY: Verdana, boldt[/b];
BACKGROUND-COLOR: #C0C0C0"></a>
</td>
</tr>
</table>
</body>

<SCRIPT Language="javascript" type="text/javascript">
function validar() {
if (window.document.inicia.OEP.value=="") {

window.document.inicia.OEP.focus();
return;
}
if (isNaN(window.document.inicia.OEP.value)) {
alert("Tiene que colocar solo numeros");
window.document.inicia.OEP.focus();
return;
}
window.document.inicia.submit();
}
</script>

<SCRIPT Language="javascript" type="text/javascript">
function maximizar() {
window.document.inicia.OEP.focus();
window.moveTo(0,0);
window.resizeTo(screen.width,screen.height);
document.inicia.OEP.focus();
}
</SCRIPT>
</HTML>