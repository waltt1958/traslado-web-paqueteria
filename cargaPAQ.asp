<HTML>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title> PAQUETES RECIBIDOS EN PAQUETERIA</title>
</HEAD>
<body bgcolor="#FF00FF" onload="maximizar()">
<H3 align= "right">Hoy es: <%=date%></H3>
<center><h1><p align="center"><u><b><font size="12"> INGRESOS DE PAQUETES EN SECTOR PAQUETERIA (exclusivamente)</font size></b></u></p> </h1></center>

<center><h1><p align="center"><u><b><font size="9"> LEA EL NRO DE OEP (sin parentesis ni espacios)</font size></b></u></p> </h1></center>
<br>
<hr size= 6 color="black"></hr>

<br>
<FORM name="inicia" action="procesaPAQ.asp" method="post">

<TABLE WIDTH= "100%" border="3" bordercolor="black">

<TR>
<TD align="center" fontsize= "8"><H2><B>PAQUETE</B></H2></TD>
</TR>

<TR>
<TD align="center"><INPUT TYPE="TEXT" NAME="OEP" onblur="validar()" OnKeyPress="if (event.keyCode==32) event.returnValue= false;" maxlength="19"></TD>
</TR>
</TABLE>
</FORM>


<a href="CARGADOenPAQ.asp" TARGET="_blank"><font size="6"><FONT COLOR="BLACK"><b>Ver e imprimir todo lo cargado</b></a>
<br>
<BR>
<a href="intermedia.asp" TARGET="_self"><font size="6"><FONT COLOR="BLACK"><b>Ver diferencias de carga</b></a>
<BR>
<BR>
<a href="intermedia.asp" TARGET="_blank"><font size="6"><FONT COLOR="BLACK"><b>Cerrar diferencias de carga hasta ayer</b></a>
<br>
<br>
<a href="cierreFORZADO.asp" TARGET="_blank"><font size="6"><FONT COLOR="BLACK"><b>Ver todos los paquetes forzados el cierre hasta ayer</b></a>
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