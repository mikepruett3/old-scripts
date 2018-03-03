<!--
Remove this commented block if code is deployed

<SCRIPT Language="VBScript">
-->
Function frmExample_onsubmit
   If instr(frmExample.txtemail.Value, "@") = 0 OR _
	instr(frmExample.txtemail.Value, ".") = 0 OR _
	Len(frmExample.txtemail.Value) < 7 Then
	window.alert "Please specify a valid e-mail address!"
	frmExample_onsubmit = False
   End IF
	
End Function
<!--
Remove this commented block if code is deployed
</SCRIPT>
-->