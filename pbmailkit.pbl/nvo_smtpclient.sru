forward
global type nvo_smtpclient from smtpclient
end type
end forward

global type nvo_smtpclient from smtpclient
end type
global nvo_smtpclient nvo_smtpclient

type variables
Public:
// Secure Protocol
Constant Integer ii_ALL = 0
Constant Integer ii_TLS10 = 1
Constant Integer ii_TLS11 = 2
Constant Integer ii_TLS12 = 3
Constant Integer ii_TLS13 = 4

// Priority
Constant Integer ii_NormalPriority = 0
Constant Integer ii_LowPriority = 1
Constant Integer ii_HighPriority = 2

//Reset 
Constant Integer ii_ResetALL=0
Constant Integer ii_ResetRecipients=1
Constant Integer ii_ResetCcs=2
Constant Integer ii_ResetBccs=3
Constant Integer ii_ResetAttachments=4
Constant Integer ii_ResetLinkedResources=5
Constant Integer ii_ResetSender=6



// Character Set Values
// Windows-1256	Arabic (Windows)
// Windows-1257	Baltic (Windows)
// iso-8859-2		Central European (ISO)
// Windows-1250	Central European (Windows)
// gb2312			Chinese Simplified (GB2312)
// hz-gb-2312		Chinese Simplified (HZ)
// big5				Chinese Traditional (Big5)
// koi8-r			Cyrilic (KOI8-R)
// Windows-1251	Cyrillic (Windows)
// Windows-1253	Greek (Windows)
// Windows-1255	Hebrew (Windows)
// iso-2022-jp		Japanese (JIS)
// ks_c_5601		Korean
// euc-kr			Korean (EUC)
// iso-8859-15		Latin 9 (ISO)
// Windows-874		Thai (Windows)
// Windows-1254	Turkish (Windows)
// UTF-7				Unicode (UTF-7)
// UTF-8				Unicode (UTF-8)
// Windows-1258	Vietnamese (Windows)
// iso-8859-1		Western European (ISO)
// Windows-1252	Western European (Windows)
end variables

forward prototypes
public function integer of_send () throws n_ex
private function string of_valid_email (string as_email) throws n_ex
public subroutine of_add_bcc (string as_bcc_name, string as_bcc_email) throws n_ex
public subroutine of_add_cc (string as_cc_name, string as_cc_email) throws n_ex
public subroutine of_add_recipient (string as_recipient_name, string as_recipient_email) throws n_ex
public subroutine of_set_bcc (string as_bcc_name[], string as_bcc_email[]) throws n_ex
public subroutine of_set_cc (string as_cc_name[], string as_cc_email[]) throws n_ex
public subroutine of_set_recipient (string as_recipient_name[], string as_recipient_email[]) throws n_ex
public subroutine of_add_attachment (string as_attachment) throws n_ex
public subroutine of_enable_tls (boolean ab_enable_tls) throws n_ex
public subroutine of_set_attachment (string as_attachment[]) throws n_ex
public subroutine of_set_encoding (string as_encoding) throws n_ex
public subroutine of_set_login (string as_user_name, string as_user_pass) throws n_ex
public subroutine of_set_message (string as_message) throws n_ex
public subroutine of_set_message (string as_message, boolean ab_html) throws n_ex
public subroutine of_set_port (integer ai_port) throws n_ex
public subroutine of_set_priority (integer ai_priority) throws n_ex
public subroutine of_set_secure_protocol (integer ai_secure_protocol) throws n_ex
public subroutine of_set_subject (string as_subject) throws n_ex
public subroutine of_set_host (string as_host) throws n_ex
public subroutine of_set_sender (string as_sender_name, string as_sender_email) throws n_ex
private function long of_parse (readonly string as_string, readonly string as_separator, ref string as_outarray[]) throws n_ex
end prototypes

public function integer of_send () throws n_ex;Integer li_rc
String ls_errorText

//send the email message
li_rc = This.Send()

//Reseteamos Variables Por si no destruimos el Objeto
This.Message.Reset(ii_ResetALL)  

CHOOSE CASE li_rc
	CASE -1
	 	 ls_errorText = "Ha ocurrido un error general."
	CASE -2
	  	ls_errorText = "No se pudo conectar al servicio a través del proxy."
	CASE -3
	 	 ls_errorText = "No se pudo resolver el host del proxy proporcionado."
	CASE -4
		  ls_errorText = "No se pudo resolver el host remoto proporcionado."
	CASE -5
		  ls_errorText = "Fallo al conectar al host."
	CASE -6
	 	 ls_errorText = "El formato del host es incorrecto o está ausente."
	CASE -7
	 	 ls_errorText = "El protocolo no está soportado."
	CASE -8
	 	 ls_errorText = "Error en la conexión SSL."
	CASE -9
	  	ls_errorText = "El certificado del servidor está revocado."
	CASE -10
	 	 ls_errorText = "La autenticación del certificado del servicio ha fallado."
	CASE -11
	  ls_errorText = "Tiempo de espera de operación agotado."
	CASE -12
	 	 ls_errorText = "El servidor remoto denegó el inicio de sesión de curl."
	CASE -13
	 	 ls_errorText = "Error al enviar datos de red."
	CASE -14
	 	 ls_errorText = "Error al recibir datos de red."
	CASE -15
	 	 ls_errorText = "Nombre de usuario o contraseña incorrectos."
	CASE -16
	 	 ls_errorText = "Error al leer el archivo local."
	CASE -17
	  	ls_errorText = "No se ha especificado ningún remitente."
	CASE -18
	 	 ls_errorText = "No se han especificado destinatarios."
	CASE ELSE
	 	 ls_errorText = "Error desconocido."
END CHOOSE

IF li_rc <> 1 THEN
	gf_throw(PopulateError(li_rc * -1,  ls_errorText))
END IF	

RETURN li_rc

end function

private function string of_valid_email (string as_email) throws n_ex;String ls_localpart, ls_domain, ls_domainpart[]
Integer li_pos, li_idx, li_max, li_asc

If trim(as_email) = "" Then
	gf_throw(PopulateError(0, "¡Hay que indciar una dirección de correo válida!"))
End If
 

If Len(as_email) > 254 Then
	gf_throw(PopulateError(1, "¡La dirección de correo electrónico no puede tener más de 254 caracteres!"))
End If

li_pos = Pos(as_email, "@")
If li_pos < 2 Then
	gf_throw(PopulateError(2, "¡La dirección de correo electrónico debe tener el carácter @!"))
End If

If LastPos(as_email, "@") > Pos(as_email, "@") Then
	gf_throw(PopulateError(3, "¡La dirección de correo electrónico no puede tener más de un carácter @!"))
End If

li_pos = Pos(as_email, " ")
If li_pos > 0 Then
	gf_throw(PopulateError(4, "¡La dirección de correo electrónico no puede contener espacios!"))
End If

// separar local y dominio
li_pos = Pos(as_email, "@")
ls_localpart = Left(as_email, li_pos - 1)
ls_domain    = Mid(as_email, li_pos + 1)

If Len(ls_localpart) > 64 Then
	gf_throw(PopulateError(5, "¡La parte del buzón de la dirección de correo electrónico no puede tener más de 64 caracteres!"))
End If

If Len(ls_domain) > 253 Then
	gf_throw(PopulateError(6, "¡La parte del dominio de la dirección de correo electrónico no puede tener más de 253 caracteres!"))
End If

If Left(ls_localpart, 1) = "." Then
	gf_throw(PopulateError(7, "¡La parte del buzón de la dirección de correo electrónico no puede comenzar con un punto!"))
	
End If

If Right(ls_localpart, 1) = "." Then
	gf_throw(PopulateError(8, "¡La parte del buzón de la dirección de correo electrónico no puede terminar con un punto!"))	
End If

If Pos(ls_localpart, "..") > 0 Then
	gf_throw(PopulateError(9, "¡La parte del buzón de la dirección de correo electrónico no puede tener más de un punto consecutivo!"))
End If

// verificar caracteres permitidos en local
li_max = Len(ls_localpart)
For li_idx = 1 To li_max
	li_asc = Asc(Mid(ls_localpart, li_idx, 1))
	choose case li_asc
		case 65 to 90
			// Letras minúsculas a-z
		case 97 to 122
			// Letras mayúsculas A-Z
		case 48 to 57
			// Dígitos 0-9
		case 33, 35 to 39, 42, 43, 45, 47, 61, 63, 94 to 96, 123 to 126
			// Caracteres !#$%&'*+-/=?^_`{|}~
		case 46
			// Punto
		case else
			gf_throw(PopulateError(10, "¡La parte del buzón de la dirección de correo electrónico contiene caracteres no válidos!"))		
	end choose
Next

If Left(ls_domain, 1) = "." Then
	gf_throw(PopulateError(11, "¡La parte del dominio de la dirección de correo electrónico no puede comenzar con un punto!"))	
End If

If Right(ls_domain, 1) = "." Then
	gf_throw(PopulateError(12, "¡La parte del dominio de la dirección de correo electrónico no puede terminar con un punto!"))
End If

li_pos = Pos(ls_domain, ".")
If li_pos = 0 Then
	gf_throw(PopulateError(13, "¡La parte del dominio de la dirección de correo electrónico debe tener al menos un punto!"))
End If

// verificar caracteres permitidos en el dominio
li_max = Len(ls_domain)
For li_idx = 1 To li_max
	li_asc = Asc(Mid(ls_domain, li_idx, 1))
	choose case li_asc
		case 65 to 90
			// Letras minúsculas a-z
		case 97 to 122
			// Letras mayúsculas A-Z
		case 48 to 57
			// Dígitos 0-9
		case 45
			// Guión
		case 46
			// Punto
		case else
			gf_throw(PopulateError(14, "¡La parte del dominio de la dirección de correo electrónico contiene caracteres no válidos!"))	
	end choose
Next

// dividir el dominio en partes
li_max =of_parse(ls_domain, ".", REF ls_domainpart[])

If li_max > 127 Then
	gf_throw(PopulateError(15, "¡La parte del dominio de la dirección de correo electrónico contiene demasiados puntos!"))
End If

For li_idx = 1 To li_max
	If Left(ls_domainpart[li_idx], 1) = "-" Then
		gf_throw(PopulateError(16, "¡La parte del dominio de la dirección de correo electrónico no puede tener un guión y un punto uno al lado del otro!"))
	End If
	If Right(ls_domainpart[li_idx], 1) = "-" Then
		gf_throw(PopulateError(17, "¡La parte del dominio de la dirección de correo electrónico no puede tener un guión y un punto uno al lado del otro!"))
	End If
Next

RETURN as_email
end function

public subroutine of_add_bcc (string as_bcc_name, string as_bcc_email) throws n_ex;as_bcc_email  = of_valid_email(as_bcc_email)
as_bcc_name  = trim(as_bcc_name)

IF as_Bcc_name = "" THEN 
		This.Message.AddBcc(as_bcc_email)
ELSE	
		This.Message.AddBcc(as_bcc_email, as_bcc_name)
END IF



end subroutine

public subroutine of_add_cc (string as_cc_name, string as_cc_email) throws n_ex;as_cc_email  = of_valid_email(as_cc_email)
as_cc_name  = trim(as_cc_name)

IF as_cc_name = "" THEN  
	This.Message.AddCc(as_cc_email)
ELSE
	This.Message.AddCc(as_cc_email, as_cc_name)
END IF
	


end subroutine

public subroutine of_add_recipient (string as_recipient_name, string as_recipient_email) throws n_ex;as_recipient_email = of_valid_email(as_recipient_email)
as_recipient_name= trim(as_recipient_name)

IF as_recipient_name = "" THEN  
	This.Message.AddRecipient(as_recipient_email)
ELSE	
	This.Message.AddRecipient(as_recipient_email, as_recipient_name)
END IF
end subroutine

public subroutine of_set_bcc (string as_bcc_name[], string as_bcc_email[]) throws n_ex;integer li_bcc, li_TotalBcc

This.Message.Reset(ii_ResetBccs)  

li_TotalBcc = UpperBound(as_bcc_email[])

FOR li_bcc = 1 TO li_TotalBcc
	of_add_bcc(as_bcc_name[li_bcc], as_bcc_email[li_bcc])
NEXT
end subroutine

public subroutine of_set_cc (string as_cc_name[], string as_cc_email[]) throws n_ex;integer li_Cc, li_TotalCc

This.Message.Reset(ii_ResetCcs)  

li_TotalCc = UpperBound(as_Cc_email[])

FOR li_Cc = 1 TO li_TotalCc
	of_add_Cc(as_Cc_name[li_Cc], as_Cc_email[li_Cc])
NEXT
end subroutine

public subroutine of_set_recipient (string as_recipient_name[], string as_recipient_email[]) throws n_ex;integer li_Recipient, li_TotalRecipient

This.Message.Reset(ii_ResetRecipients)  

li_TotalRecipient = UpperBound(as_Recipient_email[])

FOR li_Recipient = 1 TO li_TotalRecipient
	of_add_Recipient(as_Recipient_name[li_Recipient], as_Recipient_email[li_Recipient])
NEXT
end subroutine

public subroutine of_add_attachment (string as_attachment) throws n_ex;as_attachment = trim(as_attachment)

IF NOT FileExists(as_attachment) THEN
	gf_throw(PopulateError(1, "¡ El archivo adjunto "+as_attachment+" no existe !"))
END IF	

This.Message.AddAttachment(as_attachment)


end subroutine

public subroutine of_enable_tls (boolean ab_enable_tls) throws n_ex;If isnull(ab_enable_tls) Then
	gf_throw(PopulateError(1, "¡Hay que indicar si queremos usar TLS!"))
End If

This.EnableTLS = ab_enable_tls
end subroutine

public subroutine of_set_attachment (string as_attachment[]) throws n_ex;integer li_atach, li_TotalAttachment

This.Message.Reset(ii_ResetAttachments)  

li_TotalAttachment = UpperBound(as_attachment[])

FOR li_atach = 1 TO li_TotalAttachment
	of_add_attachment(as_attachment[li_atach])
NEXT
end subroutine

public subroutine of_set_encoding (string as_encoding) throws n_ex;This.message.Encoding = as_encoding	

end subroutine

public subroutine of_set_login (string as_user_name, string as_user_pass) throws n_ex;If trim(as_user_name) = "" Then
	gf_throw(PopulateError(1, "¡Hay que indicar el Nombre de Usuario!"))
End If
If trim(as_user_pass) = "" Then
	gf_throw(PopulateError(2, "¡Hay que indicar el Password de Usuario!"))
End If

This.UserName =  as_user_name
This.Password = as_user_pass
end subroutine

public subroutine of_set_message (string as_message) throws n_ex;of_set_message(as_message, FALSE)
end subroutine

public subroutine of_set_message (string as_message, boolean ab_html) throws n_ex;IF ab_html = TRUE THEN
	This.Message.HTMLBody = as_message
ELSE
	This.Message.TextBody= as_message
END IF	

end subroutine

public subroutine of_set_port (integer ai_port) throws n_ex;IF ai_port <> 25 AND ai_port <> 587 AND ai_port <> 465 THEN
	gf_throw(PopulateError(1, "¡ El Puerto "+string(ai_port)+" no es válido !"))
END IF	

This.Port =ai_port


end subroutine

public subroutine of_set_priority (integer ai_priority) throws n_ex;IF ai_priority < 0 OR ai_priority > 2 THEN
	gf_throw(PopulateError(1, "¡ La Prioridad debe estar entre 0 y 2 !"))
END IF	

This.message.Priority = ai_priority

end subroutine

public subroutine of_set_secure_protocol (integer ai_secure_protocol) throws n_ex;IF ai_secure_protocol < 0 or ai_secure_protocol > 4  THEN
	gf_throw(PopulateError(1, "¡ El Protocolo de Seguridad debe estar entre 0 y 4 !"))
END IF	

This.SecureProtocol = ai_secure_protocol

end subroutine

public subroutine of_set_subject (string as_subject) throws n_ex;This.message.Subject = trim(as_subject)

end subroutine

public subroutine of_set_host (string as_host) throws n_ex;If trim(as_host) = "" Then
	gf_throw(PopulateError(1, "¡Hay que indicar un Servidor SMTP!"))
End If

This.Host = as_host

end subroutine

public subroutine of_set_sender (string as_sender_name, string as_sender_email) throws n_ex;as_sender_email = of_valid_email(as_sender_email)
as_sender_name = trim(as_sender_name)

This.Message.SetSender(as_sender_email, as_sender_name)
end subroutine

private function long of_parse (readonly string as_string, readonly string as_separator, ref string as_outarray[]) throws n_ex;// -----------------------------------------------------------------------------
// SCRIPT:     of_Parse
//
// PURPOSE:    This function parses a string into an array.
//
// ARGUMENTS:  as_string		- The string to be separated
//					as_separator	- The separator characters
//					as_outarray		- By ref output array
//
//	RETURN:		The number of items in the array
//
// DATE        CHANGED BY	DESCRIPTION OF CHANGE / REASON
// ----------  ----------  -----------------------------------------------------
// 08/12/2017  RolandS		Initial creation
// -----------------------------------------------------------------------------

Long ll_PosEnd, ll_PosStart = 1, ll_SeparatorLen, ll_Counter = 1

If UpperBound(as_OutArray) > 0 Then as_OutArray = {""}

ll_SeparatorLen = Len(as_Separator)

ll_PosEnd = Pos(as_String, as_Separator, 1)

Do While ll_PosEnd > 0
	as_OutArray[ll_Counter] = Mid(as_String, ll_PosStart, ll_PosEnd - ll_PosStart)
	ll_PosStart = ll_PosEnd + ll_SeparatorLen
	ll_PosEnd = Pos(as_String, as_Separator, ll_PosStart)
	ll_Counter++
Loop

as_OutArray[ll_Counter] = Right(as_String, Len(as_String) - ll_PosStart + 1)

Return ll_Counter

end function

on nvo_smtpclient.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nvo_smtpclient.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

