$PBExportHeader$n_cst_smtpclient.sru
forward
global type n_cst_smtpclient from nonvisualobject
end type
end forward

global type n_cst_smtpclient from nonvisualobject
end type
global n_cst_smtpclient n_cst_smtpclient

type variables
Private:
SmtpClient in_Smtp
boolean ib_enable_tls
integer ii_secure_protocol
string is_encoding
boolean ib_html
integer ii_priority
string is_host
integer ii_port
string is_user_name, is_user_pass
string is_sender_name, is_sender_email
string is_attachment[]
string is_message
string is_subject
string is_recipient_name[], is_recipient_email[]
string is_cc_name[], is_cc_email[]
string is_bcc_name[], is_bcc_email[]
String Is_ResetArray[] 

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
private subroutine of_reset () throws n_ex
private function long of_parse (readonly string as_string, readonly string as_separator, ref string as_outarray[]) throws n_ex
public subroutine of_enable_tls (boolean ab_enable_tls) throws n_ex
end prototypes

public function integer of_send () throws n_ex;Integer li_rc, li_attachments, li_idx, li_Recipient, li_CC, li_BCC
String ls_errorText 

in_Smtp =  CREATE SMTPClient

//set the email account information     
in_Smtp.Host = is_host
in_Smtp.Port = ii_port
in_Smtp.UserName = is_user_name
in_Smtp.Password = is_user_pass
in_Smtp.EnableTLS = ib_enable_tls
in_Smtp.SecureProtocol = ii_secure_protocol

//set the email message
in_Smtp.message.Encoding = is_encoding	
in_Smtp.message.Priority = ii_priority
in_Smtp.Message.SetSender(is_sender_email, is_sender_name)

li_Recipient = UpperBound(is_recipient_email[])

FOR li_idx = 1 TO li_Recipient
	IF trim(is_recipient_name[li_idx]) = "" THEN  
		in_Smtp.Message.AddRecipient(is_recipient_email[li_idx])
	ELSE	
		in_Smtp.Message.AddRecipient(is_recipient_email[li_idx], is_recipient_name[li_idx] )
	END IF
NEXT	

li_CC = UpperBound(is_cc_email[])

FOR li_idx = 1 TO li_CC
	IF trim(is_cc_name[li_idx]) = "" THEN  
		in_Smtp.Message.AddCc(is_cc_email[li_idx])
	ELSE
		in_Smtp.Message.AddCc(is_cc_email[li_idx], is_cc_name[li_idx])
	END IF
NEXT	


li_BCC = UpperBound(is_bcc_email[])

FOR li_idx = 1 TO li_BCC
	IF trim(is_Bcc_name[li_idx]) = "" THEN 
		in_Smtp.Message.AddBcc(is_bcc_email[li_idx])
	ELSE	
		in_Smtp.Message.AddBcc(is_bcc_email[li_idx], is_bcc_name[li_idx])
	END IF
NEXT	

in_Smtp.message.Subject = is_subject

li_attachments = UpperBound(is_attachment[]) 

FOR li_idx = 1 TO li_attachments
	in_Smtp.Message.AddAttachment( is_attachment[li_idx])
NEXT

IF ib_html = TRUE THEN
	in_Smtp.Message.HTMLBody = is_message
ELSE
	in_Smtp.Message.TextBody= is_message
END IF	


//send the email message
li_rc = in_Smtp.Send()

//Reseteamos Variables Por si no destruimos el Objeto
of_reset()
Destroy in_Smtp

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
li_max = of_parse(ls_domain, ".", REF ls_domainpart[])

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

public subroutine of_add_bcc (string as_bcc_name, string as_bcc_email) throws n_ex;integer li_idx, li_TotalArray

li_TotalArray = UpperBound(is_bcc_email[])

li_idx = li_TotalArray + 1

is_bcc_email[li_idx] = of_valid_email(as_bcc_email)
is_bcc_name[li_idx] = trim(as_bcc_name)



end subroutine

public subroutine of_add_cc (string as_cc_name, string as_cc_email) throws n_ex;Integer li_idx, li_TotalArray

li_TotalArray = UpperBound(is_cc_email[])

li_idx = li_TotalArray + 1

is_cc_email[li_idx]  = of_valid_email(as_cc_email)
is_cc_name[li_idx]  = trim(as_cc_name)

end subroutine

public subroutine of_add_recipient (string as_recipient_name, string as_recipient_email) throws n_ex;Integer li_idx, li_TotalArray

li_TotalArray = UpperBound(is_recipient_email[])

li_idx = li_TotalArray + 1

is_recipient_email[li_idx] = of_valid_email(as_recipient_email)
is_recipient_name[li_idx] = trim(as_recipient_name)

end subroutine

public subroutine of_set_bcc (string as_bcc_name[], string as_bcc_email[]) throws n_ex;integer li_idx, li_TotalArray

//Reseteamos Las Arrays
is_bcc_name[] = is_ResetArray[]
is_bcc_email = is_ResetArray[]

li_TotalArray = UpperBound(as_bcc_email[])

FOR li_idx = 1 TO li_TotalArray
	of_add_bcc(as_bcc_name[li_idx], as_bcc_email[li_idx])
NEXT

end subroutine

public subroutine of_set_cc (string as_cc_name[], string as_cc_email[]) throws n_ex;integer li_idx, li_TotalArray

//Reseteamos Las Arrays
is_cc_name[] = is_ResetArray[]
is_cc_email = is_ResetArray[]

li_TotalArray = UpperBound(as_cc_email[])

FOR li_idx = 1 TO li_TotalArray
	of_add_cc(as_cc_name[li_idx], as_cc_email[li_idx])
NEXT

end subroutine

public subroutine of_set_recipient (string as_recipient_name[], string as_recipient_email[]) throws n_ex;integer li_idx, li_TotalArray

//Reseteamos las Arrays
is_recipient_name[] = is_ResetArray[]
is_recipient_email[] = is_ResetArray[]

li_TotalArray = UpperBound(as_recipient_email[])

FOR li_idx = 1 TO li_TotalArray
	 of_add_recipient(as_recipient_name[li_idx], as_recipient_email[li_idx])
NEXT

end subroutine

public subroutine of_add_attachment (string as_attachment) throws n_ex;integer li_idx, li_TotalArray

IF NOT FileExists(as_attachment) THEN
	gf_throw(PopulateError(1, "¡ El archivo adjunto "+as_attachment+" no existe !"))
END IF	

li_TotalArray = UpperBound(is_attachment[])

li_idx = li_TotalArray + 1

is_attachment[li_idx] = trim(as_attachment)

end subroutine

public subroutine of_set_attachment (string as_attachment[]) throws n_ex;integer li_idx, li_TotalArray

//Reseteamos La Array
is_attachment[] = is_ResetArray[]

IF as_attachment[] =  is_ResetArray[] THEN RETURN

li_TotalArray = UpperBound(as_attachment[])

FOR li_idx = 1 TO li_TotalArray
	of_add_attachment(as_attachment[li_idx])
NEXT
end subroutine

public subroutine of_set_encoding (string as_encoding) throws n_ex;is_encoding = as_encoding	

end subroutine

public subroutine of_set_login (string as_user_name, string as_user_pass) throws n_ex;If trim(as_user_name) = "" Then
	gf_throw(PopulateError(1, "¡Hay que indicar el Nombre de Usuario!"))
End If
If trim(as_user_pass) = "" Then
	gf_throw(PopulateError(2, "¡Hay que indicar el Password de Usuario!"))
End If

is_user_name = as_user_name
is_user_pass = as_user_pass

end subroutine

public subroutine of_set_message (string as_message) throws n_ex;of_set_message(as_message, FALSE)
end subroutine

public subroutine of_set_message (string as_message, boolean ab_html) throws n_ex;is_message = as_message
ib_html = ab_html

end subroutine

public subroutine of_set_port (integer ai_port) throws n_ex;IF ai_port <> 25 AND ai_port <> 587 AND ai_port <> 465 THEN
	gf_throw(PopulateError(1, "¡ El Puerto "+string(ai_port)+" no es válido !"))
END IF	

ii_port = ai_port
end subroutine

public subroutine of_set_priority (integer ai_priority) throws n_ex;IF ai_priority < 0 OR ai_priority > 2 THEN
	gf_throw(PopulateError(1, "¡ La Prioridad debe estar entre 0 y 2 !"))
END IF	

ii_priority = ai_priority

end subroutine

public subroutine of_set_secure_protocol (integer ai_secure_protocol) throws n_ex;IF ii_secure_protocol < 0 or ii_secure_protocol > 4  THEN
	gf_throw(PopulateError(1, "¡ El Protocolo de Seguridad debe estar entre 0 y 4 !"))
END IF	

ii_secure_protocol = ai_secure_protocol

end subroutine

public subroutine of_set_subject (string as_subject) throws n_ex;is_subject = trim(as_subject)

end subroutine

public subroutine of_set_host (string as_host) throws n_ex;If trim(as_host) = "" Then
	gf_throw(PopulateError(1, "¡Hay que indicar un Servidor SMTP!"))
End If

is_host = as_host

end subroutine

public subroutine of_set_sender (string as_sender_name, string as_sender_email) throws n_ex;is_sender_email = of_valid_email(as_sender_email)
is_sender_name = trim(as_sender_name)
end subroutine

private subroutine of_reset () throws n_ex;//Resetear Variables:
ib_enable_tls = TRUE
ii_secure_protocol=ii_ALL
is_encoding= "UTF-8"
ib_html=FALSE
ii_priority = ii_NormalPriority
is_host =""
ii_port = 587
is_user_name=""
is_user_pass=""
is_sender_name=""
is_sender_email=""
is_attachment[] =  is_ResetArray[] 
is_message= ""
is_subject= ""
is_recipient_name[] =  is_ResetArray[] 
is_recipient_email[] =  is_ResetArray[] 
is_cc_name[] =  is_ResetArray[] 
is_cc_email[] =  is_ResetArray[] 
is_bcc_name[] =  is_ResetArray[] 
is_bcc_email[] =  is_ResetArray[] 



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

public subroutine of_enable_tls (boolean ab_enable_tls) throws n_ex;If isnull(ab_enable_tls) Then
	gf_throw(PopulateError(1, "¡Hay que indicar si queremos usar TLS!"))
End If

ib_enable_tls = ab_enable_tls

end subroutine

on n_cst_smtpclient.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_cst_smtpclient.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

event destructor;if isValid(in_Smtp) then Destroy in_Smtp
end event

