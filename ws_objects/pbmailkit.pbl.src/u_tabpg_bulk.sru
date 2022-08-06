$PBExportHeader$u_tabpg_bulk.sru
forward
global type u_tabpg_bulk from u_tabpg
end type
type st_8 from statictext within u_tabpg_bulk
end type
type st_7 from statictext within u_tabpg_bulk
end type
type st_5 from statictext within u_tabpg_bulk
end type
type sle_recip_email4 from singlelineedit within u_tabpg_bulk
end type
type sle_recip_email3 from singlelineedit within u_tabpg_bulk
end type
type sle_recip_email2 from singlelineedit within u_tabpg_bulk
end type
type cbx_priority from checkbox within u_tabpg_bulk
end type
type cbx_sendhtml from checkbox within u_tabpg_bulk
end type
type st_6 from statictext within u_tabpg_bulk
end type
type sle_recip_email1 from singlelineedit within u_tabpg_bulk
end type
type st_4 from statictext within u_tabpg_bulk
end type
type st_3 from statictext within u_tabpg_bulk
end type
type sle_sender_email from singlelineedit within u_tabpg_bulk
end type
type sle_sender_name from singlelineedit within u_tabpg_bulk
end type
type st_2 from statictext within u_tabpg_bulk
end type
type mle_body from multilineedit within u_tabpg_bulk
end type
type cb_send from commandbutton within u_tabpg_bulk
end type
type sle_subject from singlelineedit within u_tabpg_bulk
end type
type st_1 from statictext within u_tabpg_bulk
end type
end forward

global type u_tabpg_bulk from u_tabpg
string text = "Bulk Email"
st_8 st_8
st_7 st_7
st_5 st_5
sle_recip_email4 sle_recip_email4
sle_recip_email3 sle_recip_email3
sle_recip_email2 sle_recip_email2
cbx_priority cbx_priority
cbx_sendhtml cbx_sendhtml
st_6 st_6
sle_recip_email1 sle_recip_email1
st_4 st_4
st_3 st_3
sle_sender_email sle_sender_email
sle_sender_name sle_sender_name
st_2 st_2
mle_body mle_body
cb_send cb_send
sle_subject sle_subject
st_1 st_1
end type
global u_tabpg_bulk u_tabpg_bulk

type variables
String is_currentdirectory

end variables

on u_tabpg_bulk.create
int iCurrent
call super::create
this.st_8=create st_8
this.st_7=create st_7
this.st_5=create st_5
this.sle_recip_email4=create sle_recip_email4
this.sle_recip_email3=create sle_recip_email3
this.sle_recip_email2=create sle_recip_email2
this.cbx_priority=create cbx_priority
this.cbx_sendhtml=create cbx_sendhtml
this.st_6=create st_6
this.sle_recip_email1=create sle_recip_email1
this.st_4=create st_4
this.st_3=create st_3
this.sle_sender_email=create sle_sender_email
this.sle_sender_name=create sle_sender_name
this.st_2=create st_2
this.mle_body=create mle_body
this.cb_send=create cb_send
this.sle_subject=create sle_subject
this.st_1=create st_1
iCurrent=UpperBound(this.Control)
this.Control[iCurrent+1]=this.st_8
this.Control[iCurrent+2]=this.st_7
this.Control[iCurrent+3]=this.st_5
this.Control[iCurrent+4]=this.sle_recip_email4
this.Control[iCurrent+5]=this.sle_recip_email3
this.Control[iCurrent+6]=this.sle_recip_email2
this.Control[iCurrent+7]=this.cbx_priority
this.Control[iCurrent+8]=this.cbx_sendhtml
this.Control[iCurrent+9]=this.st_6
this.Control[iCurrent+10]=this.sle_recip_email1
this.Control[iCurrent+11]=this.st_4
this.Control[iCurrent+12]=this.st_3
this.Control[iCurrent+13]=this.sle_sender_email
this.Control[iCurrent+14]=this.sle_sender_name
this.Control[iCurrent+15]=this.st_2
this.Control[iCurrent+16]=this.mle_body
this.Control[iCurrent+17]=this.cb_send
this.Control[iCurrent+18]=this.sle_subject
this.Control[iCurrent+19]=this.st_1
end on

on u_tabpg_bulk.destroy
call super::destroy
destroy(this.st_8)
destroy(this.st_7)
destroy(this.st_5)
destroy(this.sle_recip_email4)
destroy(this.sle_recip_email3)
destroy(this.sle_recip_email2)
destroy(this.cbx_priority)
destroy(this.cbx_sendhtml)
destroy(this.st_6)
destroy(this.sle_recip_email1)
destroy(this.st_4)
destroy(this.st_3)
destroy(this.sle_sender_email)
destroy(this.sle_sender_name)
destroy(this.st_2)
destroy(this.mle_body)
destroy(this.cb_send)
destroy(this.sle_subject)
destroy(this.st_1)
end on

event constructor;call super::constructor;// save current directory
is_currentdirectory = GetCurrentDirectory()

sle_sender_email.text = of_getreg("BulkSenderEmail", "")
sle_sender_name.text  = of_getreg("BulkSenderName", "")
sle_recip_email1.text = of_getreg("RecipEmail1", "")
sle_recip_email2.text = of_getreg("RecipEmail2", "")
sle_recip_email3.text = of_getreg("RecipEmail3", "")
sle_recip_email4.text = of_getreg("RecipEmail4", "")
sle_subject.text      = of_getreg("BulkSubject", "")
mle_body.text         = of_getreg("BulkBody", "")

If of_getreg("BulkSendHTML", "N") = "Y" Then
	cbx_sendhtml.checked = True
Else
	cbx_sendhtml.checked = False
End If

If of_getreg("BulkPriority", "None") = "High" Then
	cbx_priority.checked = True
Else
	cbx_priority.checked = False
End If

end event

event destructor;call super::destructor;of_setreg("BulkSenderEmail", sle_sender_email.text)
of_setreg("BulkSenderName",  sle_sender_name.text)
of_setreg("RecipEmail1", sle_recip_email1.text)
of_setreg("RecipEmail2", sle_recip_email2.text)
of_setreg("RecipEmail3", sle_recip_email3.text)
of_setreg("RecipEmail4", sle_recip_email4.text)
of_setreg("BulkSubject", sle_subject.text)
of_setreg("BulkBody", mle_body.text)

If cbx_sendhtml.checked Then
	of_setreg("BulkSendHTML", "Y")
Else
	of_setreg("BulkSendHTML", "N")
End If

If cbx_priority.checked Then
	of_setreg("BulkPriority", "High")
Else
	of_setreg("BulkPriority", "None")
End If

end event

event ue_pagechanged;call super::ue_pagechanged;sle_sender_email.SetFocus()

end event

type st_8 from statictext within u_tabpg_bulk
integer x = 37
integer y = 560
integer width = 288
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "To Email #4:"
boolean focusrectangle = false
end type

type st_7 from statictext within u_tabpg_bulk
integer x = 37
integer y = 432
integer width = 288
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "To Email #3:"
boolean focusrectangle = false
end type

type st_5 from statictext within u_tabpg_bulk
integer x = 37
integer y = 304
integer width = 288
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "To Email #2:"
boolean focusrectangle = false
end type

type sle_recip_email4 from singlelineedit within u_tabpg_bulk
integer x = 329
integer y = 552
integer width = 809
integer height = 84
integer taborder = 60
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_recip_email3 from singlelineedit within u_tabpg_bulk
integer x = 329
integer y = 424
integer width = 809
integer height = 84
integer taborder = 50
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_recip_email2 from singlelineedit within u_tabpg_bulk
integer x = 329
integer y = 296
integer width = 809
integer height = 84
integer taborder = 40
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type cbx_priority from checkbox within u_tabpg_bulk
integer x = 841
integer y = 1360
integer width = 407
integer height = 68
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "High Priority"
end type

type cbx_sendhtml from checkbox within u_tabpg_bulk
integer x = 329
integer y = 1360
integer width = 407
integer height = 68
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Send as HTML"
end type

type st_6 from statictext within u_tabpg_bulk
integer x = 37
integer y = 176
integer width = 288
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "To Email #1:"
boolean focusrectangle = false
end type

type sle_recip_email1 from singlelineedit within u_tabpg_bulk
integer x = 329
integer y = 168
integer width = 809
integer height = 84
integer taborder = 30
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type st_4 from statictext within u_tabpg_bulk
integer x = 37
integer y = 48
integer width = 288
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "From Email:"
boolean focusrectangle = false
end type

type st_3 from statictext within u_tabpg_bulk
integer x = 1230
integer y = 48
integer width = 302
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "From Name:"
boolean focusrectangle = false
end type

type sle_sender_email from singlelineedit within u_tabpg_bulk
integer x = 329
integer y = 40
integer width = 809
integer height = 84
integer taborder = 10
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type sle_sender_name from singlelineedit within u_tabpg_bulk
integer x = 1536
integer y = 40
integer width = 809
integer height = 84
integer taborder = 20
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type st_2 from statictext within u_tabpg_bulk
integer x = 46
integer y = 928
integer width = 242
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Body:"
alignment alignment = right!
boolean focusrectangle = false
end type

type mle_body from multilineedit within u_tabpg_bulk
integer x = 329
integer y = 920
integer width = 2016
integer height = 364
integer taborder = 80
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type cb_send from commandbutton within u_tabpg_bulk
integer x = 1938
integer y = 1344
integer width = 407
integer height = 100
integer taborder = 90
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Send Email"
end type

event clicked;n_pbnismtp ln_smtp

Boolean lb_Html, lb_Priority
Integer li_port, li_authmethod, li_conntype, li_rc, li_idx, li_max
String ls_server, ls_body, ls_userid, ls_passwd, ls_subject
String ls_senderName, ls_senderMail, ls_recipientsMails[], ls_recipientsNames[], ls_charset, ls_errmsg

SetPointer(HourGlass!)

ChangeDirectory(is_currentdirectory)

 ln_smtp  = CREATE n_pbnismtp

// get settings
ls_server     = of_getreg("Server", "")
ls_userid     = of_getreg("Userid", "")
ls_passwd     = of_getreg("Password", "")
li_port       = Integer(of_getreg("Port", "25"))
li_authmethod = Integer(of_getreg("AuthMethod", "2"))
li_conntype   = Integer(of_getreg("ConnType", "0"))
ls_charset    = of_getreg("Charset", "windows-1252")

// input field edits
If ls_server = "" Then
	MessageBox("Edit Error", &
		"You must specify Server on the Settings tab first!", StopSign!)
	Return
End If

If sle_sender_email.text = "" Then
	sle_sender_email.SetFocus()
	MessageBox("Edit Error", &
		"From Email is a required field!", StopSign!)
	Return
End If
If Not ln_smtp.of_ValidEmail(sle_sender_email.text, ls_errmsg) Then
	sle_sender_email.SetFocus()
	MessageBox("From Email Format Error", ls_errmsg, StopSign!)
	Return
End If

If sle_recip_email1.text = "" Then
	sle_recip_email1.SetFocus()
	MessageBox("Edit Error", &
		"To Email #1 is a required field!", StopSign!)
	Return
End If

If sle_subject.text = "" Then
	sle_subject.SetFocus()
	MessageBox("Edit Error", &
		"Subject is a required field!", StopSign!)
	Return
End If

If mle_body.text = "" Then
	mle_body.SetFocus()
	MessageBox("Edit Error", &
		"Body is a required field!", StopSign!)
	Return
End If

// wrap message in HTML
If cbx_sendhtml.Checked Then
	ls_body  = "<html><body bgcolor='#FFFFFF' topmargin=8 leftmargin=8><h2>"
	ls_body += of_replaceall(mle_body.text, "~r~n", "<br>") + "</h2>"
	ls_body += "</body></html>"
	lb_Html = True
Else
	ls_body = mle_body.text
	lb_Html = False
End If

// get other email properties
ls_subject   = sle_subject.text
ls_senderName    = sle_sender_name.text
ls_senderMail = sle_sender_email.text
lb_Priority  = cbx_priority.Checked

// get recipients
If sle_recip_email1.text <> "" Then
	If Not ln_smtp.of_ValidEmail(sle_recip_email1.text, ls_errmsg) Then
		sle_recip_email1.SetFocus()
		MessageBox("To Email #1 Format Error", ls_errmsg, StopSign!)
		Return
	End If
	li_max = UpperBound(ls_recipientsMails[]) + 1
	ls_recipientsNames[li_max] = "" // sle_recip_name1.text
	ls_recipientsMails[li_max] = sle_recip_email1.text
End If
If sle_recip_email2.text <> "" Then
	If Not ln_smtp.of_ValidEmail(sle_recip_email2.text, ls_errmsg) Then
		sle_recip_email1.SetFocus()
		MessageBox("To Email #2 Format Error", ls_errmsg, StopSign!)
		Return
	End If
	li_max = UpperBound(ls_recipientsMails[]) + 1
	ls_recipientsNames[li_max] = "" // sle_recip_name2.text
	ls_recipientsMails[li_max] = sle_recip_email2.text
End If
If sle_recip_email3.text <> "" Then
	If Not ln_smtp.of_ValidEmail(sle_recip_email3.text, ls_errmsg) Then
		sle_recip_email1.SetFocus()
		MessageBox("To Email #3 Format Error", ls_errmsg, StopSign!)
		Return
	End If
	li_max = UpperBound(ls_recipientsMails[]) + 1
	ls_recipientsNames[li_max] = "" // sle_recip_name3.text
	ls_recipientsMails[li_max] = sle_recip_email3.text
End If
If sle_recip_email4.text <> "" Then
	If Not ln_smtp.of_ValidEmail(sle_recip_email4.text, ls_errmsg) Then
		sle_recip_email1.SetFocus()
		MessageBox("To Email #4 Format Error", ls_errmsg, StopSign!)
		Return
	End If
	li_max = UpperBound(ls_recipientsMails[]) + 1
	ls_recipientsNames[li_max] = "" // sle_recip_name4.text
	ls_recipientsMails[li_max] = sle_recip_email4.text
End If

SetPointer(HourGlass!)

// connect to the server
try
	ln_smtp.of_SetMailerName("PBMailkit 1.0") //Modificación Topwiz
	// set server settings
	ln_smtp.of_SetSMTPServer(ls_server)
	If ls_userid <> "" Then
		ln_smtp.of_SetUserNamePassword(ls_userid, ls_passwd)
	End If
	ln_smtp.of_SetPort(li_port)
	ln_smtp.of_SetAuthMethod(li_AuthMethod)
	ln_smtp.of_SetConnectionType(li_ConnType)
	ln_smtp.of_SetCharSet(ls_Charset)
	If lb_Priority Then
		ln_smtp.of_SetPriority(ln_smtp.HighPriority)
	End If
	// connect to the server
	li_rc = ln_smtp.of_SmtpConnect()
catch ( NullObjectError noe1 )
	MessageBox("SmtpConnect: Null Object Exception", &
					noe1.getMessage(), StopSign!)
catch ( PBXRuntimeError pbxre1 )
	MessageBox("SmtpConnect: PBX Exception", &
					pbxre1.getMessage(), StopSign!)
catch ( Throwable oe1 )
	MessageBox("SmtpConnect: Other Exception", &
					oe1.getMessage(), StopSign!)
finally
	If li_rc = 1 Then
		// Success
	Else
		ls_errmsg = ln_smtp.of_GetLastErrorMessage()
		MessageBox("SmtpConnect Error: " + String(li_rc), &
						ls_errmsg, StopSign!)
	End If
end try

// send the emails
For li_idx = 1 To li_max
	try
		// set message properties
		ln_smtp.of_SetSenderEmail(ls_senderName, ls_senderMail)
		ln_smtp.of_SetRecipientEmail(ls_recipientsNames[li_idx], ls_recipientsMails[li_idx])
		ln_smtp.of_SetSubject(ls_subject)
		ln_smtp.of_SetMessage(ls_body, lb_Html)
		// send the email
		li_rc = ln_smtp.of_SmtpSend()
	catch ( NullObjectError noe2 )
		MessageBox("SmtpSend: Null Object Exception", &
						noe2.getMessage(), StopSign!)
	catch ( PBXRuntimeError pbxre2 )
		MessageBox("SmtpSend: PBX Exception", &
						pbxre2.getMessage(), StopSign!)
	catch ( Throwable oe2 )
		MessageBox("SmtpSend: Other Exception", &
						oe2.getMessage(), StopSign!)
	finally
		If li_rc = 1 Then
			// Success
		Else
			ls_errmsg = ln_smtp.of_GetLastErrorMessage()
			MessageBox("SmtpSend Error: " + String(li_rc), &
							ls_errmsg, StopSign!)
			Exit	// Break out of loop
		End If
	end try
Next

// disconnect from the server
try
	li_rc = ln_smtp.of_SmtpDisconnect()
catch ( NullObjectError noe3 )
	MessageBox("SmtpDisconnect: Null Object Exception", &
					noe3.getMessage(), StopSign!)
catch ( PBXRuntimeError pbxre3 )
	MessageBox("SmtpDisconnect: PBX Exception", &
					pbxre3.getMessage(), StopSign!)
catch ( Throwable oe3 )
	MessageBox("SmtpDisconnect: Other Exception", &
					oe3.getMessage(), StopSign!)
finally
	If li_rc = 1 Then
		// Success
	Else
		ls_errmsg = ln_smtp.of_GetLastErrorMessage()
		MessageBox("SmtpDisconnect Error: " + String(li_rc), &
						ls_errmsg, StopSign!)
	End If
end try

MessageBox(this.text, "Message Sent!")
Destroy  ln_smtp
end event

type sle_subject from singlelineedit within u_tabpg_bulk
integer x = 329
integer y = 800
integer width = 2016
integer height = 84
integer taborder = 70
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
borderstyle borderstyle = stylelowered!
end type

type st_1 from statictext within u_tabpg_bulk
integer x = 46
integer y = 808
integer width = 242
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Subject:"
alignment alignment = right!
boolean focusrectangle = false
end type

