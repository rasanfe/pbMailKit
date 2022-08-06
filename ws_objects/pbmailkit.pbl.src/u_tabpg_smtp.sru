$PBExportHeader$u_tabpg_smtp.sru
forward
global type u_tabpg_smtp from u_tabpg
end type
type cbx_receipt from checkbox within u_tabpg_smtp
end type
type sle_cc_name from singlelineedit within u_tabpg_smtp
end type
type st_9 from statictext within u_tabpg_smtp
end type
type sle_cc_email from singlelineedit within u_tabpg_smtp
end type
type st_8 from statictext within u_tabpg_smtp
end type
type cbx_priority from checkbox within u_tabpg_smtp
end type
type cb_del from commandbutton within u_tabpg_smtp
end type
type cb_addfile from commandbutton within u_tabpg_smtp
end type
type st_7 from statictext within u_tabpg_smtp
end type
type lb_attachments from listbox within u_tabpg_smtp
end type
type cbx_sendhtml from checkbox within u_tabpg_smtp
end type
type st_6 from statictext within u_tabpg_smtp
end type
type sle_recip_email from singlelineedit within u_tabpg_smtp
end type
type st_4 from statictext within u_tabpg_smtp
end type
type st_3 from statictext within u_tabpg_smtp
end type
type sle_sender_email from singlelineedit within u_tabpg_smtp
end type
type sle_sender_name from singlelineedit within u_tabpg_smtp
end type
type st_2 from statictext within u_tabpg_smtp
end type
type mle_body from multilineedit within u_tabpg_smtp
end type
type cb_send from commandbutton within u_tabpg_smtp
end type
type sle_subject from singlelineedit within u_tabpg_smtp
end type
type st_1 from statictext within u_tabpg_smtp
end type
type sle_recip_name from singlelineedit within u_tabpg_smtp
end type
type st_5 from statictext within u_tabpg_smtp
end type
end forward

global type u_tabpg_smtp from u_tabpg
string text = "Send Email"
cbx_receipt cbx_receipt
sle_cc_name sle_cc_name
st_9 st_9
sle_cc_email sle_cc_email
st_8 st_8
cbx_priority cbx_priority
cb_del cb_del
cb_addfile cb_addfile
st_7 st_7
lb_attachments lb_attachments
cbx_sendhtml cbx_sendhtml
st_6 st_6
sle_recip_email sle_recip_email
st_4 st_4
st_3 st_3
sle_sender_email sle_sender_email
sle_sender_name sle_sender_name
st_2 st_2
mle_body mle_body
cb_send cb_send
sle_subject sle_subject
st_1 st_1
sle_recip_name sle_recip_name
st_5 st_5
end type
global u_tabpg_smtp u_tabpg_smtp

type variables
String is_currentdirectory

end variables

on u_tabpg_smtp.create
int iCurrent
call super::create
this.cbx_receipt=create cbx_receipt
this.sle_cc_name=create sle_cc_name
this.st_9=create st_9
this.sle_cc_email=create sle_cc_email
this.st_8=create st_8
this.cbx_priority=create cbx_priority
this.cb_del=create cb_del
this.cb_addfile=create cb_addfile
this.st_7=create st_7
this.lb_attachments=create lb_attachments
this.cbx_sendhtml=create cbx_sendhtml
this.st_6=create st_6
this.sle_recip_email=create sle_recip_email
this.st_4=create st_4
this.st_3=create st_3
this.sle_sender_email=create sle_sender_email
this.sle_sender_name=create sle_sender_name
this.st_2=create st_2
this.mle_body=create mle_body
this.cb_send=create cb_send
this.sle_subject=create sle_subject
this.st_1=create st_1
this.sle_recip_name=create sle_recip_name
this.st_5=create st_5
iCurrent=UpperBound(this.Control)
this.Control[iCurrent+1]=this.cbx_receipt
this.Control[iCurrent+2]=this.sle_cc_name
this.Control[iCurrent+3]=this.st_9
this.Control[iCurrent+4]=this.sle_cc_email
this.Control[iCurrent+5]=this.st_8
this.Control[iCurrent+6]=this.cbx_priority
this.Control[iCurrent+7]=this.cb_del
this.Control[iCurrent+8]=this.cb_addfile
this.Control[iCurrent+9]=this.st_7
this.Control[iCurrent+10]=this.lb_attachments
this.Control[iCurrent+11]=this.cbx_sendhtml
this.Control[iCurrent+12]=this.st_6
this.Control[iCurrent+13]=this.sle_recip_email
this.Control[iCurrent+14]=this.st_4
this.Control[iCurrent+15]=this.st_3
this.Control[iCurrent+16]=this.sle_sender_email
this.Control[iCurrent+17]=this.sle_sender_name
this.Control[iCurrent+18]=this.st_2
this.Control[iCurrent+19]=this.mle_body
this.Control[iCurrent+20]=this.cb_send
this.Control[iCurrent+21]=this.sle_subject
this.Control[iCurrent+22]=this.st_1
this.Control[iCurrent+23]=this.sle_recip_name
this.Control[iCurrent+24]=this.st_5
end on

on u_tabpg_smtp.destroy
call super::destroy
destroy(this.cbx_receipt)
destroy(this.sle_cc_name)
destroy(this.st_9)
destroy(this.sle_cc_email)
destroy(this.st_8)
destroy(this.cbx_priority)
destroy(this.cb_del)
destroy(this.cb_addfile)
destroy(this.st_7)
destroy(this.lb_attachments)
destroy(this.cbx_sendhtml)
destroy(this.st_6)
destroy(this.sle_recip_email)
destroy(this.st_4)
destroy(this.st_3)
destroy(this.sle_sender_email)
destroy(this.sle_sender_name)
destroy(this.st_2)
destroy(this.mle_body)
destroy(this.cb_send)
destroy(this.sle_subject)
destroy(this.st_1)
destroy(this.sle_recip_name)
destroy(this.st_5)
end on

event constructor;call super::constructor;// save current directory
is_currentdirectory = GetCurrentDirectory()

sle_sender_email.text = of_getreg("SenderEmail", "")
sle_sender_name.text  = of_getreg("SenderName", "")
sle_recip_email.text  = of_getreg("RecipEmail", "")
sle_recip_name.text   = of_getreg("RecipName", "")
sle_subject.text      = of_getreg("Subject", "")
mle_body.text         = of_getreg("Body", "")

If of_getreg("SendHTML", "N") = "Y" Then
	cbx_sendhtml.checked = True
Else
	cbx_sendhtml.checked = False
End If

If of_getreg("Priority", "None") = "High" Then
	cbx_priority.checked = True
Else
	cbx_priority.checked = False
End If

If of_getreg("Receipt", "N") = "Y" Then
	cbx_receipt.checked = True
Else
	cbx_receipt.checked = False
End If

end event

event destructor;call super::destructor;of_setreg("SenderEmail", sle_sender_email.text)
of_setreg("SenderName", sle_sender_name.text)
of_setreg("RecipEmail", sle_recip_email.text)
of_setreg("RecipName", sle_recip_name.text)
of_setreg("Subject", sle_subject.text)
of_setreg("Body", mle_body.text)

If cbx_sendhtml.checked Then
	of_setreg("SendHTML", "Y")
Else
	of_setreg("SendHTML", "N")
End If

If cbx_priority.checked Then
	of_setreg("Priority", "High")
Else
	of_setreg("Priority", "None")
End If

If cbx_receipt.checked Then
	of_setreg("Receipt", "Y")
Else
	of_setreg("Receipt", "N")
End If

end event

event ue_pagechanged;call super::ue_pagechanged;sle_sender_email.SetFocus()

end event

type cbx_receipt from checkbox within u_tabpg_smtp
integer x = 1280
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
string text = "Read Receipt"
end type

type sle_cc_name from singlelineedit within u_tabpg_smtp
integer x = 1536
integer y = 296
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

type st_9 from statictext within u_tabpg_smtp
integer x = 1285
integer y = 304
integer width = 247
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "CC Name:"
boolean focusrectangle = false
end type

type sle_cc_email from singlelineedit within u_tabpg_smtp
integer x = 329
integer y = 296
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

type st_8 from statictext within u_tabpg_smtp
integer x = 82
integer y = 304
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
string text = "CC Email:"
boolean focusrectangle = false
end type

type cbx_priority from checkbox within u_tabpg_smtp
integer x = 805
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

type cb_del from commandbutton within u_tabpg_smtp
integer x = 37
integer y = 1120
integer width = 261
integer height = 100
integer taborder = 110
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Delete"
end type

event clicked;// delete attachment

Integer li_row

li_row = lb_attachments.SelectedIndex()
If li_row > 0 Then
	lb_attachments.DeleteItem(li_row)
End If

end event

type cb_addfile from commandbutton within u_tabpg_smtp
integer x = 37
integer y = 992
integer width = 261
integer height = 100
integer taborder = 100
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Add"
end type

event clicked;String ls_pathname, ls_filename, ls_filter
Integer li_rc

ls_filter = "All files,*.*"

li_rc = GetFileOpenName("Select File to Attach", &
		ls_pathname, ls_filename, "", ls_filter)

If li_rc = 1 Then
	lb_attachments.AddItem(ls_pathname)
End If

end event

type st_7 from statictext within u_tabpg_smtp
integer x = 329
integer y = 928
integer width = 315
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "Attachments:"
boolean focusrectangle = false
end type

type lb_attachments from listbox within u_tabpg_smtp
integer x = 329
integer y = 992
integer width = 2016
integer height = 292
integer taborder = 90
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
boolean vscrollbar = true
boolean sorted = false
borderstyle borderstyle = stylelowered!
end type

type cbx_sendhtml from checkbox within u_tabpg_smtp
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

type st_6 from statictext within u_tabpg_smtp
integer x = 91
integer y = 180
integer width = 233
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "To Email:"
boolean focusrectangle = false
end type

type sle_recip_email from singlelineedit within u_tabpg_smtp
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

type st_4 from statictext within u_tabpg_smtp
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

type st_3 from statictext within u_tabpg_smtp
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

type sle_sender_email from singlelineedit within u_tabpg_smtp
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

type sle_sender_name from singlelineedit within u_tabpg_smtp
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

type st_2 from statictext within u_tabpg_smtp
integer x = 46
integer y = 544
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

type mle_body from multilineedit within u_tabpg_smtp
integer x = 329
integer y = 536
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

type cb_send from commandbutton within u_tabpg_smtp
integer x = 1938
integer y = 1344
integer width = 407
integer height = 100
integer taborder = 120
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Send Email"
end type

event clicked;n_pbnismtp ln_smtp 
Boolean lb_Html, lb_Priority, lb_Receipt
Integer li_port, li_idx, li_max, li_authmethod, li_conntype, li_rc
String ls_server, ls_body, ls_userid, ls_passwd, ls_subject, ls_attach
String ls_charset, ls_errmsg, ls_work
String ls_senderName, ls_recipientName, ls_ccrecipName
String ls_senderMail, ls_recipientMail, ls_ccrecipMail

SetPointer(HourGlass!)

ChangeDirectory(is_currentdirectory)

 ln_smtp  = CREATE n_pbnismtp

// get settings
ls_server     = of_getreg("Server", "")
ls_userid     = of_getreg("Userid", "")
ls_passwd     = of_getreg("Password", "6qk#6dfd[2b:")
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

If sle_recip_email.text = "" Then
	sle_recip_email.SetFocus()
	MessageBox("Edit Error", &
		"To Email is a required field!", StopSign!)
	Return
End If
If Not ln_smtp.of_ValidEmail(sle_recip_email.text, ls_errmsg) Then
	sle_recip_email.SetFocus()
	MessageBox("To Email Format Error", ls_errmsg, StopSign!)
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

// build attachment list 
//li_max = lb_attachments.TotalItems()
//For li_idx = 1 To li_max
//	ln_smtp.of_AddToString(ls_attach, lb_attachments.Text(li_idx))
//Next

// wrap message in HTML
If cbx_sendhtml.Checked Then
	ls_body  = "<html><body bgcolor='#FFFFFF' topmargin=8 leftmargin=8><h2>"
	ls_body += of_replaceall(mle_body.text, "~r~n", "<br>") + "</h2>"
	For li_idx = 1 To li_max
		ls_work = lb_attachments.Text(li_idx)
		ls_work = Mid(ls_work, LastPos(ls_work, "\") + 1)
		ls_body += "<br><h2>" + ls_work + "</h2><br>"
		ls_body += "<img src='cid:attachment" + String(li_idx) + "'/>"
	Next
	ls_body += "</body></html>"
	lb_Html = True
Else
	ls_body = mle_body.text
	lb_Html = False
End If

// get other email properties
ls_subject   = sle_subject.text
ls_recipientName = sle_recip_name.text
ls_recipientMail = sle_recip_email.text
ls_senderName    = sle_sender_name.text
ls_senderMail    = sle_sender_email.text
ls_ccrecipName   = sle_cc_name.text
ls_ccrecipMail   = sle_cc_email.text
lb_Priority  = cbx_priority.Checked
lb_Receipt   = cbx_receipt.Checked

// send the email
try
	SetPointer(HourGlass!)

	// set server settings
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

	// set message properties
	ln_smtp.of_SetSenderEmail(ls_senderName, ls_senderMail)
	ln_smtp.of_SetRecipientEmail(ls_recipientName, ls_recipientMail)
	If ls_ccrecipMail <> "" Then
		ln_smtp.of_SetCCRecipientEmail(ls_ccrecipName, ls_ccrecipMail)
	End If
	ln_smtp.of_SetSubject(ls_subject)
	ln_smtp.of_SetMessage(ls_body, lb_Html)
	If lb_Priority Then
		ln_smtp.of_SetPriority(ln_smtp.HighPriority)
	End If
	ln_smtp.of_SetReadReceiptRequested(lb_Receipt)
	
	
//	If ls_attach <> "" Then
//		ln_smtp.of_SetAttachment(ls_attach)
//	End If

	//Set Atachments
	li_max = lb_attachments.TotalItems()
	For li_idx = 1 To li_max
		ls_attach = lb_attachments.Text(li_idx)
		ln_smtp.of_SetAttachment(ls_attach)
	Next

	// send the email
	li_rc = ln_smtp.of_Send()

catch ( NullObjectError noe )
	MessageBox("Null Object Exception", noe.getMessage(), StopSign!)
catch ( PBXRuntimeError pbxre )
	MessageBox("PBX Exception", pbxre.getMessage(), StopSign!)
catch ( Throwable oe )
	MessageBox("Other Exception", oe.getMessage(), StopSign!)
finally
	If li_rc = 1 Then
		MessageBox(this.text, "Message Sent!")
	Else
		ls_errmsg = ln_smtp.of_GetLastErrorMessage()
		MessageBox(this.text + " Error: " + String(li_rc), &
						ls_errmsg, StopSign!)
	End If
end try

Destroy  ln_smtp
end event

type sle_subject from singlelineedit within u_tabpg_smtp
integer x = 329
integer y = 416
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

type st_1 from statictext within u_tabpg_smtp
integer x = 46
integer y = 424
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

type sle_recip_name from singlelineedit within u_tabpg_smtp
integer x = 1536
integer y = 168
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

type st_5 from statictext within u_tabpg_smtp
integer x = 1285
integer y = 176
integer width = 247
integer height = 64
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long textcolor = 33554432
long backcolor = 67108864
string text = "To Name:"
boolean focusrectangle = false
end type

