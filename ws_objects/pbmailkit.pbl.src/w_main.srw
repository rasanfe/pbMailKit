$PBExportHeader$w_main.srw
forward
global type w_main from window
end type
type cb_cancel from commandbutton within w_main
end type
type tab_main from u_tab_main within w_main
end type
type tab_main from u_tab_main within w_main
end type
end forward

global type w_main from window
integer width = 2574
integer height = 1972
boolean titlebar = true
string title = "PBMAILKIT EMail"
boolean controlmenu = true
long backcolor = 67108864
string icon = "AppIcon!"
boolean center = true
cb_cancel cb_cancel
tab_main tab_main
end type
global w_main w_main

on w_main.create
this.cb_cancel=create cb_cancel
this.tab_main=create tab_main
this.Control[]={this.cb_cancel,&
this.tab_main}
end on

on w_main.destroy
destroy(this.cb_cancel)
destroy(this.tab_main)
end on

type cb_cancel from commandbutton within w_main
integer x = 2158
integer y = 1728
integer width = 334
integer height = 100
integer taborder = 80
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Cancel"
boolean cancel = true
end type

event clicked;Close(Parent)

end event

type tab_main from u_tab_main within w_main
integer x = 37
integer y = 32
integer taborder = 20
end type

