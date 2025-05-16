forward
global type w_main from window
end type
type st_info from statictext within w_main
end type
type p_logo from picture within w_main
end type
type st_myversion from statictext within w_main
end type
type st_platform from statictext within w_main
end type
type cb_cancel from commandbutton within w_main
end type
type tab_main from u_tab_main within w_main
end type
type tab_main from u_tab_main within w_main
end type
type r_1 from rectangle within w_main
end type
end forward

global type w_main from window
integer width = 2574
integer height = 2192
boolean titlebar = true
string title = "PBMAILKIT EMail"
boolean controlmenu = true
string icon = "AppIcon!"
boolean center = true
st_info st_info
p_logo p_logo
st_myversion st_myversion
st_platform st_platform
cb_cancel cb_cancel
tab_main tab_main
r_1 r_1
end type
global w_main w_main

forward prototypes
public subroutine wf_version (statictext ast_version, statictext ast_patform)
end prototypes

public subroutine wf_version (statictext ast_version, statictext ast_patform);String ls_version, ls_platform
environment env
integer rtn

rtn = GetEnvironment(env)

IF rtn <> 1 THEN 
	ls_version = string(year(today()))
	ls_platform="32"
ELSE
	ls_version = "20"+ string(env.pbmajorrevision)+ "." + string(env.pbbuildnumber)
	ls_platform= string(env.ProcessBitness)
END IF

ls_platform += " Bits"

ast_version.text=ls_version
ast_patform.text=ls_platform
end subroutine

on w_main.create
this.st_info=create st_info
this.p_logo=create p_logo
this.st_myversion=create st_myversion
this.st_platform=create st_platform
this.cb_cancel=create cb_cancel
this.tab_main=create tab_main
this.r_1=create r_1
this.Control[]={this.st_info,&
this.p_logo,&
this.st_myversion,&
this.st_platform,&
this.cb_cancel,&
this.tab_main,&
this.r_1}
end on

on w_main.destroy
destroy(this.st_info)
destroy(this.p_logo)
destroy(this.st_myversion)
destroy(this.st_platform)
destroy(this.cb_cancel)
destroy(this.tab_main)
destroy(this.r_1)
end on

event open;wf_version(st_myversion, st_platform)
end event

type st_info from statictext within w_main
integer x = 18
integer y = 2032
integer width = 1353
integer height = 64
integer textsize = -7
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 8421504
string text = "Copyright © Ramón San Félix Ramón  rsrsystem.soft@gmail.com"
alignment alignment = right!
boolean focusrectangle = false
end type

type p_logo from picture within w_main
integer x = 5
integer y = 4
integer width = 1253
integer height = 248
boolean originalsize = true
string picturename = "logo.jpg"
boolean focusrectangle = false
end type

type st_myversion from statictext within w_main
integer x = 2121
integer y = 56
integer width = 402
integer height = 64
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Versión"
alignment alignment = right!
boolean focusrectangle = false
end type

type st_platform from statictext within w_main
integer x = 2121
integer y = 144
integer width = 402
integer height = 64
integer textsize = -12
integer weight = 400
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Arial"
long textcolor = 16777215
long backcolor = 33521664
string text = "Bits"
alignment alignment = right!
boolean focusrectangle = false
end type

type cb_cancel from commandbutton within w_main
integer x = 2158
integer y = 1992
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
integer y = 296
integer taborder = 20
end type

type r_1 from rectangle within w_main
long linecolor = 33554432
linestyle linestyle = transparent!
integer linethickness = 4
long fillcolor = 33521664
integer width = 2551
integer height = 260
end type

