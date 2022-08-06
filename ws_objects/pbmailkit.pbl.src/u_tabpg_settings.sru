$PBExportHeader$u_tabpg_settings.sru
forward
global type u_tabpg_settings from u_tabpg
end type
type dw_settings from datawindow within u_tabpg_settings
end type
type cb_save from commandbutton within u_tabpg_settings
end type
end forward

global type u_tabpg_settings from u_tabpg
string text = "Server Settings"
dw_settings dw_settings
cb_save cb_save
end type
global u_tabpg_settings u_tabpg_settings

on u_tabpg_settings.create
int iCurrent
call super::create
this.dw_settings=create dw_settings
this.cb_save=create cb_save
iCurrent=UpperBound(this.Control)
this.Control[iCurrent+1]=this.dw_settings
this.Control[iCurrent+2]=this.cb_save
end on

on u_tabpg_settings.destroy
call super::destroy
destroy(this.dw_settings)
destroy(this.cb_save)
end on

event constructor;call super::constructor;// initialize the DataWindow
dw_settings.Reset()
dw_settings.InsertRow(0)

dw_settings.SetItem(1, "server", of_getreg("Server", ""))
dw_settings.SetItem(1, "userid", of_getreg("Userid", ""))
dw_settings.SetItem(1, "password", of_getreg("Password", ""))
dw_settings.SetItem(1, "port", Long(of_getreg("Port", "25")))
dw_settings.SetItem(1, "conntype", Long(of_getreg("ConnType", "1")))
dw_settings.SetItem(1, "authmethod", Long(of_getreg("AuthMethod", "5")))
dw_settings.SetItem(1, "characterset", of_getreg("Charset", "utf-8"))

end event

event ue_pagechanged;call super::ue_pagechanged;dw_settings.SetFocus()

end event

type dw_settings from datawindow within u_tabpg_settings
integer x = 73
integer y = 64
integer width = 2199
integer height = 1244
integer taborder = 20
string title = "none"
string dataobject = "d_settings"
boolean border = false
boolean livescroll = true
end type

type cb_save from commandbutton within u_tabpg_settings
integer x = 1938
integer y = 1344
integer width = 407
integer height = 100
integer taborder = 80
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
string text = "Save Settings"
end type

event clicked;dw_settings.AcceptText()

of_setreg("Server", dw_settings.GetItemString(1, "server"))
of_setreg("Userid", dw_settings.GetItemString(1, "userid"))
of_setreg("Password", dw_settings.GetItemString(1, "password"))
of_setreg("Port", String(dw_settings.GetItemNumber(1, "port")))
of_setreg("ConnType", String(dw_settings.GetItemNumber(1, "conntype")))
of_setreg("AuthMethod", String(dw_settings.GetItemNumber(1, "authmethod")))
of_setreg("Charset", dw_settings.GetItemString(1, "characterset"))

end event

