$PBExportHeader$u_tab.sru
$PBExportComments$Base tab object
forward
global type u_tab from tab
end type
end forward

global type u_tab from tab
integer width = 1582
integer height = 1000
integer taborder = 10
integer textsize = -8
integer weight = 400
fontcharset fontcharset = ansi!
fontpitch fontpitch = variable!
fontfamily fontfamily = swiss!
string facename = "Tahoma"
long backcolor = 79416533
boolean raggedright = true
boolean boldselectedtext = true
integer selectedtab = 1
end type
global u_tab u_tab

event selectionchanged;// trigger event on new tab page
If newindex > 0 Then
	Control[newindex].Event Dynamic ue_pagechanged(oldindex)
End If

end event

