$PBExportHeader$u_tabpg.sru
$PBExportComments$Base tabpage object
forward
global type u_tabpg from userobject
end type
end forward

global type u_tabpg from userobject
integer width = 2414
integer height = 1504
long backcolor = 79416533
long tabtextcolor = 33554432
long picturemaskcolor = 536870912
event ue_pagechanged ( integer oldindex )
end type
global u_tabpg u_tabpg

type prototypes
Subroutine DebugMsg( &
	String lpOutputString &
	) Library "kernel32.dll" Alias For "OutputDebugStringW"

end prototypes

forward prototypes
public function string of_getreg (string as_entry, string as_default)
public subroutine of_setreg (string as_entry, string as_value)
public function string of_replaceall (string as_oldstring, string as_findstr, string as_replace)
end prototypes

public function string of_getreg (string as_entry, string as_default);String ls_regkey, ls_regvalue

ls_regkey = "HKEY_CURRENT_USER\Software\rsrsystem\pbmailkit"

RegistryGet(ls_regkey, as_entry, ls_regvalue)
If ls_regvalue = "" Then
	ls_regvalue = as_default
End If

Return ls_regvalue

end function

public subroutine of_setreg (string as_entry, string as_value);String ls_regkey

ls_regkey = "HKEY_CURRENT_USER\Software\rsrsystem\pbmailkit"

RegistrySet(ls_regkey, as_entry, as_value)

end subroutine

public function string of_replaceall (string as_oldstring, string as_findstr, string as_replace);String ls_newstring
Long ll_findstr, ll_replace, ll_pos

// get length of strings
ll_findstr = Len(as_findstr)
ll_replace = Len(as_replace)

// find first occurrence
ls_newstring = as_oldstring
ll_pos = Pos(ls_newstring, as_findstr)

Do While ll_pos > 0
	// replace old with new
	ls_newstring = Replace(ls_newstring, ll_pos, ll_findstr, as_replace)
	// find next occurrence
	ll_pos = Pos(ls_newstring, as_findstr, (ll_pos + ll_replace))
Loop

Return ls_newstring

end function

on u_tabpg.create
end on

on u_tabpg.destroy
end on

