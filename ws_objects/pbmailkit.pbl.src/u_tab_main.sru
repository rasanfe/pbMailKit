$PBExportHeader$u_tab_main.sru
forward
global type u_tab_main from u_tab
end type
type tabpage_smtp from u_tabpg_smtp within u_tab_main
end type
type tabpage_smtp from u_tabpg_smtp within u_tab_main
end type
type tabpage_bulk from u_tabpg_bulk within u_tab_main
end type
type tabpage_bulk from u_tabpg_bulk within u_tab_main
end type
type tabpage_settings from u_tabpg_settings within u_tab_main
end type
type tabpage_settings from u_tabpg_settings within u_tab_main
end type
end forward

global type u_tab_main from u_tab
integer width = 2455
integer height = 1636
tabpage_smtp tabpage_smtp
tabpage_bulk tabpage_bulk
tabpage_settings tabpage_settings
end type
global u_tab_main u_tab_main

on u_tab_main.create
this.tabpage_smtp=create tabpage_smtp
this.tabpage_bulk=create tabpage_bulk
this.tabpage_settings=create tabpage_settings
int iCurrent
call super::create
iCurrent=UpperBound(this.Control)
this.Control[iCurrent+1]=this.tabpage_smtp
this.Control[iCurrent+2]=this.tabpage_bulk
this.Control[iCurrent+3]=this.tabpage_settings
end on

on u_tab_main.destroy
call super::destroy
destroy(this.tabpage_smtp)
destroy(this.tabpage_bulk)
destroy(this.tabpage_settings)
end on

type tabpage_smtp from u_tabpg_smtp within u_tab_main
integer x = 18
integer y = 100
integer width = 2418
integer height = 1520
end type

type tabpage_bulk from u_tabpg_bulk within u_tab_main
integer x = 18
integer y = 100
integer width = 2418
integer height = 1520
end type

type tabpage_settings from u_tabpg_settings within u_tab_main
integer x = 18
integer y = 100
integer width = 2418
integer height = 1520
end type

