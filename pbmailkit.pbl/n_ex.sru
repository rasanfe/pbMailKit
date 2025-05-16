forward
global type n_ex from exception
end type
end forward

global type n_ex from exception
end type
global n_ex n_ex

type variables
/**********************************************************************************************************************
How to use:	http://code.intfast.ca/viewtopic.php?t=1
***********************************************************************************************************************
Licence:		No licence needed, this utility is absolutely free
***********************************************************************************************************************
Developer:	Michael Zuskin > http://linkedin.com/in/zuskin | http://code.intfast.ca/
**********************************************************************************************************************/

protected:

	int		ii_err_num
	int		ii_line
	string	is_class
	string	is_script
end variables

forward prototypes
public subroutine of_populate (integer ai_err_num, string as_err_msg, string as_class, string as_script, integer ai_line)
public subroutine of_msg ()
protected subroutine of_write_to_log ()
public function string of_get_error ()
public function integer of_get_number ()
end prototypes

public subroutine of_populate (integer ai_err_num, string as_err_msg, string as_class, string as_script, integer ai_line);/**********************************************************************************************************************
Dscr:	Writes into this object all the information related to the class and the script the exception is thrown from.
		Called from f_throw().
		Later, that information will be extracted from this object by of_msg() and, optionally, of_write_to_log().
***********************************************************************************************************************
Arg:	int		ai_err_num
		string	as_err_msg
		string	as_class
		string	as_script
		int		ai_line
Change:		Ramón San Félix Ramón  16-11-2022  Renombro funión uf_populate a of_populate
**********************************************************************************************************************/
this.ii_err_num = ai_err_num
this.SetMessage(as_err_msg)
this.is_class = as_class
this.is_script = as_script
this.ii_line = ai_line

return
end subroutine

public subroutine of_msg ();/**********************************************************************************************************************
Dscr:	Displays a message with all the exception-related information.
		That information has been previously written to this object by f_throw(), which called of_populate().
		of_msg() should be called when the exception is caught in "catch" section of a "try...end try" block:
		
		try
			of_function_which_throws_n_ex_using_f_throw() // an exception can be thrown...
			if ... then f_throw(PopulateError(0, "Oooops...")) // another exception can be thrown...
		catch(n_ex e)
			e.of_msg()
		end try
Change:		Ramón San Félix Ramón  16-11-2022  Renombro funión uf_msg a of_msg
**********************************************************************************************************************/
string	ls_msg
string	ls_msg_title
string ls_error

ls_msg_title = gf_iif(this.ii_err_num > 0, "EXCEPTION #" + String(this.ii_err_num) + " THROWN", "EXCEPTION THROWN")

if Len(this.is_class) > 0 then ls_msg = "Class: " + this.is_class + "~r~n"
if Len(this.is_script) > 0 then ls_msg += "Script: " + this.is_script + "~r~n"
if this.ii_line > 0 then ls_msg += "Line: " + String(ii_line) + "~r~n"
if ls_msg <> "" then ls_msg += "~r~n~r~n"

ls_error = this.GetMessage() // that message has previously been written to this object by f_throw()
ls_msg += ls_error

this.of_write_to_log()

// If your allplication displays error messages using it's own mechanism rather than calls MessageBox()
// each time an error occurs, then change the following line accordingly:
MessageBox(ls_msg_title, ls_msg, Exclamation!, OK!, 1)
	
// If you use PFC, then uncomment the following line and comment out the previous line:
//gnv_app.inv_error.of_Message(ls_msg_title, ls_msg)

return
end subroutine

protected subroutine of_write_to_log ();/**********************************************************************************************************************
Dscr:	Registers the thrown exception (writes to a log table or a file, and/or sends an email to the support staff
		or a developer) in addition to the error message, displayed in of_msg.
		Called from of_msg() just before displaying the error message.
Change:		Ramón San Félix Ramón  16-11-2022  Renombro funión uf_write_to_log a of_write_to_log
**********************************************************************************************************************/

//####### If you need to somehow register the thrown exception (to write to a log table or a file, and/or send
//####### an email to the support staff or a developer) in addition to the error message, displayed in of_msg,
//####### then uncomment the following fragment and customize it according to your needs:

//string	ls_msg
//
//if this.ii_err_num > 0 then
//	ls_msg = "#" + String(this.ii_err_num) + ": "
//end if
//
//if Len(this.is_class) > 0 then ls_msg += "Class: " + this.is_class
//if ls_msg <> "" then ls_msg += "; "
//
//if Len(this.is_script) > 0 then ls_msg += "Script: " + this.is_script
//if ls_msg <> "" then ls_msg += "; "
//
//if this.ii_line > 0 then ls_msg += "Line: " + String(this.ii_line)
//if ls_msg <> "" then ls_msg += "; "
//
//ls_msg += this.GetMessage()
//
///******* TODO: Write ls_msg to log and/or send it by email to the support staff or a developer *******/
//
//return


end subroutine

public function string of_get_error ();string ls_error

ls_error = this.GetMessage() // that message has previously been written to this object by f_throw()

RETURN ls_error

end function

public function integer of_get_number ();integer li_error

li_error = ii_err_num

RETURN li_error

end function

on n_ex.create
call super::create
TriggerEvent( this, "constructor" )
end on

on n_ex.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

