$PBExportHeader$nvo_mailkitsmptrsr.sru
forward
global type nvo_mailkitsmptrsr from dotnetobject
end type
end forward

global type nvo_mailkitsmptrsr from dotnetobject
event ue_error ( )
end type
global nvo_mailkitsmptrsr nvo_mailkitsmptrsr

type variables

PUBLIC:
String is_assemblypath = "MailKitNetSmtp.dll"
String is_classname = "MailKitNetSmtp.MailKitSmptRSR"

/* Exception handling -- Indicates how proxy handles .NET exceptions */
Boolean ib_CrashOnException = False

/*      Error types       */
Constant Int SUCCESS        =  0 // No error since latest reset
Constant Int LOAD_FAILURE   = -1 // Failed to load assembly
Constant Int CREATE_FAILURE = -2 // Failed to create .NET object
Constant Int CALL_FAILURE   = -3 // Call to .NET function failed

/* Latest error -- Public reset via of_ResetError */
PRIVATEWRITE Long il_ErrorType   
PRIVATEWRITE Long il_ErrorNumber 
PRIVATEWRITE String is_ErrorText 

PRIVATE:
/*  .NET object creation */
Boolean ib_objectCreated

/* Error handler -- Public access via of_SetErrorHandler/of_ResetErrorHandler/of_GetErrorHandler
    Triggers "ue_Error" event for each error when no current error handler */
PowerObject ipo_errorHandler // Each error triggers <ErrorHandler, ErrorEvent>
String is_errorEvent
end variables

forward prototypes
public subroutine of_seterrorhandler (powerobject apo_newhandler, string as_newevent)
public subroutine of_signalerror ()
private subroutine of_setdotneterror (string as_failedfunction, string as_errortext)
public subroutine of_reseterror ()
public function boolean of_createondemand ()
private subroutine of_setassemblyerror (long al_errortype, string as_actiontext, long al_errornumber, string as_errortext)
public subroutine of_geterrorhandler (ref powerobject apo_currenthandler,ref string as_currentevent)
public subroutine of_reseterrorhandler ()
public function long of_send()
public subroutine  of_setmessage(string as_pbmessage)
public subroutine  of_setmessage(string as_pbmessage,boolean abln_pbhtml)
public subroutine  of_setrecipientemail(string as_pbrecipientname,string as_pbrecipientmail)
public subroutine  of_setccrecipientemail(string as_pbccrecipientname,string as_pbccrecipientmail)
public subroutine  of_setbccrecipientemail(string as_pbbccrecipientname,string as_pbbccrecipientmaill)
public subroutine  of_setreplytoemail(string as_pbreplytoname,string as_pbreplytomail)
public subroutine  of_setsenderemail(string as_pbsendername,string as_pbsendermaill)
public subroutine  of_setsmtpserver(string as_pbsmtpserver)
public subroutine  of_setsubject(string as_pbsubject)
public subroutine  of_setattachment(string as_pbattachment)
public subroutine  of_setcharset(string as_pbcharset)
public subroutine  of_setusernamepassword(string as_pbusername,string as_pbpassword)
public subroutine  of_setport(long al_pbport)
public subroutine  of_setauthmethod(long al_pbauthmethod)
public subroutine  of_setconnectiontype(long al_pbconnecttype)
public function string of_getlasterrormessage()
public subroutine  of_setmailername(string as_pbmailername)
public subroutine  of_setpriority(long al_pbpriority)
public subroutine  of_setprioritynone()
public subroutine  of_setprioritylow()
public subroutine  of_setprioritynormal()
public subroutine  of_setpriorityhigh()
public subroutine  of_setreadreceiptrequested(boolean abln_pbreadreceipt)
public function long of_smtpconnect()
public function long of_smtpsend()
public function long of_smtpdisconnect()
end prototypes

event ue_error ( );
/*-----------------------------------------------------------------------------------------*/
/*  Handler undefined or call failed (event undefined) => Signal object itself */
/*-----------------------------------------------------------------------------------------*/
end event

public subroutine of_seterrorhandler (powerobject apo_newhandler, string as_newevent);
//*-----------------------------------------------------------------*/
//*    of_seterrorhandler:  
//*                       Register new error handler (incl. event)
//*-----------------------------------------------------------------*/

This.ipo_errorHandler = apo_newHandler
This.is_errorEvent = Trim(as_newEvent)
end subroutine

public subroutine of_signalerror ();
//*-----------------------------------------------------------------------------*/
//* PRIVATE of_SignalError
//* Triggers error event on previously defined error handler.
//* Calls object's own UE_ERROR when handler or its event is undefined.
//*
//* Handler is "DEFINED" when
//* 	1) <ErrorEvent> is non-empty
//*	2) <ErrorHandler> refers to valid object
//*	3) <ErrorEvent> is actual event on <ErrorHandler>
//*-----------------------------------------------------------------------------*/

Boolean lb_handlerDefined
If This.is_errorEvent > '' Then
	If Not IsNull(This.ipo_errorHandler) Then
		lb_handlerDefined = IsValid(This.ipo_errorHandler)
	End If
End If

If lb_handlerDefined Then
	/* Try to call defined handler*/
	Long ll_status
	ll_status = This.ipo_errorHandler.TriggerEvent(This.is_errorEvent)
	If ll_status = 1 Then Return
End If

/* Handler undefined or call failed (event undefined) => Signal object itself*/
This.event ue_Error( )
end subroutine

private subroutine of_setdotneterror (string as_failedfunction, string as_errortext);
//*----------------------------------------------------------------------------------------*/
//* PRIVATE of_setDotNETError
//* Sets error description for specified error condition exposed by call to .NET  
//*
//* Error description layout
//*			| Call <failedFunction> failed.<EOL>
//*			| Error Text: <errorText> (*)
//* (*): Line skipped when <ErrorText> is empty
//*----------------------------------------------------------------------------------------*/

/* Format description*/
String ls_error
ls_error = "Call " + as_failedFunction + " failed."
If Len(Trim(as_errorText)) > 0 Then
	ls_error += "~r~nError Text: " + as_errorText
End If

/* Retain state in instance variables*/
This.il_ErrorType = This.CALL_FAILURE
This.is_ErrorText = ls_error
This.il_ErrorNumber = 0
end subroutine

public subroutine of_reseterror ();
//*--------------------------------------------*/
//* PUBLIC of_ResetError
//* Clears previously registered error
//*--------------------------------------------*/

This.il_ErrorType = This.SUCCESS
This.is_ErrorText = ''
This.il_ErrorNumber = 0
end subroutine

public function boolean of_createondemand ();
//*--------------------------------------------------------------*/
//*  PUBLIC   of_createOnDemand( )
//*  Return   True:  .NET object created
//*               False: Failed to create .NET object
//*  Loads .NET assembly and creates instance of .NET class.
//*  Uses .NET Core when loading .NET assembly.
//*  Signals error If an error occurs.
//*  Resets any prior error when load + create succeeds.
//*--------------------------------------------------------------*/

This.of_ResetError( )
If This.ib_objectCreated Then Return True // Already created => DONE

Long ll_status 
String ls_action

/* Load assembly using .NET Core */
ls_action = 'Load ' + This.is_AssemblyPath
DotNetAssembly lnv_assembly
lnv_assembly = Create DotNetAssembly
ll_status = lnv_assembly.LoadWithDotNetCore(This.is_AssemblyPath)

/* Abort when load fails */
If ll_status <> 1 Then
	This.of_SetAssemblyError(This.LOAD_FAILURE, ls_action, ll_status, lnv_assembly.ErrorText)
	This.of_SignalError( )
	Return False // Load failed => ABORT
End If

/*   Create .NET object */
ls_action = 'Create ' + This.is_ClassName
ll_status = lnv_assembly.CreateInstance(is_ClassName, This)

/* Abort when create fails */
If ll_status <> 1 Then
	This.of_SetAssemblyError(This.CREATE_FAILURE, ls_action, ll_status, lnv_assembly.ErrorText)
	This.of_SignalError( )
	Return False // Load failed => ABORT
End If

This.ib_objectCreated = True
Return True
end function

private subroutine of_setassemblyerror (long al_errortype, string as_actiontext, long al_errornumber, string as_errortext);
//*----------------------------------------------------------------------------------------------*/
//* PRIVATE of_setAssemblyError
//* Sets error description for specified error condition report by an assembly function
//*
//* Error description layout
//* 		| <actionText> failed.<EOL>
//* 		| Error Number: <errorNumber><EOL>
//* 		| Error Text: <errorText> (*)
//*  (*): Line skipped when <ErrorText> is empty
//*----------------------------------------------------------------------------------------------*/

/*    Format description */
String ls_error
ls_error = as_actionText + " failed.~r~n"
ls_error += "Error Number: " + String(al_errorNumber) + "."
If Len(Trim(as_errorText)) > 0 Then
	ls_error += "~r~nError Text: " + as_errorText
End If

/*  Retain state in instance variables */
This.il_ErrorType = al_errorType
This.is_ErrorText = ls_error
This.il_ErrorNumber = al_errorNumber
end subroutine

public subroutine of_geterrorhandler (ref powerobject apo_currenthandler,ref string as_currentevent);
//*-------------------------------------------------------------------------*/
//* PUBLIC of_GetErrorHandler
//* Return as REF-parameters current error handler (incl. event)
//*-------------------------------------------------------------------------*/

apo_currentHandler = This.ipo_errorHandler
as_currentEvent = This.is_errorEvent
end subroutine

public subroutine of_reseterrorhandler ();
//*---------------------------------------------------*/
//* PUBLIC of_ResetErrorHandler
//* Removes current error handler (incl. event)
//*---------------------------------------------------*/

SetNull(This.ipo_errorHandler)
SetNull(This.is_errorEvent)
end subroutine

public function long of_send();
//*-----------------------------------------------------------------*/
//*  .NET function : Send
//*   Return : Long
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function
Long ll_result

/* Set the dotnet function name */
ls_function = "Send"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		SetNull(ll_result)
		Return ll_result
	End If

	/* Trigger the dotnet function */
	ll_result = This.send()
	Return ll_result
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

	/*  Indicate error occurred */
	SetNull(ll_result)
	Return ll_result
End Try
end function

public subroutine  of_setmessage(string as_pbmessage);
//*-----------------------------------------------------------------*/
//*  .NET function : SetMessage
//*   Argument:
//*              String as_pbmessage
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetMessage"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setmessage(as_pbmessage)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setmessage(string as_pbmessage,boolean abln_pbhtml);
//*-----------------------------------------------------------------*/
//*  .NET function : SetMessage
//*   Argument:
//*              String as_pbmessage
//*              Boolean abln_pbhtml
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetMessage"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setmessage(as_pbmessage,abln_pbhtml)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setrecipientemail(string as_pbrecipientname,string as_pbrecipientmail);
//*-----------------------------------------------------------------*/
//*  .NET function : SetRecipientEmail
//*   Argument:
//*              String as_pbrecipientname
//*              String as_pbrecipientmail
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetRecipientEmail"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setrecipientemail(as_pbrecipientname,as_pbrecipientmail)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setccrecipientemail(string as_pbccrecipientname,string as_pbccrecipientmail);
//*-----------------------------------------------------------------*/
//*  .NET function : SetCCRecipientEmail
//*   Argument:
//*              String as_pbccrecipientname
//*              String as_pbccrecipientmail
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetCCRecipientEmail"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setccrecipientemail(as_pbccrecipientname,as_pbccrecipientmail)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setbccrecipientemail(string as_pbbccrecipientname,string as_pbbccrecipientmaill);
//*-----------------------------------------------------------------*/
//*  .NET function : SetBCCRecipientEmail
//*   Argument:
//*              String as_pbbccrecipientname
//*              String as_pbbccrecipientmaill
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetBCCRecipientEmail"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setbccrecipientemail(as_pbbccrecipientname,as_pbbccrecipientmaill)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setreplytoemail(string as_pbreplytoname,string as_pbreplytomail);
//*-----------------------------------------------------------------*/
//*  .NET function : SetReplyToEmail
//*   Argument:
//*              String as_pbreplytoname
//*              String as_pbreplytomail
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetReplyToEmail"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setreplytoemail(as_pbreplytoname,as_pbreplytomail)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setsenderemail(string as_pbsendername,string as_pbsendermaill);
//*-----------------------------------------------------------------*/
//*  .NET function : SetSenderEmail
//*   Argument:
//*              String as_pbsendername
//*              String as_pbsendermaill
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetSenderEmail"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setsenderemail(as_pbsendername,as_pbsendermaill)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setsmtpserver(string as_pbsmtpserver);
//*-----------------------------------------------------------------*/
//*  .NET function : SetSMTPServer
//*   Argument:
//*              String as_pbsmtpserver
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetSMTPServer"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setsmtpserver(as_pbsmtpserver)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setsubject(string as_pbsubject);
//*-----------------------------------------------------------------*/
//*  .NET function : SetSubject
//*   Argument:
//*              String as_pbsubject
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetSubject"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setsubject(as_pbsubject)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setattachment(string as_pbattachment);
//*-----------------------------------------------------------------*/
//*  .NET function : SetAttachment
//*   Argument:
//*              String as_pbattachment
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetAttachment"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setattachment(as_pbattachment)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setcharset(string as_pbcharset);
//*-----------------------------------------------------------------*/
//*  .NET function : SetCharSet
//*   Argument:
//*              String as_pbcharset
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetCharSet"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setcharset(as_pbcharset)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setusernamepassword(string as_pbusername,string as_pbpassword);
//*-----------------------------------------------------------------*/
//*  .NET function : SetUsernamePassword
//*   Argument:
//*              String as_pbusername
//*              String as_pbpassword
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetUsernamePassword"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setusernamepassword(as_pbusername,as_pbpassword)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setport(long al_pbport);
//*-----------------------------------------------------------------*/
//*  .NET function : SetPort
//*   Argument:
//*              Long al_pbport
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetPort"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setport(al_pbport)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setauthmethod(long al_pbauthmethod);
//*-----------------------------------------------------------------*/
//*  .NET function : SetAuthMethod
//*   Argument:
//*              Long al_pbauthmethod
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetAuthMethod"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setauthmethod(al_pbauthmethod)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setconnectiontype(long al_pbconnecttype);
//*-----------------------------------------------------------------*/
//*  .NET function : SetConnectionType
//*   Argument:
//*              Long al_pbconnecttype
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetConnectionType"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setconnectiontype(al_pbconnecttype)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public function string of_getlasterrormessage();
//*-----------------------------------------------------------------*/
//*  .NET function : GetLastErrorMessage
//*   Return : String
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function
String ls_result

/* Set the dotnet function name */
ls_function = "GetLastErrorMessage"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		SetNull(ls_result)
		Return ls_result
	End If

	/* Trigger the dotnet function */
	ls_result = This.getlasterrormessage()
	Return ls_result
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

	/*  Indicate error occurred */
	SetNull(ls_result)
	Return ls_result
End Try
end function

public subroutine  of_setmailername(string as_pbmailername);
//*-----------------------------------------------------------------*/
//*  .NET function : SetMailerName
//*   Argument:
//*              String as_pbmailername
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetMailerName"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setmailername(as_pbmailername)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setpriority(long al_pbpriority);
//*-----------------------------------------------------------------*/
//*  .NET function : SetPriority
//*   Argument:
//*              Long al_pbpriority
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetPriority"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setpriority(al_pbpriority)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setprioritynone();
//*-----------------------------------------------------------------*/
//*  .NET function : SetPriorityNone
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetPriorityNone"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setprioritynone()
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setprioritylow();
//*-----------------------------------------------------------------*/
//*  .NET function : SetPriorityLow
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetPriorityLow"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setprioritylow()
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setprioritynormal();
//*-----------------------------------------------------------------*/
//*  .NET function : SetPriorityNormal
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetPriorityNormal"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setprioritynormal()
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setpriorityhigh();
//*-----------------------------------------------------------------*/
//*  .NET function : SetPriorityHigh
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetPriorityHigh"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setpriorityhigh()
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public subroutine  of_setreadreceiptrequested(boolean abln_pbreadreceipt);
//*-----------------------------------------------------------------*/
//*  .NET function : SetReadReceiptRequested
//*   Argument:
//*              Boolean abln_pbreadreceipt
//*   Return : (None)
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function

/* Set the dotnet function name */
ls_function = "SetReadReceiptRequested"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		Return 
	End If

	/* Trigger the dotnet function */
	This.setreadreceiptrequested(abln_pbreadreceipt)
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

End Try
end subroutine

public function long of_smtpconnect();
//*-----------------------------------------------------------------*/
//*  .NET function : SmtpConnect
//*   Return : Long
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function
Long ll_result

/* Set the dotnet function name */
ls_function = "SmtpConnect"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		SetNull(ll_result)
		Return ll_result
	End If

	/* Trigger the dotnet function */
	ll_result = This.smtpconnect()
	Return ll_result
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

	/*  Indicate error occurred */
	SetNull(ll_result)
	Return ll_result
End Try
end function

public function long of_smtpsend();
//*-----------------------------------------------------------------*/
//*  .NET function : SmtpSend
//*   Return : Long
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function
Long ll_result

/* Set the dotnet function name */
ls_function = "SmtpSend"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		SetNull(ll_result)
		Return ll_result
	End If

	/* Trigger the dotnet function */
	ll_result = This.smtpsend()
	Return ll_result
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

	/*  Indicate error occurred */
	SetNull(ll_result)
	Return ll_result
End Try
end function

public function long of_smtpdisconnect();
//*-----------------------------------------------------------------*/
//*  .NET function : SmtpDisconnect
//*   Return : Long
//*-----------------------------------------------------------------*/
/* .NET  function name */
String ls_function
Long ll_result

/* Set the dotnet function name */
ls_function = "SmtpDisconnect"

Try
	/* Create .NET object */
	If Not This.of_createOnDemand( ) Then
		SetNull(ll_result)
		Return ll_result
	End If

	/* Trigger the dotnet function */
	ll_result = This.smtpdisconnect()
	Return ll_result
Catch(runtimeerror re_error)

	If This.ib_CrashOnException Then Throw re_error

	/*   Handle .NET error */
	This.of_SetDotNETError(ls_function, re_error.text)
	This.of_SignalError( )

	/*  Indicate error occurred */
	SetNull(ll_result)
	Return ll_result
End Try
end function

on nvo_mailkitsmptrsr.create
call super::create
TriggerEvent( this, "constructor" )
end on

on nvo_mailkitsmptrsr.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

