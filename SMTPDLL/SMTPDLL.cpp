// SMTPDLL.cpp : Defines the entry point for the DLL application.
//

#include "stdafx.h"
#include "SMTP.h"

BOOL APIENTRY DllMain( HANDLE hModule, 
                       DWORD  ul_reason_for_call, 
                       LPVOID lpReserved
					 )
{
	if (DLL_PROCESS_ATTACH==ul_reason_for_call){
		InitWinsock(); 
	}else if (DLL_PROCESS_DETACH==ul_reason_for_call){
		CleanWinsock();
	}
		
	return TRUE;
}

bool __stdcall SendEmail(const char *SMTPServer,const char *EMailName,const char *Password,const char *Dest,const char *Sorc,const char *Subject,const char *BodyCaption)
{
	return SMTPSend(SMTPServer,EMailName,Password,Dest,Sorc,Subject,BodyCaption);
}