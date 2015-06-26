
#include <stdio.h>
#include <string.h>
#include <windows.h>
#include <winsock.h>
#include <memory.h>
#include <malloc.h>

#pragma comment(lib,"ws2_32")

void InitWinsock(void);
void CleanWinsock(void);
char*ResolveHost(const char *HostName);
bool SMTPSend(const char *SMTPServer,const char *EMailName,const char *Password,const char *Dest,const char *Sorc,const char *Subject,const char *BodyCaption);

#define VER_WINSOCK                1

#define LEN_EMAIL_NAME             64
#define LEN_EMAIL_PASSOWRD         LEN_EMAIL_NAME
#define LEN_EMAIL_SERVER           LEN_EMAIL_PASSOWRD
#define LEN_EMAIL_SUBJECT          LEN_EMAIL_SERVER
#define LEN_EMAIL_CAPTION          65536

#define PORT_SMTP                  25
#define PORT_LOCAL                 10132

#define CMD_MAIL_TEST              "noop "
#define CMD_MAIL_HELO              "helo "
#define CMD_MAIL_LOGI              "auth login "
#define CMD_MAIL_DATA              "data "
#define CMD_MAIL_DATA_EXIT         "."
#define CMD_MAIL_SEND              "mail from:<"
#define CMD_MAIL_RECV              "rcpt to:<"
#define CMD_MAIL_QUIT              "quit "

#define BUF_SEND                   64
#define BUF_RECV                   BUF_SEND * 4

char   *base64_encode(const char* data, int data_len);
char   *base64_decode(const char* data, int data_len);

#define STR_LEFT                   ">"
#define STR_CRLF                   "\r\n\0"
#define STR_MAIL_FROM              "from:"
#define STR_MAIL_TO                "to:"
#define STR_MAIL_SUBJECT           "subject:"

/*

		LCatro give you that code about how to use it..



VC++ 6:



// EMailAttack.cpp : Defines the entry point for the console application.
//

#include "stdafx.h"
#include "SMTP.h"

char   *EMail         = (char *)malloc(LEN_EMAIL_NAME);
char   *Password      = (char *)malloc(LEN_EMAIL_PASSOWRD);
char   *SMTPServer    = (char *)malloc(LEN_EMAIL_SERVER);
char   *EMailName     = (char *)malloc(LEN_EMAIL_NAME);
char   *TargeEMail    = (char *)malloc(LEN_EMAIL_NAME);
char   *Subject       = (char *)malloc(LEN_EMAIL_SUBJECT);
char   *EMailDataBody = (char *)malloc(LEN_EMAIL_CAPTION);

void Clean(void){
	free(EMail);
	EMail=NULL;
	free(Password);
	Password=NULL;
	free(SMTPServer);
	SMTPServer=NULL;
	free(EMailName);
	EMailName=NULL;
	free(TargeEMail);
	TargeEMail=NULL;
	free(Subject);
	Subject=NULL;
	free(EMailDataBody);
	EMailDataBody=NULL;
}

int main(int argc, char* argv[])
{
	if (NULL==EMail || NULL==Password || NULL==SMTPServer || NULL==EMailName || NULL==TargeEMail){
		printf("malloc memory for save email name and password and SMTP-server had err!\n");
		return -1;
	}

	InitWinsock();
	
	printf("Input your E-Mail Name:");
	scanf("%s",EMailName);
	printf("Input your E-Mail Password:");
	scanf("%s",Password);
	printf("Input the E-Mail SMTP Server Host:");
	scanf("%s",SMTPServer);
	printf("Input your E-Mail:");
	scanf("%s",EMail);
	printf("Input the targe E-Mail:");
	scanf("%s",TargeEMail);
	printf("EMail subject:");
	scanf("%s",Subject);
	printf("EMail Data:");
	scanf("%s",EMailDataBody);

	SMTPSend((const char *)SMTPServer,(const char *)EMailName,(const char *)Password,(const char *)TargeEMail,(const char *)EMail,(const char *)Subject,(const char *)EMailDataBody);

	Clean();
	CleanWinsock();

	getchar();
	return 0;
}



*/