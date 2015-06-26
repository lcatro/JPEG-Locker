
#include "stdafx.h"
#include "SMTP.h"

const char base[] = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=";   

char *base64_encode(const char* data, int data_len)   
{   
    //int data_len = strlen(data);   
    int prepare = 0;   
    int ret_len;   
    int temp = 0;   
    char *ret = NULL;   
    char *f = NULL;   
    int tmp = 0;   
    char changed[4];   
    int i = 0;   
    ret_len = data_len / 3;   
    temp = data_len % 3;   
    if (temp > 0)   
    {   
        ret_len += 1;   
    }   
    ret_len = ret_len*4 + 1;   
    ret = (char *)malloc(ret_len);   
       
    if ( ret == NULL)   
		return NULL;

    memset(ret, 0, ret_len);   
    f = ret;   
    while (tmp < data_len)   
    {   
        temp = 0;   
        prepare = 0;   
        memset(changed, '\0', 4);   
        while (temp < 3)   
        {   
            //printf("tmp = %d\n", tmp);   
            if (tmp >= data_len)   
            {   
                break;   
            }   
            prepare = ((prepare << 8) | (data[tmp] & 0xFF));   
            tmp++;   
            temp++;   
        }   
        prepare = (prepare<<((3-temp)*8));   
        //printf("before for : temp = %d, prepare = %d\n", temp, prepare);   
        for (i = 0; i < 4 ;i++ )   
        {   
            if (temp < i)   
            {   
                changed[i] = 0x40;   
            }   
            else 
            {   
                changed[i] = (prepare>>((3-i)*6)) & 0x3F;   
            }   
            *f = base[changed[i]];   
            //printf("%.2X", changed[i]);   
            f++;   
        }   
    }   
    *f = '\0';   
       
    return ret;   
       
}   

void InitWinsock(void)
{
	WSADATA WSAData;
	memset(&WSAData,0,sizeof(WSADATA));
	WSAStartup(VER_WINSOCK,&WSAData);
}

void CleanWinsock(void)
{
	WSACleanup();
}

char *ResolveHost(const char *HostName)
{
	hostent *HostData=NULL;
	char *IP=NULL;
	HostData=gethostbyname(HostName);
	if (NULL!=HostData){
		IP=inet_ntoa(*((struct in_addr *)HostData->h_addr_list[0]));
		return IP;
	}else
		return NULL;
}

bool SMTPSend(const char *SMTPServer,const char *EMailName,const char *Password,const char *Dest,const char *Sorc,const char *Subject,const char *BodyCaption)
{
	bool rtn=false;

	SOCKET sock=socket(AF_INET,SOCK_STREAM,IPPROTO_TCP);
	if (sock!=0){

		sockaddr_in Local;
		memset(&Local,0,sizeof(sockaddr_in));
		Local.sin_family=AF_INET;
		Local.sin_port=htons(PORT_LOCAL);

		if (SOCKET_ERROR==bind(sock,(const sockaddr *)&Local,sizeof(sockaddr)))
			goto Clean;

		char *SMTPIP=ResolveHost((const char *)SMTPServer);

		sockaddr_in SMTP;
		memset(&SMTP,0,sizeof(sockaddr_in));
		SMTP.sin_addr.S_un.S_addr=inet_addr(SMTPIP);
		SMTP.sin_family=AF_INET;
		SMTP.sin_port=htons(PORT_SMTP);

		if (SOCKET_ERROR!=connect(sock,(const sockaddr *)&SMTP,sizeof(sockaddr))){
			printf("Connect to SMTP_Server OK!\n");

			char SendBuf[BUF_SEND]={0};
			char RecvBuf[BUF_RECV]={0};

			strcat((char *)SendBuf,(const char *)CMD_MAIL_HELO);
			strcat((char *)SendBuf,(const char *)SMTPServer);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			recv(sock,RecvBuf,BUF_RECV,0);
			memset(SendBuf,0,BUF_SEND);
			memset(RecvBuf,0,BUF_RECV);

			strcat((char *)SendBuf,(const char *)CMD_MAIL_LOGI);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			recv(sock,RecvBuf,BUF_RECV,0);
			recv(sock,RecvBuf,BUF_RECV,0);
			memset(SendBuf,0,BUF_SEND);
			memset(RecvBuf,0,BUF_RECV);

			char *EMailEncode=base64_encode((const char *)EMailName,strlen(EMailName));
			char *PasswordEncode=base64_encode((const char *)Password,strlen(Password));

			memcpy(SendBuf,(const void *)EMailEncode,strlen(EMailEncode));
			strcat(SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			recv(sock,RecvBuf,BUF_RECV,0);
			free(EMailEncode);
			EMailEncode=NULL;
			memset(SendBuf,0,BUF_SEND);
			memset(RecvBuf,0,BUF_RECV);

			memcpy(SendBuf,(const void *)PasswordEncode,strlen(PasswordEncode));
			strcat(SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			recv(sock,RecvBuf,BUF_RECV,0);
			free(PasswordEncode);
			PasswordEncode=NULL;
			memset(SendBuf,0,BUF_SEND);
			memset(RecvBuf,0,BUF_RECV);

			strcat((char *)SendBuf,(const char *)CMD_MAIL_SEND);
			strcat((char *)SendBuf,(const char *)Sorc);
			strcat((char *)SendBuf,(const char *)STR_LEFT);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			recv(sock,RecvBuf,BUF_RECV,0);
			memset(SendBuf,0,BUF_SEND);
			memset(RecvBuf,0,BUF_RECV);

			strcat((char *)SendBuf,(const char *)CMD_MAIL_RECV);
			strcat((char *)SendBuf,(const char *)Dest);
			strcat((char *)SendBuf,(const char *)STR_LEFT);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			recv(sock,RecvBuf,BUF_RECV,0);
			memset(SendBuf,0,BUF_SEND);
			memset(RecvBuf,0,BUF_RECV);

			strcat((char *)SendBuf,(const char *)CMD_MAIL_DATA);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			recv(sock,RecvBuf,BUF_SEND,0);
			memset(SendBuf,0,BUF_SEND);
			memset(RecvBuf,0,BUF_RECV);

			strcat((char *)SendBuf,(const char *)STR_MAIL_FROM);
			strcat((char *)SendBuf,(const char *)Sorc);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			memset(SendBuf,0,BUF_SEND);

			strcat((char *)SendBuf,(const char *)STR_MAIL_TO);
			strcat((char *)SendBuf,(const char *)Dest);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			memset(SendBuf,0,BUF_SEND);

			strcat((char *)SendBuf,(const char *)STR_MAIL_SUBJECT);
			strcat((char *)SendBuf,(const char *)Subject);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			memset(SendBuf,0,BUF_SEND);

			strcat((char *)SendBuf,(const char *)BodyCaption);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);
			memset(SendBuf,0,BUF_SEND);

			strcat((char *)SendBuf,(const char *)STR_CRLF);
			strcat((char *)SendBuf,(const char *)CMD_MAIL_DATA_EXIT);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);			
			memset(SendBuf,0,BUF_SEND);
			recv(sock,RecvBuf,BUF_SEND,0);
			printf("Recv:%s\n",RecvBuf);

			strcat((char *)SendBuf,(const char *)CMD_MAIL_QUIT);
			strcat((char *)SendBuf,(const char *)STR_CRLF);
			send(sock,(const char *)SendBuf,strlen(SendBuf),0);

			rtn=true;
		}else{
			printf("Connect to SMTP-Server ERR!\n");
		}
	}else{
		printf("Create socket ERR!\n");
	}

Clean:
	closesocket(sock);
	return rtn;
}
