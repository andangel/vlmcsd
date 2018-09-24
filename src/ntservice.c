#ifndef CONFIG
#define CONFIG "config.h"
#endif // CONFIG
#include CONFIG

#include "ntservice.h"
#include "shared_globals.h"
#include "vlmcsd.h"
#include "output.h"
#include "helpers.h"

#ifdef _NTSERVICE

SERVICE_STATUS          gSvcStatus;
SERVICE_STATUS_HANDLE   gSvcStatusHandle;

VOID WINAPI ServiceCtrlHandler(DWORD dwCtrl)
{
	// Handle the requested control code.
	switch (dwCtrl)
	{
	case SERVICE_CONTROL_STOP:
	case SERVICE_CONTROL_SHUTDOWN:

		ServiceShutdown = TRUE;
		ReportServiceStatus(SERVICE_STOP_PENDING, NO_ERROR, 0);

		// Remove PID file and free ressources
		cleanup();
#			if __CYGWIN__ || defined(USE_MSRPC)
		ReportServiceStatus(SERVICE_STOPPED, NO_ERROR, 0);
#			endif // __CYGWIN__

	default:
		break;
	}
}

static VOID WINAPI ServiceMain(const int argc_unused, CARGV argv_unused)
{
	// Register the handler function for the service
	//注册该服务的处理函数
	if (!((gSvcStatusHandle = RegisterServiceCtrlHandler(NT_SERVICE_NAME, ServiceCtrlHandler))))
	{
		return;
	}

	// These SERVICE_STATUS members remain as set here
	//这些SERVICE_STATUS成员保持不变
	gSvcStatus.dwServiceType = SERVICE_WIN32_OWN_PROCESS;
	gSvcStatus.dwServiceSpecificExitCode = 0;

	// Run the actual program
	//运行实际的程序
	ReportServiceStatus(SERVICE_STOPPED, newmain(), 3000);
}

SERVICE_TABLE_ENTRY NTServiceDispatchTable[] = {
	{
		(LPSTR)NT_SERVICE_NAME,
		(LPSERVICE_MAIN_FUNCTION)ServiceMain
	},
	{
		NULL,
		NULL
	}
};

VOID ReportServiceStatus(const DWORD dwCurrentState, const DWORD dwWin32ExitCode, const DWORD dwWaitHint)
{
	static DWORD dwCheckPoint = 1;

	// Fill in the SERVICE_STATUS structure.
	//填写SERVICE_STATUS结构
	gSvcStatus.dwCurrentState = dwCurrentState;
	gSvcStatus.dwWin32ExitCode = dwWin32ExitCode;
	gSvcStatus.dwWaitHint = dwWaitHint;

	if (dwCurrentState == SERVICE_START_PENDING)
		gSvcStatus.dwControlsAccepted = 0;
	else
		gSvcStatus.dwControlsAccepted = SERVICE_ACCEPT_STOP;

	if ((dwCurrentState == SERVICE_RUNNING) ||
		(dwCurrentState == SERVICE_STOPPED))
		gSvcStatus.dwCheckPoint = 0;
	else
		gSvcStatus.dwCheckPoint = dwCheckPoint++;

	// Report the status of the service to the SCM.
	SetServiceStatus(gSvcStatusHandle, &gSvcStatus);
}

/*VOID ServiceReportEvent(char *szFunction)
{
	HANDLE hEventSource;
	const char *eventStrings[2];
	TCHAR Buffer[80];

	hEventSource = RegisterEventSource(NULL, NT_SERVICE_NAME);

	if (hEventSource)
	{
		snprintf(Buffer, 80, "%s failed with %d", szFunction, GetLastError());

		eventStrings[0] = NT_SERVICE_NAME;
		eventStrings[1] = Buffer;

		ReportEvent(hEventSource,        // event log handle
					EVENTLOG_ERROR_TYPE, // event type
					0,                   // event category
					00,           // event identifier
					NULL,                // no security identifier
					2,                   // size of lpszStrings array
					0,                   // no binary data
					eventStrings,         // array of strings
					NULL);               // no binary data

		DeregisterEventSource(hEventSource);
	}
}*/

//Returns 0=Error, 1=Success, 2=Doesn't exist


static uint_fast8_t OpenAndRemoveService(DWORD *dwPreviousState, SC_HANDLE *schSCManager)
{
	SERVICE_STATUS status;
	uint_fast8_t i;
	SC_HANDLE installedService;
	uint_fast8_t result = 1;
	BOOL closeManager = FALSE;

	// Allow NULL for both Arguments
	//两个参数都允许NULL
	if (!dwPreviousState) dwPreviousState = (DWORD*)alloca(sizeof(*dwPreviousState));
	if (!schSCManager)
	{
		schSCManager = (SC_HANDLE*)alloca(sizeof(*schSCManager));
		closeManager = TRUE;
	}

	*schSCManager = OpenSCManager(
		NULL,                    // local computer 本地计算机
		NULL,                    // ServicesActive database ServicesActive数据库
		SC_MANAGER_ALL_ACCESS);  // full access rights 完全访问权限

	if (!*schSCManager) return 0;

	if (!((installedService = OpenService(*schSCManager, NT_SERVICE_NAME, SERVICE_ALL_ACCESS))))
	{
		result = 2;
	}
	else
	{
		*dwPreviousState = SERVICE_STOPPED;
		if (QueryServiceStatus(installedService, &status)) *dwPreviousState = status.dwCurrentState;

		ControlService(installedService, SERVICE_CONTROL_STOP, &status);

		for (i = 0; i < 10; i++)
		{
			QueryServiceStatus(installedService, &status);
			// Give it 100 ms after it reported SERVICE_STOPPED. Subsequent CreateService will fail otherwise
			//在报告SERVICE_STOPPED 100毫秒后给它。 否则，后续的CreateService将失败
			Sleep(100);
			if (status.dwCurrentState == SERVICE_STOPPED) break;
		}

		if (!DeleteService(installedService)) result = 0;
		CloseServiceHandle(installedService);
	}

	if (closeManager) CloseServiceHandle(*schSCManager);
	return result;
}
//
//服务安装
//
static VOID ServiceInstaller(const char *restrict ServiceUser, const char *const ServicePassword)
{
	SC_HANDLE schSCManager; //服务控制管理程序维护的登记数据库的句柄，由系统函数OpenSCManager 返回
	SC_HANDLE schService;
	char szPath[MAX_PATH] = "\"";
	//GetModuleFileName获取模块文件名
	if (!GetModuleFileName(NULL, szPath + sizeof(char), MAX_PATH - 1))
	{
		errorout("无法安装服务 (%d)\n", (uint32_t)GetLastError());
		return;
	}

	strcat(szPath, "\"");

	int i;
	for (i = 1; i < global_argc; i++)
	{
		// Strip unneccessary parameters, especially the password
		//剥去不必要的参数，特别是密码
		if (!strcmp(global_argv[i], "-s")) continue;

		if (!strcmp(global_argv[i], "-W") ||
			!strcmp(global_argv[i], "-U"))
		{
			i++;
			continue;
		}

		strcat(szPath, " ");

		if (strchr(global_argv[i], ' '))
		{
			strcat(szPath, "\"");
			strcat(szPath, global_argv[i]);
			strcat(szPath, "\"");
		}
		else
			strcat(szPath, global_argv[i]);
	}

	// Get a handle to the SCM database.
	//获取SCM数据库的句柄。

	SERVICE_STATUS status;
	DWORD dwPreviousState;

	if (!OpenAndRemoveService(&dwPreviousState, &schSCManager))
	{
		errorout("服务删除失败 (%d)\n", (uint32_t)GetLastError());
		return;
	}

	char *tempUser = NULL;

	if (ServiceUser)
	{
		// Shortcuts for some well known users
		if (!strcasecmp(ServiceUser, "/l")) ServiceUser = "NT AUTHORITY\\LocalService";
		if (!strcasecmp(ServiceUser, "/n")) ServiceUser = "NT AUTHORITY\\NetworkService";

		// Allow Local Users without .\ , e.g. "johndoe" instead of ".\johndoe"
		//允许没有。\的本地用户，例如 “johndoe”而不是“。\ johndoe”
		if (!strchr(ServiceUser, '\\'))
		{
			tempUser = (char*)vlmcsd_malloc(strlen(ServiceUser) + 3);
			strcpy(tempUser, ".\\");
			strcat(tempUser, ServiceUser);
			ServiceUser = tempUser;
		}
	}

	schService = CreateService(
		schSCManager,				//  //服务控制管理程序维护的登记数据库的句柄，由系统函数OpenSCManager 返回
		NT_SERVICE_NAME,			// name of service 服务名称
		NT_SERVICE_DISPLAY_NAME,	// service name to display 显示服务名称
		SERVICE_ALL_ACCESS,			// desired access
		SERVICE_WIN32_OWN_PROCESS,	// service type 服务类型
		SERVICE_AUTO_START,			// start type 启动类型
		SERVICE_ERROR_NORMAL,		// error control type 错误控制类型
		szPath,						// path to service's binary 服务二进制路径
		NULL,						// no load ordering group 空载顺序组
		NULL,						// no tag identifier 没有标签标识符
		"tcpip\0",			        // depends on TCP/IP 取决于TCP/IP
		ServiceUser,				// LocalSystem account LocalSystem帐户
		ServicePassword);			// no password 没有密码

#	if __clang__ && (__CYGWIN__ || __MINGW64__ )
	// Workaround for clang not understanding some GCC asm syntax used in <w32api/psdk_inc/intrin-impl.h>
	ZeroMemory((char*)ServicePassword, strlen(ServicePassword));
#	else
	SecureZeroMemory((char*)ServicePassword, strlen(ServicePassword));
#	endif
	if (tempUser) free(tempUser);

	if (schService == NULL)
	{
		errorout("创建服务失败 (%u)\n", (uint32_t)GetLastError());
		CloseServiceHandle(schSCManager);
		return;
	}
	else
	{
		errorout("服务安装成功\n");

		if (dwPreviousState == SERVICE_RUNNING)
		{
			printf("重新启动 " NT_SERVICE_NAME " 服务 => ");
			status.dwCurrentState = SERVICE_STOPPED;

			if (StartService(schService, 0, NULL))
			{
				for (i = 0; i < 10; i++)
				{
					if (!QueryServiceStatus(schService, &status) || status.dwCurrentState != SERVICE_START_PENDING) break;
					Sleep(100);
				}

				if (status.dwCurrentState == SERVICE_RUNNING)
					printf("成功\n");
				else if (status.dwCurrentState == SERVICE_START_PENDING)
					printf("Not ready within a second\n");
				else
					errorout("Error\n");
			}
			else
				errorout("Error %u\n", (uint32_t)GetLastError());
		}
	}

	CloseServiceHandle(schService);
	CloseServiceHandle(schSCManager);
}

int NtServiceInstallation(const int_fast8_t installService, const char *restrict ServiceUser, const char *const ServicePassword)
{
	if (IsNTService) return 0;

	if (installService == 1) // Install
	{
		ServiceInstaller(ServiceUser, ServicePassword);
		return(0);
	}

	if (installService == 2) // Remove
	{
		switch (OpenAndRemoveService(NULL, NULL))
		{
		case 0:
			errorout("错误 删除服务. %s\n", NT_SERVICE_NAME);
			return(!0);
		case 1:
			printf("服务 %s 成功移除.\n", NT_SERVICE_NAME);
			return(0);
		default:
			errorout("服务 %s不存在.\n", NT_SERVICE_NAME);
			return(!0);
		}
	}

	// Do nothing

	return(0);
}
#endif // _NTSERVICE
