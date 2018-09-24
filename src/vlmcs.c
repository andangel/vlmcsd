#ifndef CONFIG
#define CONFIG "config.h"
#endif // CONFIG
#include CONFIG

#ifndef _GNU_SOURCE
#define _GNU_SOURCE
#endif

#ifndef _CRT_SECURE_NO_WARNINGS
#define _CRT_SECURE_NO_WARNINGS
#endif

#include "vlmcs.h"
#if _MSC_VER
#include <Shlwapi.h>
#endif
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <errno.h>
#include <stdint.h>
#if _MSC_VER
#include "wingetopt.h"
#else
#include <getopt.h>
#endif
#include <sys/types.h>
#include <sys/stat.h>
#ifndef _MSC_VER
#include <unistd.h>
#endif
#ifndef _WIN32
#include <sys/ioctl.h>
#include <termios.h>
#else // _WIN32
#endif // _WIN32
#include "endian.h"
#include "shared_globals.h"
#include "output.h"
#ifndef USE_MSRPC
#include "network.h"
#include "rpc.h"
#else // USE_MSRPC
#include "msrpc-client.h"
#endif // USE_MSRPC
#include "kms.h"
#include "helpers.h"
#include "dns_srv.h"


#define VLMCS_OPTION_GRAB_INI 1
#define VLMCS_OPTION_NO_GRAB_INI 2

//#define kmsVersionMinor 0 // Currently constant. May change in future KMS versions

#ifndef IS_LIBRARY

// Function Prototypes
static void CreateRequestBase(REQUEST *Request);


// KMS Parameters
#ifndef NO_VERBOSE_LOG
static int_fast8_t verbose = FALSE;
#endif

static int_fast8_t VMInfo = FALSE;
static int_fast8_t dnsnames = TRUE;
static int FixedRequests = 0;
static DWORD LicenseStatus = 0x02;
static const char *CMID = NULL;
static const char *CMID_prev = NULL;
static const char *WorkstationName = NULL;
static int BindingExpiration = 43200; //30 days
static const char *RemoteAddr;
static int_fast8_t ReconnectForEachRequest = FALSE;
#ifndef USE_MSRPC
static int AddressFamily = AF_UNSPEC;
#else
static int AddressFamily = 0;
#endif // USE_MSRPC
static int_fast8_t incompatibleOptions = 0;
static const char* fn_ini_client = NULL;
//static int_fast16_t kmsVersionMinor = 0;
static const char* ePidGroup[] = { "Windows", "Office2010", "Office2013", "Office2016" };
static int32_t ActiveProductIndex = 0;
static int32_t NCountPolicy = 0;
static GUID AppGuid, KmsGuid, SkuGuid;
static uint16_t MinorVersion = 0;
static uint16_t MajorVersion;

//#if !MULTI_CALL_BINARY
//uint8_t DefaultKmsData[]={0};
//__pure size_t getDefaultKmsDataSize() { return (size_t)0; }
//#endif // !MULTI_CALL_BINARY

#ifndef NO_DNS
static int_fast8_t NoSrvRecordPriority = FALSE;
#endif // NO_DNS


typedef char iniFileEpidLines[4][256];

typedef struct
{
	const char* first[16];
	const char* second[16];
	const char* tld[22];
} DnsNames;


// Some names for the DNS name random generator
static DnsNames ClientDnsNames =
{
	{ "www", "ftp", "kms", "hack-me", "smtp", "ns1", "mx1", "ns1", "pop3", "imap", "mail", "dns", "headquarter", "we-love", "_vlmcs._tcp", "ceo-laptop" },
	{ ".microsoft", ".apple", ".amazon", ".samsung", ".adobe", ".google", ".yahoo", ".facebook", ".ubuntu", ".oracle", ".borland", ".htc", ".acer", ".windows", ".linux", ".sony" },
	{ ".com", ".net", ".org", ".cn", ".co.uk", ".de", ".com.tw", ".us", ".fr", ".it", ".me", ".info", ".biz", ".co.jp", ".ua", ".at", ".es", ".pro", ".by", ".ru", ".pl", ".kr" }
};

// Request Count Control Variables
static int RequestsToGo = 1;
static BOOL firstRequestSent = FALSE;


static void string2UuidOrExit(const char *const restrict input, GUID *const restrict guid)
{
	if (strlen(input) != GUID_STRING_LENGTH || !string2UuidLE(input, guid))
	{
		errorout("严重的:命令行包含无效的GUID.\n");
		exit(VLMCSD_EINVAL);
	}
}


#ifndef NO_HELP

__noreturn static void clientUsage(const char* const programName)
{
	errorout(
		"vlmcs %s \n\n"
#		ifndef NO_DNS
		"Usage: %s [options] [ <host>[:<port>] | .<domain> | - ] [options]\n\n"
#		else // DNS
		"Usage: %s [options] [<host>[:<port>]] [options]\n\n"
#		endif // DNS

		"选项:\n\n"

#		ifndef NO_VERBOSE_LOG
		"  -v Be verbose\n"
#		endif
		"  -l <app>\n"
		"  -4 Force V4 protocol\n"
		"  -5 Force V5 protocol\n"
		"  -6 Force V6 protocol\n"
#		ifndef USE_MSRPC
		"  -i <IpVersion> 使用 (4 or 6)IP协议\n"
#		endif // USE_MSRPC
#		ifndef NO_EXTERNAL_DATA
		"  -j <file> 读取外部的 KMS 数据文件 <file>\n"
#		endif // NO_EXTERNAL_DATA
		"  -e Show some valid examples\n"
		"  -x Show valid Apps\n"
		"  -d no DNS names, use Netbios names (no effect if -w is used)\n"
		"  -V show version information and exit\n\n"

		"Advanced options:\n\n"

		"  -a <AppGUID> Use custom Application GUID\n"
		"  -s <ActGUID> Use custom Activation Configuration GUID\n"
		"  -k <KmsGUID> Use custom KMS GUID\n"
		"  -c <ClientGUID> Use custom Client GUID. Default: Use random\n"
		"  -o <PreviousClientGUID> Use custom Prevoius Client GUID. Default: ZeroGUID\n"
		"  -K <ProtocolVersion> Use a specific (possibly invalid) protocol version\n"
		"  -w <Workstation> Use custom workstation name. Default: Use random\n"
		"  -r <RequiredClientCount> Fake required clients\n"
		"  -n <Requests> Fixed # of requests (Default: Enough to charge)\n"
		"  -m Pretend to be a virtual machine\n"
		"  -G <file> Get ePID/HwId data and write to <file>. Can't be used with -l, -4, -5, -6, -a, -s, -k, -r and -n\n"
#		ifndef USE_MSRPC
		"  -T Use a new TCP connection for each request.\n"
		"  -N <0|1> disable or enable NDR64. Default: 1\n"
		"  -B <0|1> disable or enable RPC bind time feature negotiation. Default: 1\n"
#		endif // USE_MSRPC
		"  -t <LicenseStatus> Use specfic license status (0 <= T <= 6)\n"
		"  -g <BindingExpiration> Use a specfic binding expiration time in minutes. Default 43200\n"
#		ifndef NO_DNS
		"  -P Ignore priority and weight in DNS SRV records\n"
#		endif // NO_DNS
#		ifndef USE_MSRPC
		"  -p Don't use multiplexed RPC bind\n"
#		endif // USE_MSRPC
		"\n"

		"<port>:\t\tTCP port name of the KMS to use. Default 1688.\n"
		"<host>:\t\thost name of the KMS to use. Default 127.0.0.1\n"
#		ifndef NO_DNS
		".<domain>:\tfind KMS server in <domain> via DNS\n"
#		endif // NO_DNS
		"<app>:\t\t(Type %s -x to see a list of valid apps)\n\n",
		Version, programName, programName
	);

	exit(VLMCSD_EINVAL);
}

__pure static int getLineWidth(void)
{
#ifdef TERMINAL_FIXED_WIDTH // For Toolchains that to not have winsize
	return TERMINAL_FIXED_WIDTH;
#else // Can determine width of terminal
#ifndef _WIN32

	struct winsize w;

	if (ioctl(STDOUT_FILENO, TIOCGWINSZ, &w))
	{
		return 80; // Return this if stdout is not a tty
	}

	return w.ws_col;

#else // _WIN32

	CONSOLE_SCREEN_BUFFER_INFO csbiInfo;
	HANDLE hStdout = GetStdHandle(STD_OUTPUT_HANDLE);

	if (!GetConsoleScreenBufferInfo(hStdout, &csbiInfo))
	{
		return 80; // Return this if stdout is not a Console
	}

	return csbiInfo.srWindow.Right - csbiInfo.srWindow.Left;

#endif // WIN32

#endif // Can determine width of terminal

}

__noreturn static void showProducts(PRINTFUNC p)
{
	int cols = getLineWidth();
	int itemsPerLine;
	uint8_t i;
	int32_t index;


	uint_fast8_t longestString = 0;
	int32_t k, items = KmsData->SkuItemCount;

	p("You may use these product names or numbers:\n\n");

	for (index = 0; index < KmsData->SkuItemCount; index++)
	{
		uint_fast8_t len = (uint_fast8_t)strlen(KmsData->SkuItemList[index].Name);
		if (len > longestString) longestString = len;
	}

	itemsPerLine = cols / (longestString + 10);
	if (!itemsPerLine) itemsPerLine = 1;
	uint8_t lines = items / itemsPerLine;
	if (items % itemsPerLine) lines++;

	for (i = 0; i < lines; i++)
	{
		for (k = 0; k < itemsPerLine; k++)
		{
			uint8_t j;
			index = k * lines + i;

			if (index >= items) break;

			p("%3u = %s", index + 1, KmsData->SkuItemList[index].Name);

			for (j = 0; j < longestString + 4 - strlen(KmsData->SkuItemList[index].Name); j++)
			{
				p(" ");
			}
		}

		p("\n");
	}

	p("\n");

	exit(0);
}

__noreturn static void examples(const char* const programName)
{
	printf(
		"\n使用V4协议请求激活的 Office 2013 192.168.1.5:1688\n"
		"\t%s -l \"Office 2013 Professional\" -4 192.168.1.5\n"
		"\t%s -l \"Office 2013 Professional\" -4 192.168.1.5:1688\n\n"

		"使用V4协议请求激活 Windows Server 2012 localhost:1688\n"
		"\t%s -4 -l \"Windows Server 2012\" -k 8665cb71-468c-4aa3-a337-cb9bc9d5eaac\n"
		"\t%s -4 -l \"Windows Server 2012\"\n"
		"\t%s -4 -l \"Windows Server 2012\" [::1]:1688\n"
		"\t%s -4 -l \"Windows Server 2012\" 127.0.0.2:1688\n\n"

		"发送 100,000 次请求 to localhost:1688\n"
		"\t%s -n 100000\n\n"

		"从10.0.0.1:4711请求激活Windows 8并假装是史蒂夫鲍尔默\n"
		"\t%s -l \"Windows 8 Professional\" -w steveb1.redmond.microsoft.com 10.0.0.1:4711\n\n",
		programName, programName, programName, programName, programName, programName, programName, programName
	);

	exit(0);
}

#else // NO_HELP


__noreturn static void clientUsage(const char* const programName)
{
	errorout("Incorrect parameter specified.\n");
	exit(VLMCSD_EINVAL);
}


#endif // NO_HELP


static void parseProtocolVersion(void)
{
	char *endptr_major, *endptr_minor, *period = strchr(optarg, (int)'.');

	if (!period)
	{
		errorout("Fatal: Protocol version must be in the format #.#\n");
		exit(VLMCSD_EINVAL);
	}

	long major = strtol(optarg, &endptr_major, 10);
	long minor = strtol(period + 1, &endptr_minor, 10);

	if ((*endptr_major && *endptr_major != '.') || *endptr_minor || *optarg == '.' || !period[1])
	{
		errorout("Fatal: Protocol version must be in the format #.#\n");
		exit(VLMCSD_EINVAL);
	}

	if (major < 0 || major > 0xffff || minor < 0 || minor > 0xffff)
	{
		errorout("Fatal: Major and minor protocol version number must be between 0 and 65535\n");
		exit(VLMCSD_EINVAL);
	}

	MajorVersion = (uint16_t)major;
	MinorVersion = (uint16_t)minor;
}


static int32_t findLicensePackByName(const char* const name)
{
	int32_t i;

	for (i = KmsData->SkuItemCount - 1; i >= 0; i--)
	{
		if (!strcasecmp(name, KmsData->SkuItemList[i].Name)) return i;
	}

	return i;
}

static const char* const client_optstring = "+N:B:i:j:l:a:s:k:c:w:r:n:t:g:G:o:K:pPTv456mexdV";


//We handle only "-j". Many other options do not run without a loaded database
static void parseCommandLinePass0(const int argc, CARGV argv)
{
	int o;
	optReset();

	for (opterr = 0; (o = getopt(argc, (char* const*)argv, client_optstring)) > 0; ) switch (o)
	{
#	ifndef NO_EXTERNAL_DATA
	case 'j': // Set "License Pack" and protocol version (e.g. Windows8, Office2013v5, ...)
		fn_data = optarg;
#		ifndef NO_INTERNAL_DATA
		ExplicitDataLoad = TRUE;
#		endif // NO_INTERNAL_DATA
		break;
#	endif // NO_EXTERNAL_DATA

	default:
		break;
	}
}

//First pass. We handle only "-l". Since -a -k -s -4 -5 and -6 are exceptions to -l, we process -l first
static void parseCommandLinePass1(const int argc, CARGV argv)
{
	int o;
	optReset();

	for (opterr = 0; (o = getopt(argc, (char* const*)argv, client_optstring)) > 0; ) switch (o)
	{
	case 'l': // Set "License Pack" and protocol version (e.g. Windows8, Office2013v5, ...)
		if (stringToInt(optarg, 1, KmsData->SkuItemCount, (unsigned int*)&ActiveProductIndex))
		{
			ActiveProductIndex--;
			break;
		}

		ActiveProductIndex = findLicensePackByName(optarg);
		if (ActiveProductIndex < 0)
		{
			errorout("Invalid client application. \"%s\" is not valid for -l.\n\n", optarg);
#ifndef NO_HELP
			showProducts(&errorout);
#endif // !NO_HELP
		}

		break;

	default:
		break;
	}

	int32_t kmsIndex = KmsData->SkuItemList[ActiveProductIndex].KmsIndex;
	int32_t appIndex = KmsData->SkuItemList[ActiveProductIndex].AppIndex;

	MajorVersion = (uint16_t)KmsData->SkuItemList[ActiveProductIndex].ProtocolVersion;
	NCountPolicy = (uint32_t)KmsData->SkuItemList[ActiveProductIndex].NCountPolicy;
	memcpy(&SkuGuid, &KmsData->SkuItemList[ActiveProductIndex].Guid, sizeof(GUID));
	memcpy(&KmsGuid, &KmsData->KmsItemList[kmsIndex].Guid, sizeof(GUID));
	memcpy(&AppGuid, &KmsData->AppItemList[appIndex].Guid, sizeof(GUID));
}


// Second Pass. Handle all options except "-l"
static void parseCommandLinePass2(const char *const programName, const int argc, CARGV argv)
{
	int o;
	optReset();

	for (opterr = 0; (o = getopt(argc, (char* const*)argv, client_optstring)) > 0; ) switch (o)
	{
#ifndef NO_HELP

	case 'j':
		break;

	case 'e': // Show examples

		examples(programName);

	case 'x': // Show Apps

		showProducts(&printf);

#endif // NO_HELP

#			ifndef NO_DNS

	case 'P':

		NoSrvRecordPriority = TRUE;
		break;

#			endif // NO_DNS

	case 'G':

		incompatibleOptions |= VLMCS_OPTION_GRAB_INI;
		fn_ini_client = optarg;
		break;

#	ifndef USE_MSRPC

	case 'N':
		if (!getArgumentBool(&UseClientRpcNDR64, optarg)) clientUsage(programName);
		break;

	case 'B':
		if (!getArgumentBool(&UseClientRpcBTFN, optarg)) clientUsage(programName);
		break;

	case 'i':

		switch (getOptionArgumentInt((char)o, 4, 6))
		{
		case 4:
			AddressFamily = AF_INET;
			break;
		case 6:
			AddressFamily = AF_INET6;
			break;
		default:
			errorout("IPv5 不存在.\n");
			exit(VLMCSD_EINVAL);
		}

		break;

	case 'p': // Multiplexed RPC

		UseMultiplexedRpc = FALSE;
		break;

#	endif // USE_MSRPC

	case 'n': // Fixed number of Requests (regardless, whether they are required)

		incompatibleOptions |= VLMCS_OPTION_NO_GRAB_INI;
		FixedRequests = getOptionArgumentInt((char)o, 1, INT_MAX);
		break;

	case 'r': // Fake minimum required client count

		incompatibleOptions |= VLMCS_OPTION_NO_GRAB_INI;
		NCountPolicy = getOptionArgumentInt((char)o, 0, INT_MAX);
		break;

	case 'c': // use a specific client GUID

		// If using a constant Client ID, send only one request unless /N= explicitly specified
		if (!FixedRequests) FixedRequests = 1;

		CMID = optarg;
		break;

	case 'o': // use a specific previous client GUID

		CMID_prev = optarg;
		break;

	case 'a': // Set specific App Id

		incompatibleOptions |= VLMCS_OPTION_NO_GRAB_INI;
		string2UuidOrExit(optarg, &AppGuid);
		break;

	case 'g': // Set custom "grace" time in minutes (default 30 days)

		BindingExpiration = getOptionArgumentInt((char)o, 0, INT_MAX);
		break;

	case 's': // Set specfic SKU ID

		incompatibleOptions |= VLMCS_OPTION_NO_GRAB_INI;
		string2UuidOrExit(optarg, &SkuGuid);
		break;

	case 'k': // Set specific KMS ID

		incompatibleOptions |= VLMCS_OPTION_NO_GRAB_INI;
		string2UuidOrExit(optarg, &KmsGuid);
		break;

	case '4': // Force V4 protocol
	case '5': // Force V5 protocol
	case '6': // Force V5 protocol

		incompatibleOptions |= VLMCS_OPTION_NO_GRAB_INI;
		MajorVersion = o - 0x30;
		MinorVersion = 0;
		break;

	case 'K': // Use specific protocol (may be invalid)

		parseProtocolVersion();
		break;

	case 'd': // Don't use DNS names

		dnsnames = FALSE;
		break;

#			ifndef NO_VERBOSE_LOG

	case 'v': // Be verbose

		verbose = TRUE;
		break;

#			endif // NO_VERBOSE_LOG

	case 'm': // Pretend to be a virtual machine

		VMInfo = TRUE;
		break;

	case 'w': // WorkstationName (max. 63 chars)

		WorkstationName = optarg;

		if (strlen(WorkstationName) > 63)
		{
			errorout("\007WARNING! 截断工作站名称到63个字符 (%s).\n", WorkstationName);
		}

		break;

	case 't':

		LicenseStatus = getOptionArgumentInt((char)o, 0, 0x7fffffff);
		if ((unsigned int)LicenseStatus > 6) errorout("Warning: 正确的许可证状态为0 <=许可证状态<= 6.\n");
		break;

#			ifndef USE_MSRPC

	case 'T':

		ReconnectForEachRequest = TRUE;
		break;

#			endif // USE_MSRPC

	case 'l':
		incompatibleOptions |= VLMCS_OPTION_NO_GRAB_INI;
		break;

#	ifndef NO_VERSION_INFORMATION

	case 'V':
#				if defined(__s390__) && !defined(__zarch__) && !defined(__s390x__)
		printf("vlmcs %s %i-bit\n", Version, sizeof(void*) == 4 ? 31 : (int)sizeof(void*) << 3);
#				else
		printf("vlmcs %s %i-bit\n", Version, (int)sizeof(void*) << 3);
#				endif // defined(__s390__) && !defined(__zarch__) && !defined(__s390x__)
		printPlatform();
		printCommonFlags();
		printClientFlags();
		exit(0);

#			endif // NO_VERSION_INFORMATION

	default:
		clientUsage(programName);
	}
	if ((incompatibleOptions & (VLMCS_OPTION_NO_GRAB_INI | VLMCS_OPTION_GRAB_INI)) == (VLMCS_OPTION_NO_GRAB_INI | VLMCS_OPTION_GRAB_INI))
		clientUsage(programName);
}


///*
// * Compares 2 GUIDs where one is host-endian and the other is little-endian (network byte order)
// */
//int_fast8_t IsEqualGuidLEHE(const GUID* const guid1, const GUID* const guid2)
//{
//	GUID tempGuid;
//	LEGUID(&tempGuid, guid2);
//	return IsEqualGUID(guid1, &tempGuid);
//}


#ifndef USE_MSRPC
static void checkRpcLevel(const REQUEST* request, RESPONSE* response)
{
	if (!RpcFlags.HasNDR32)
		errorout("\nWARNING: 服务器的RPC协议不支持NDR32.\n");

	if (UseClientRpcBTFN && UseClientRpcNDR64 && RpcFlags.HasNDR64 && !RpcFlags.HasBTFN)
		errorout("\nWARNING: 服务器的RPC协议具有NDR64但没有BTFN.\n");

	//#	ifndef NO_BASIC_PRODUCT_LIST
	//	if (!IsEqualGuidLEHE(&request->KMSID, &ProductList[15].guid) && UseClientRpcBTFN && !RpcFlags.HasBTFN)
	//		errorout("\nWARNING: A server with pre-Vista RPC activated a product other than Office 2010.\n");
	//#	endif // NO_BASIC_PRODUCT_LIST
}
#endif // USE_MSRPC


static void displayResponse(const RESPONSE_RESULT result, const REQUEST* request, RESPONSE* response, BYTE *hwid)
{
	fflush(stdout);

	if (!result.RpcOK)				errorout("\n\007ERROR:	非零RPC结果代码.\n");
	if (!result.DecryptSuccess)		errorout("\n\007ERROR: V5 / V6响应的解密失败.\n");
	if (!result.IVsOK)				errorout("\n\007ERROR: AES CBC 初始化向量(IVs) 不匹配.\n");
	if (!result.PidLengthOK)		errorout("\n\007ERROR: PID的长度无效.\n");
	if (!result.HashOK)				errorout("\n\007ERROR: 计算的哈希与响应中的哈希不匹配.\n");
	if (!result.ClientMachineIDOK)	errorout("\n\007ERROR:客户端计算机请求和响应的GUID不匹配.\n");
	if (!result.TimeStampOK)		errorout("\n\007ERROR: 请求和响应的时间戳不匹配.\n");
	if (!result.VersionOK)			errorout("\n\007ERROR: 请求和响应的协议版本不匹配.\n");
	if (!result.HmacSha256OK)		errorout("\n\007ERROR: 密钥哈希消息身份验证代码（HMAC）不正确.\n");
	if (!result.IVnotSuspicious)	errorout("\nWARNING: The KMS server is an emulator because the response uses an IV following KMSv5 rules in KMSv6 protocol.\n");

	if (result.effectiveResponseSize != result.correctResponseSize)
	{
		errorout("\n\007WARNING: Size of RPC payload (KMS Message) should be %u but is %u.", result.correctResponseSize, result.effectiveResponseSize);
	}

#	ifndef USE_MSRPC
	checkRpcLevel(request, response);
#	endif // USE_MSRPC

	if (!result.DecryptSuccess) return; // Makes no sense to display anything

	char ePID[3 * PID_BUFFER_SIZE];
	if (!ucs2_to_utf8(response->KmsPID, ePID, PID_BUFFER_SIZE, 3 * PID_BUFFER_SIZE))
	{
		memset(ePID + 3 * PID_BUFFER_SIZE - 3, 0, 3);
	}

	// Read KMSPID from Response
#	ifndef NO_VERBOSE_LOG
	if (!verbose)
#	endif // NO_VERBOSE_LOG
	{
		printf(" -> %s", ePID);

		if (LE16(response->MajorVer) > 5)
		{
#			ifndef _WIN32
			printf(" (%016llX)", (unsigned long long)BE64(*(uint64_t*)hwid));
#			else // _WIN32
			printf(" (%016I64X)", (unsigned long long)BE64(*(uint64_t*)hwid));
#			endif // _WIN32
		}

		printf("\n");
	}
#	ifndef NO_VERBOSE_LOG
	else
	{
		printf(
			"\n\n来自KMS服务器的响应\n========================\n\n"
			"Size of KMS Response            : %u (0x%x)\n", result.effectiveResponseSize, result.effectiveResponseSize
		);

		logResponseVerbose(ePID, hwid, response, &printf);
		printf("\n");
	}
#	endif // NO_VERBOSE_LOG
}


static void connectRpc(RpcCtx *s)
{
#	ifdef NO_DNS

	RpcDiag_t rpcDiag;

	*s = connectToAddress(RemoteAddr, AddressFamily, FALSE);
	if (*s == INVALID_RPCCTX)
	{
		errorout("Fatal: Could not connect to %s\n", RemoteAddr);
		exit(SOCKET_ECONNABORTED);
	}

	if (verbose)
		printf("\nPerforming RPC bind ...\n");

	RpcStatus status;
	if ((status = rpcBindClient(*s, verbose, &rpcDiag)))
	{
		errorout("Fatal: Could not bind RPC\n");
		exit(status);
	}

	if (verbose) printf("... successful\n");

#	else // DNS

	static kms_server_dns_ptr* serverlist = NULL;
	static int numServers = 0;
	//static int_fast8_t ServerListAlreadyPrinted = FALSE;
	int i;

	if (!strcmp(RemoteAddr, "-") || *RemoteAddr == '.') // Get KMS server via DNS SRV record
	{
		if (!serverlist)
			numServers = getKmsServerList(&serverlist, RemoteAddr);

		if (numServers < 1)
		{
			errorout("Fatal: 找不到KMS服务器\n");
			exit(SOCKET_ECONNABORTED);
		}

		if (!NoSrvRecordPriority) sortSrvRecords(serverlist, numServers);

#		ifndef NO_VERBOSE_LOG
		if (verbose /*&& !ServerListAlreadyPrinted*/)
		{
			for (i = 0; i < numServers; i++)
			{
				printf(
					"Found %-40s (priority: %hu, weight: %hu, randomized weight: %i)\n",
					serverlist[i]->serverName,
					serverlist[i]->priority, serverlist[i]->weight,
					NoSrvRecordPriority ? 0 : serverlist[i]->random_weight
				);
			}

			printf("\n");
			//ServerListAlreadyPrinted = TRUE;
		}
#		endif // NO_VERBOSE_LOG
	}
	else // Just use the server supplied on the command line
	{
		if (!serverlist)
		{
			serverlist = (kms_server_dns_ptr*)vlmcsd_malloc(sizeof(kms_server_dns_ptr));
			*serverlist = (kms_server_dns_ptr)vlmcsd_malloc(sizeof(kms_server_dns_t));

			numServers = 1;
			strncpy((*serverlist)->serverName, RemoteAddr, sizeof((*serverlist)->serverName));
		}
	}

	for (i = 0; i < numServers; i++)
	{
		*s = connectToAddress(serverlist[i]->serverName, AddressFamily, (*RemoteAddr == '.' || *RemoteAddr == '-'));

		if (*s == INVALID_RPCCTX) continue;
		RpcDiag_t rpcDiag;

#		ifndef NO_VERBOSE_LOG
		if (verbose) printf("\n执行RPC绑定 ...\n");

		if (rpcBindClient(*s, verbose, &rpcDiag))
#		else
		if (rpcBindClient(*s, FALSE, &rpcDiag))
#		endif
		{
			errorout("Warning:无法绑定RPC\n");
			continue;
		}

#		ifndef NO_VERBOSE_LOG
		if (verbose) printf("... 成功\n");
#		endif

		return;
	}

	errorout("Fatal: 无法连接到任何KMS服务器\n");
	exit(SOCKET_ECONNABORTED);

#	endif // DNS
}

#endif // IS_LIBRARY

int SendActivationRequest(const RpcCtx sock, RESPONSE *baseResponse, REQUEST *baseRequest, RESPONSE_RESULT *result, BYTE *const hwid)
{
	size_t requestSize, responseSize;
	BYTE *request, *response;
	int status;

	result->mask = 0;

	if (LE16(baseRequest->MajorVer) < 5)
		request = CreateRequestV4(&requestSize, baseRequest);
	else
		request = CreateRequestV6(&requestSize, baseRequest);

	if (!((status = rpcSendRequest(sock, request, requestSize, &response, &responseSize))))
	{
		if (LE16(((RESPONSE*)(response))->MajorVer) == 4)
		{
			RESPONSE_V4 response_v4;
			*result = DecryptResponseV4(&response_v4, (const int)responseSize, response, request);
			memcpy(baseResponse, &response_v4.ResponseBase, sizeof(RESPONSE));
		}
		else
		{
			RESPONSE_V6 response_v6;
			*result = DecryptResponseV6(&response_v6, (int)responseSize, response, request, hwid);
			memcpy(baseResponse, &response_v6.ResponseBase, sizeof(RESPONSE));
		}

		result->RpcOK = TRUE;
	}

	if (response) free(response);
	free(request);
	return status;
}

#ifndef IS_LIBRARY

static int sendRequest(RpcCtx *const s, REQUEST *const request, RESPONSE *const response, hwid_t hwid, RESPONSE_RESULT *const result)
{
	CreateRequestBase(request);

	if (*s == INVALID_RPCCTX)
		connectRpc(s);
	else
	{
		// Check for lame KMS emulators that close the socket after each request
		int_fast8_t disconnected = isDisconnected(*s);

		if (disconnected)
			errorout("\nWarning: 服务器关闭RPC连接（可能是非多任务KMS模拟器）\n");

		if (ReconnectForEachRequest || disconnected)
		{
			closeRpc(*s);
			connectRpc(s);
		}
	}

	printf("发送激活请求 (KMS V%u) ", MajorVersion);
	fflush(stdout);

	return SendActivationRequest(*s, response, request, result, hwid);
}


static void displayRequestError(RpcCtx *const s, const int status, const int currentRequest, const int totalRequests)
{
	errorout("\nError 0x%08X while sending request %u of %u\n", status, currentRequest, RequestsToGo + totalRequests);

	switch (status)
	{
	case 0xC004F042: // not licensed
		errorout("The KMS server has declined to activate the requested product\n");
		break;

	case 0x8007000D:  // e.g. v6 protocol on a v5 server
		errorout("The KMS host you are using is unable to handle your product. It only supports legacy versions\n");
		break;

	case 0xC004F06C:
		errorout("The time stamp differs too much from the KMS server time\n");
		break;

	case 0xC004D104:
		errorout("The security processor reported that invalid data was used\n");
		break;

	case 1:
		errorout("An RPC protocol error has occured\n");
		closeRpc(*s);
		connectRpc(s);
		break;

	default:
#		if _WIN32
		errorout("%s\n", win_strerror(status));
#		endif // _WIN32
		break;
	}
}


static void newIniBackupFile(const char* const restrict fname)
{
	FILE *restrict f = fopen(fname, "wb");

	if (!f)
	{
		int error = errno;
		errorout("Fatal: Cannot create %s: %s\n", fname, strerror(error));
		exit(error);
	}

	if (fclose(f))
	{
		int error = errno;
		errorout("Fatal: Cannot write to %s: %s\n", fname, strerror(error));
		vlmcsd_unlink(fname);
		exit(error);
	}
}


static void updateIniFile(iniFileEpidLines* const restrict lines)
{
	int_fast8_t lineWritten[vlmcsd_countof(*lines)];
#	if !_MSC_VER
	struct stat statbuf;
#	endif
	uint_fast8_t i;
	int_fast8_t iniFileExistedBefore = TRUE;
	unsigned int lineNumber;

	memset(lineWritten, FALSE, sizeof(lineWritten));

	char* restrict fn_bak = (char*)vlmcsd_malloc(strlen(fn_ini_client) + 2);

	strcpy(fn_bak, fn_ini_client);
	strcat(fn_bak, "~");

#	if _MSC_VER
	if (!PathFileExists(fn_ini_client))
	{
		iniFileExistedBefore = FALSE;
		newIniBackupFile(fn_bak);
	}
#	else
	if (stat(fn_ini_client, &statbuf))
	{
		if (errno != ENOENT)
		{
			int error = errno;
			errorout("Fatal: %s: %s\n", fn_ini_client, strerror(error));
			exit(error);
		}
		else
		{
			iniFileExistedBefore = FALSE;
			newIniBackupFile(fn_bak);
		}
	}
#	endif
	else
	{
		vlmcsd_unlink(fn_bak); // Required for Windows. Most Unix systems don't need it.
		if (rename(fn_ini_client, fn_bak))
		{
			int error = errno;
			errorout("Fatal: Cannot create %s: %s\n", fn_bak, strerror(error));
			exit(error);
		}
	}

	printf("\n%s file %s\n", iniFileExistedBefore ? "Updating" : "Creating", fn_ini_client);

	FILE *restrict in, *restrict out;

	in = fopen(fn_bak, "rb");

	if (!in)
	{
		int error = errno;
		errorout("Fatal: Cannot open %s: %s\n", fn_bak, strerror(error));
		exit(error);
	}

	out = fopen(fn_ini_client, "wb");

	if (!out)
	{
		int error = errno;
		errorout("Fatal: Cannot create %s: %s\n", fn_ini_client, strerror(error));
		exit(error);
	}

	char sourceLine[256];

	for (lineNumber = 1; fgets(sourceLine, sizeof(sourceLine), in); lineNumber++)
	{
		for (i = 0; i < vlmcsd_countof(*lines); i++)
		{
			if (*(*lines)[i] && !strncasecmp(sourceLine, (*lines)[i], strlen(ePidGroup[i])))
			{
				if (lineWritten[i]) break;

				fprintf(out, "%s", (*lines)[i]);
				printf("line %2i: %s", lineNumber, (*lines)[i]);
				lineWritten[i] = TRUE;
				break;
			}
		}

		if (i >= vlmcsd_countof(*lines))
		{
			fprintf(out, "%s", sourceLine);
		}
	}

	if (ferror(in))
	{
		int error = errno;
		errorout("Fatal: Cannot read from %s: %s\n", fn_bak, strerror(error));
		exit(error);
	}

	fclose(in);

	for (i = 0; i < vlmcsd_countof(*lines); i++)
	{
		if (!lineWritten[i] && *(*lines)[i])
		{
			fprintf(out, "%s", (*lines)[i]);
			printf("line %2i: %s", lineNumber + i, (*lines)[i]);
		}
	}

	if (fclose(out))
	{
		int error = errno;
		errorout("Fatal: Cannot write to %s: %s\n", fn_ini_client, strerror(error));
		exit(error);
	}

	if (!iniFileExistedBefore) vlmcsd_unlink(fn_bak);

	free(fn_bak);
}


static void grabServerData()
{
	RpcCtx s = INVALID_RPCCTX;
	WORD MajorVer = 6;
	iniFileEpidLines lines;

	static char* Licenses[vlmcsd_countof(lines)] =
	{
		(char*)"212a64dc-43b1-4d3d-a30c-2fc69d2095c6", // Vista
		(char*)"e85af946-2e25-47b7-83e1-bebcebeac611", // Office 2010
		(char*)"e6a6f1bf-9d40-40c3-aa9f-c77ba21578c0", // Office 2013
		(char*)"85b5f61b-320b-4be3-814a-b76b2bfafc82", // Office 2016
	};

	uint_fast8_t i;
	int32_t j;
	RESPONSE response;
	RESPONSE_RESULT result;
	REQUEST request;
	hwid_t hwid;
	int status;
	size_t len;

	for (i = 0; i < vlmcsd_countof(lines); i++) *lines[i] = 0;

	for (i = 0; i < vlmcsd_countof(Licenses) && MajorVer > 3; i++)
	{
		GUID guid;
		string2UuidLE(Licenses[i], &guid);
		int32_t kmsIndex = getProductIndex(&guid, KmsData->KmsItemList, KmsData->KmsItemCount, NULL, NULL);

		if (kmsIndex < 0)
		{
			errorout("Warning: KMS GUID %s not in database.\n", Licenses[i]);
			continue;
		}

		ActiveProductIndex = ~0;

		for (j = KmsData->SkuItemCount; j >= 0; j--)
		{
			if (KmsData->SkuItemList[j].KmsIndex == kmsIndex)
			{
				ActiveProductIndex = j;
				break;
			}
		}

		if (ActiveProductIndex == ~0)
		{
			errorout("Warning: KMS GUID %s not in database.\n", Licenses[i]);
			continue;
		}

		int32_t appIndex = KmsData->SkuItemList[ActiveProductIndex].AppIndex;

		NCountPolicy = (uint32_t)KmsData->SkuItemList[ActiveProductIndex].NCountPolicy;
		memcpy(&SkuGuid, &KmsData->SkuItemList[ActiveProductIndex].Guid, sizeof(GUID));
		memcpy(&KmsGuid, &KmsData->KmsItemList[kmsIndex].Guid, sizeof(GUID));
		memcpy(&AppGuid, &KmsData->AppItemList[appIndex].Guid, sizeof(GUID));
		MajorVersion = (uint16_t)MajorVer;

		status = sendRequest(&s, &request, &response, hwid, &result);
		printf("%-11s", ePidGroup[i]);

		if (status)
		{
			displayRequestError(&s, status, i + 7 - MajorVer, 9 - MajorVer);

			if (status == 1) break;

			if ((status & 0xF0000000) == 0x80000000)
			{
				MajorVer--;
				i--;
			}

			continue;
		}

		printf("%i of %i", (int)(i + 7 - MajorVer), (int)(10 - MajorVer));
		displayResponse(result, &request, &response, hwid);

		char ePID[3 * PID_BUFFER_SIZE];

		if (!ucs2_to_utf8(response.KmsPID, ePID, PID_BUFFER_SIZE, 3 * PID_BUFFER_SIZE))
		{
			memset(ePID + 3 * PID_BUFFER_SIZE - 3, 0, 3);
		}

		vlmcsd_snprintf(lines[i], sizeof(lines[0]), "%s = %s", ePidGroup[i], ePID);

		if (response.MajorVer > 5)
		{
			len = strlen(lines[i]);
			vlmcsd_snprintf(lines[i] + len, sizeof(lines[0]) - len, " / %02X %02X %02X %02X %02X %02X %02X %02X", hwid[0], hwid[1], hwid[2], hwid[3], hwid[4], hwid[5], hwid[6], hwid[7]);
		}

		len = strlen(lines[i]);
		vlmcsd_snprintf(lines[i] + len, sizeof(lines[0]) - len, "\n");

	}

	if (strcmp(fn_ini_client, "-"))
	{
		updateIniFile(&lines);
	}
	else
	{
		printf("\n");
		for (i = 0; i < vlmcsd_countof(lines); i++) printf("%s", lines[i]);
	}
}


int client_main(int argc, CARGV argv)
{
#if defined(_WIN32) && !defined(USE_MSRPC)

	// Windows Sockets must be initialized

	WSADATA wsadata;
	int error;

	if ((error = WSAStartup(0x0202, &wsadata)))
	{
		errorout("Fatal: Could not initialize Windows sockets (Error: %d).\n", error);
		return error;
	}

#endif // _WIN32

#ifdef _NTSERVICE

	// We are not a service
	IsNTService = FALSE;

#endif // _NTSERVICE

	randomNumberInit();

	//#	ifndef NO_EXTERNAL_DATA
	//	ExplicitDataLoad = TRUE;
	//#	endif // NO_EXTERNAL_DATA

	parseCommandLinePass0(argc, argv);

	int_fast8_t useDefaultHost = FALSE;

	if (optind < argc)
		RemoteAddr = argv[optind];
	else
		useDefaultHost = TRUE;

	int hostportarg = optind;

	if (optind < argc - 1)
	{
		parseCommandLinePass0(argc - hostportarg, argv + hostportarg);

		if (optind < argc - hostportarg)
			clientUsage(argv[0]);
	}

	loadKmsData();

	if (!KmsData->AppItemCount || !KmsData->SkuItemCount || !KmsData->KmsItemCount)
	{
		errorout("Fatal: Incomplete KMS data file\n");
		exit(VLMCSD_EINVAL);
	}

	parseCommandLinePass1(argc, argv);

	if (optind < argc - 1)
	{
		parseCommandLinePass1(argc - hostportarg, argv + hostportarg);
	}

	parseCommandLinePass2(argv[0], argc, argv);

	if (optind < argc - 1)
		parseCommandLinePass2(argv[0], argc - hostportarg, argv + hostportarg);

	if (useDefaultHost)
	{
#	ifndef USE_MSRPC
		RemoteAddr = AddressFamily == AF_INET6 ? "::1" : "127.0.0.1";
#	else
		RemoteAddr = "127.0.0.1";
#	endif
	}

	if (fn_ini_client != NULL)
		grabServerData();
	else
	{
		int requests;
		RpcCtx s = INVALID_RPCCTX;

		for (requests = 0, RequestsToGo = NCountPolicy == 1 ? 1 : NCountPolicy - 1; RequestsToGo; requests++)
		{
			RESPONSE response;
			REQUEST request;
			RESPONSE_RESULT result;
			hwid_t hwid;

			int status = sendRequest(&s, &request, &response, hwid, &result);

			if (FixedRequests) RequestsToGo = FixedRequests - requests - 1;

			if (status)
			{
				displayRequestError(&s, status, requests + 1, RequestsToGo + requests + 1);
				if (!FixedRequests)	RequestsToGo = 0;
			}
			else
			{
				if (!FixedRequests)
				{
					if (firstRequestSent && NCountPolicy - (int)response.Count >= RequestsToGo)
					{
						errorout("\nThe KMS server does not increment it's active clients. Aborting...\n");
						RequestsToGo = 0;
					}
					else
					{
						RequestsToGo = NCountPolicy - response.Count;
						if (RequestsToGo < 0) RequestsToGo = 0;
					}
				}

				fflush(stderr);
				printf("%i of %i ", requests + 1, RequestsToGo + requests + 1);
				displayResponse(result, &request, &response, hwid);
				firstRequestSent = TRUE;
			}
		}
	}

	return 0;
}


// Create Base KMS Client Request
static void CreateRequestBase(REQUEST *Request)
{
	Request->MinorVer = LE16(MinorVersion);
	Request->MajorVer = LE16(MajorVersion);
	Request->VMInfo = LE32(VMInfo);
	Request->LicenseStatus = LE32(LicenseStatus);
	Request->BindingExpiration = LE32(BindingExpiration);
	Request->N_Policy = LE32(NCountPolicy);

	memcpy(&Request->ActID, &SkuGuid, sizeof(GUID));
	memcpy(&Request->KMSID, &KmsGuid, sizeof(GUID));
	memcpy(&Request->AppID, &AppGuid, sizeof(GUID));

	getUnixTimeAsFileTime(&Request->ClientTime);

	{
		if (CMID)
		{
			string2UuidOrExit(CMID, &Request->CMID);
		}
		else
		{
			get16RandomBytes(&Request->CMID);

			// Set reserved UUID bits
			Request->CMID.Data4[0] &= 0x3F;
			Request->CMID.Data4[0] |= 0x80;

			// Set UUID type 4 (random UUID)
			Request->CMID.Data3 &= LE16(0xfff);
			Request->CMID.Data3 |= LE16(0x4000);
		}

		if (CMID_prev)
		{
			string2UuidOrExit(CMID_prev, &Request->CMID_prev);
		}
		else
		{
			memset(&Request->CMID_prev, 0, sizeof(Request->CMID_prev));
		}
	}

	static const char alphanum[] = "0123456789" "ABCDEFGHIJKLMNOPQRSTUVWXYZ" /*"abcdefghijklmnopqrstuvwxyz" */;

	if (WorkstationName)
	{
		utf8_to_ucs2(Request->WorkstationName, WorkstationName, WORKSTATION_NAME_BUFFER, WORKSTATION_NAME_BUFFER * 3);
	}
	else if (dnsnames)
	{
		int len, len2;
		unsigned int index = rand() % vlmcsd_countof(ClientDnsNames.first);
		len = (int)utf8_to_ucs2(Request->WorkstationName, ClientDnsNames.first[index], WORKSTATION_NAME_BUFFER, WORKSTATION_NAME_BUFFER * 3);

		index = rand() % vlmcsd_countof(ClientDnsNames.second);
		len2 = (int)utf8_to_ucs2(Request->WorkstationName + len, ClientDnsNames.second[index], WORKSTATION_NAME_BUFFER, WORKSTATION_NAME_BUFFER * 3);

		index = rand() % vlmcsd_countof(ClientDnsNames.tld);
		utf8_to_ucs2(Request->WorkstationName + len + len2, ClientDnsNames.tld[index], WORKSTATION_NAME_BUFFER, WORKSTATION_NAME_BUFFER * 3);
	}
	else
	{
		unsigned int size = (rand() % 14) + 1;
		const unsigned char *dummy;
		unsigned int i;

		for (i = 0; i < size; i++)
		{
			Request->WorkstationName[i] = utf8_to_ucs2_char((unsigned char*)alphanum + (rand() % (sizeof(alphanum) - 1)), &dummy);
		}

		Request->WorkstationName[size] = 0;
	}

#	ifndef NO_VERBOSE_LOG
	if (verbose)
	{
		printf("\nRequest Parameters\n==================\n\n");
		logRequestVerbose(Request, &printf);
		printf("\n");
	}
#	endif // NO_VERBOSE_LOG
}

#if _MSC_VER && !defined(_DEBUG)&& !MULTI_CALL_BINARY
int __stdcall WinStartUp(void)
{
	WCHAR **szArgList;
	int argc;
	szArgList = CommandLineToArgvW(GetCommandLineW(), &argc);

	int i;
	char **argv = (char**)vlmcsd_malloc(sizeof(char*)*argc);

	for (i = 0; i < argc; i++)
	{
		int size = WideCharToMultiByte(CP_UTF8, 0, szArgList[i], -1, argv[i], 0, NULL, NULL);
		argv[i] = (char*)vlmcsd_malloc(size);
		WideCharToMultiByte(CP_UTF8, 0, szArgList[i], -1, argv[i], size, NULL, NULL);
	}

	exit(client_main(argc, argv));
}
#endif // _MSC_VER && !defined(_DEBUG)&& !MULTI_CALL_BINARY


#endif // IS_LIBRARY
