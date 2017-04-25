
#include "stdafx.h"

#include <stdio.h>
#include <wchar.h>
#include "RawPrinter-AddIn.h"
#include <string>

static wchar_t *g_PropNames[] = {L"PrinterName"};
static wchar_t *g_MethodNames[] = {L"OpenPrinter", L"ClosePrinter", L"SendRaw", L"StartDocument", L"EndDocument"};

static wchar_t *g_PropNamesRu[] = {L"ИмяПринтера"};
static wchar_t *g_MethodNamesRu[] = {L"Открыть", L"Закрыть", L"ОтправитьДанные", L"НачатьДокумент", L"ЗакончитьДокумент"};

static const wchar_t g_kClassNames[] = L"CAddInRawPrinter"; //"|OtherClass1|OtherClass2";
static IAddInDefBase *pAsyncEvent = NULL;

uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len = 0);
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len = 0);
uint32_t getLenShortWcharStr(const WCHAR_T* Source);
char* ConvToUtf8(const WCHAR_T *W);


//---------------------------------------------------------------------------//
long GetClassObject(const WCHAR_T* wsName, IComponentBase** pInterface)
{
	if(!*pInterface)
	{
		*pInterface= new CAddInRawPrinter;
		return (long)*pInterface;
	}
	return 0;
}
//---------------------------------------------------------------------------//
long DestroyObject(IComponentBase** pIntf)
{
	if(!*pIntf)
		return -1;

	delete *pIntf;
	*pIntf = 0;
	return 0;
}
//---------------------------------------------------------------------------//
const WCHAR_T* GetClassNames()
{
	static WCHAR_T* names = 0;
	if (!names)
		::convToShortWchar(&names, g_kClassNames);
	return names;
}
//---------------------------------------------------------------------------//

// CAddInRawPrinter
//---------------------------------------------------------------------------//
CAddInRawPrinter::CAddInRawPrinter()
{
	m_iMemory = 0;
	m_iConnect = 0;
	hPrinter = 0;

	logfile.open("C:\\Temp\\RawPrinterLog.txt", std::ios_base::trunc);
}
//---------------------------------------------------------------------------//
CAddInRawPrinter::~CAddInRawPrinter()
{
	logfile.close();
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::Init(void* pConnection)
{ 
	m_iConnect = (IAddInDefBase*)pConnection;
	return m_iConnect != NULL;
}
//---------------------------------------------------------------------------//
long CAddInRawPrinter::GetInfo()
{ 
	// Component should put supported component technology version 
	// This component supports 2.0 version
	return 2000; 
}
//---------------------------------------------------------------------------//
void CAddInRawPrinter::Done()
{
}
/////////////////////////////////////////////////////////////////////////////
// ILanguageExtenderBase
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
	wchar_t *wsExtension = L"RawPrinter";
	int iActualSize = ::wcslen(wsExtension) + 1;
	WCHAR_T* dest = 0;

	if (m_iMemory)
	{
		logfile << "RegisterExtension... ";
		if(m_iMemory->AllocMemory((void**)wsExtensionName, iActualSize * sizeof(WCHAR_T)))
			::convToShortWchar(wsExtensionName, wsExtension, iActualSize);
		logfile << "OK" << std::endl;
		return true;
	}

	return false; 
}
//---------------------------------------------------------------------------//
long CAddInRawPrinter::GetNProps()
{ 
	// You may delete next lines and add your own implementation code here
	return eProp_Last;
}
//---------------------------------------------------------------------------//
long CAddInRawPrinter::FindProp(const WCHAR_T* wsPropName)
{ 
	long plPropNum = -1;
	wchar_t* propName = 0;

	::convFromShortWchar(&propName, wsPropName);
	plPropNum = findName(g_PropNames, propName, eProp_Last);

	if (plPropNum == -1)
		plPropNum = findName(g_PropNamesRu, propName, eProp_Last);

	delete[] propName;

	return plPropNum;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInRawPrinter::GetPropName(long lPropNum, long lPropAlias)
{ 
	if (lPropNum >= eProp_Last)
		return NULL;

	wchar_t *wsCurrentName = NULL;
	WCHAR_T *wsPropName = NULL;
	int iActualSize = 0;

	switch(lPropAlias)
	{
	case 0: // First language
		wsCurrentName = g_PropNames[lPropNum];
		break;
	case 1: // Second language
		wsCurrentName = g_PropNamesRu[lPropNum];
		break;
	default:
		return 0;
	}

	iActualSize = wcslen(wsCurrentName)+1;

	if (m_iMemory && wsCurrentName)
	{
		logfile << "GetPropName...";
		if (m_iMemory->AllocMemory((void**)&wsPropName, iActualSize * sizeof(WCHAR_T)))
			::convToShortWchar(&wsPropName, wsCurrentName, iActualSize);
		logfile << "OK" << std::endl;
	}

	return wsPropName;
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::GetPropVal(const long lPropNum, tVariant* pvarPropVal)
{ 
	switch(lPropNum)
	{
	case eProp_PrinterName: 
		{
			logfile << "GetPropVal(PrinterName)...";
			int nameLength = wcslen(PrinterName);
			wchar_t *result = NULL;

			m_iMemory->AllocMemory(reinterpret_cast<void**>(&result), sizeof(wchar_t) * (nameLength + 1));
			wcscpy(result, PrinterName);

			TV_VT(pvarPropVal) = VTYPE_PWSTR;
			pvarPropVal->pwstrVal = result;

			logfile << "OK" << std::endl;
		}
		break;
	default:
		return false;
	}

	return true;
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::SetPropVal(const long lPropNum, tVariant *varPropVal)
{ 
	return false;
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::IsPropReadable(const long lPropNum)
{ 
	switch(lPropNum)
	{ 
	case eProp_PrinterName:
		return true;
	default:
		return false;
	}

	return false;
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::IsPropWritable(const long lPropNum)
{
	return false;
}
//---------------------------------------------------------------------------//
long CAddInRawPrinter::GetNMethods()
{ 
	return eMeth_Last;
}
//---------------------------------------------------------------------------//
long CAddInRawPrinter::FindMethod(const WCHAR_T* wsMethodName)
{ 
	long plMethodNum = -1;
	wchar_t* name = 0;

	::convFromShortWchar(&name, wsMethodName);

	plMethodNum = findName(g_MethodNames, name, eMeth_Last);

	if (plMethodNum == -1)
		plMethodNum = findName(g_MethodNamesRu, name, eMeth_Last);

	return plMethodNum;
}
//---------------------------------------------------------------------------//
const WCHAR_T* CAddInRawPrinter::GetMethodName(const long lMethodNum, const long lMethodAlias)
{ 
	if (lMethodNum >= eMeth_Last)
		return NULL;

	wchar_t *wsCurrentName = NULL;
	WCHAR_T *wsMethodName = NULL;
	int iActualSize = 0;

	switch(lMethodAlias)
	{
	case 0: // First language
		wsCurrentName = g_MethodNames[lMethodNum];
		break;
	case 1: // Second language
		wsCurrentName = g_MethodNamesRu[lMethodNum];
		break;
	default: 
		return 0;
	}

	iActualSize = wcslen(wsCurrentName)+1;

	if (m_iMemory && wsCurrentName)
	{
		logfile << "GetMethodName...";
		if(m_iMemory->AllocMemory((void**)&wsMethodName, iActualSize * sizeof(WCHAR_T)))
			::convToShortWchar(&wsMethodName, wsCurrentName, iActualSize);
		logfile << "OK" << std::endl;
	}

	return wsMethodName;
}
//---------------------------------------------------------------------------//
long CAddInRawPrinter::GetNParams(const long lMethodNum)
{ 
	switch(lMethodNum)
	{ 
	case eMeth_Open:
		return 1;
	case eMeth_Close:
		return 0;
	case eMeth_SendRaw:
		return 1;
	case eMeth_StartDocument:
		return 2;
	case eMeth_EndDocument:
		return 0;
	default:
		return 0;
	}

	return 0;
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::GetParamDefValue(const long lMethodNum, const long lParamNum,
	tVariant *pvarParamDefValue)
{ 
	TV_VT(pvarParamDefValue)= VTYPE_EMPTY;

	switch(lMethodNum)
	{ 
	case eMeth_Open:
	case eMeth_Close:
	case eMeth_SendRaw:
		// There are no parameter values by default 
		break;
	default:
		return false;
	}

	return false;
} 
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::HasRetVal(const long lMethodNum)
{ 
	return false;
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::CallAsProc(const long lMethodNum,
	tVariant* paParams, const long lSizeArray)
{ 
	switch(lMethodNum)
	{ 
	case eMeth_Open:
		{
			logfile << "CallAsProc(Open)...";
			if (hPrinter) {
				ClosePrinter(hPrinter);
				hPrinter = NULL;
			}

			WCHAR_T *m_PrinterName = paParams[0].pwstrVal;
			uint32_t len = paParams[0].wstrLen;
			uint32_t sz = sizeof(WCHAR_T)*(len + 1);

			if (PrinterName) {
				// TODO: Разобраться с освобождением памяти
				// m_iMemory->FreeMemory(reinterpret_cast<void**>(&PrinterName));
				PrinterName = NULL;
			}

			{
				m_iMemory->AllocMemory(reinterpret_cast<void**>(&PrinterName), sz);
			}
			memcpy(reinterpret_cast<void*>(PrinterName), reinterpret_cast<void*>(m_PrinterName), sz);
			PrinterName[len] = 0;


			wchar_t *wp_Name = NULL;
			::convFromShortWchar(&wp_Name, PrinterName, len + 1);

			if (!OpenPrinterW(wp_Name, &hPrinter, NULL)) {
				logfile << " <OpenPrinter Error> " << std::endl;

				wchar_t buf[512];
				wsprintf(buf, L"OpenPrinterW(%s) failed with code: %u", wp_Name, GetLastError());
				addError(ADDIN_E_FAIL, L"Printer error", buf, 1);
			}

			delete [] wp_Name;

			logfile << "OK" << std::endl;

		}
		break;

	case eMeth_Close:

		logfile << "CallAsProc(Close)...";

		if (PrinterName) {
			// TODO: разобраться с освобождением памяти
			// m_iMemory->FreeMemory(reinterpret_cast<void**>(&PrinterName));
			PrinterName = NULL;
		}
		ClosePrinter(hPrinter);

		logfile << "OK" << std::endl;

		break;

	case eMeth_SendRaw:

		{
			logfile << "CallAsProc(SendRaw)...";

			WCHAR_T *wc = paParams[0].pwstrVal;
			char *utf8 = ConvToUtf8(wc);

			DWORD len = strlen(utf8), sent;
			if (!WritePrinter(hPrinter, utf8, len, &sent)) {
				logfile << "<WritePrinter Error>";
				addError(ADDIN_E_FAIL, L"Failed to send data to printer!", L"Failed!", 2);
			}

			delete [] utf8;

			logfile << "OK" << std::endl;

		}
		break;

	case eMeth_StartDocument:

		{

			logfile << "CallAsProc(StartDocument)...";

			wchar_t *doc_name = NULL;
			wchar_t *data_type = NULL;
			::convFromShortWchar(&doc_name, paParams[0].pwstrVal);
			::convFromShortWchar(&data_type, paParams[1].pwstrVal);

			DOC_INFO_1 doc;
			doc.pDocName = doc_name;
			doc.pOutputFile = NULL;
			doc.pDatatype = data_type;

			DWORD print_job = StartDocPrinter(hPrinter, 1, (LPBYTE)&doc);

			delete [] doc_name;
			delete [] data_type;

			logfile << "OK" << std::endl;
		}

		break;

	case eMeth_EndDocument:

		logfile << "CallAsProc(EndDocument)...";
		EndDocPrinter(hPrinter);
		logfile << "OK" << std::endl;

		break;

	default:
		return false;
	}

	return true;
}
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::CallAsFunc(const long lMethodNum,
	tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray)
{ 
	return false;
}
//---------------------------------------------------------------------------//
//---------------------------------------------------------------------------//
void CAddInRawPrinter::SetLocale(const WCHAR_T* loc)
{
#ifndef __linux__
	_wsetlocale(LC_ALL, loc);
#else
	//We convert in char* char_locale
	//also we establish locale
	//setlocale(LC_ALL, char_locale);
#endif
}
/////////////////////////////////////////////////////////////////////////////
// LocaleBase
//---------------------------------------------------------------------------//
bool CAddInRawPrinter::setMemManager(void* mem)
{
	m_iMemory = (IMemoryManager*)mem;
	return m_iMemory != 0;
}
//---------------------------------------------------------------------------//
void CAddInRawPrinter::addError(uint32_t wcode, const wchar_t* source, 
	const wchar_t* descriptor, long code)
{
	if (m_iConnect)
	{
		WCHAR_T *err = 0;
		WCHAR_T *descr = 0;

		::convToShortWchar(&err, source);
		::convToShortWchar(&descr, descriptor);

		m_iConnect->AddError(wcode, err, descr, code);
		delete[] err;
		delete[] descr;
	}
}
//---------------------------------------------------------------------------//
long CAddInRawPrinter::findName(wchar_t* names[], const wchar_t* name, 
	const uint32_t size) const
{
	long ret = -1;
	for (uint32_t i = 0; i < size; i++)
	{
		if (!wcscmp(names[i], name))
		{
			ret = i;
			break;
		}
	}
	return ret;
}
//---------------------------------------------------------------------------//
uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len)
{
	if (!len)
		len = ::wcslen(Source)+1;

	if (!*Dest)
		*Dest = new WCHAR_T[len];

	WCHAR_T* tmpShort = *Dest;
	wchar_t* tmpWChar = (wchar_t*) Source;
	uint32_t res = 0;

	::memset(*Dest, 0, len*sizeof(WCHAR_T));
	do
	{
		*tmpShort++ = (WCHAR_T)*tmpWChar++;
		++res;
	}
	while (len-- && *tmpWChar);

	return res;
}
//---------------------------------------------------------------------------//
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len)
{
	if (!len)
		len = getLenShortWcharStr(Source)+1;

	if (!*Dest)
		*Dest = new wchar_t[len];

	wchar_t* tmpWChar = *Dest;
	WCHAR_T* tmpShort = (WCHAR_T*)Source;
	uint32_t res = 0;

	::memset(*Dest, 0, len*sizeof(wchar_t));
	do
	{
		*tmpWChar++ = (wchar_t)*tmpShort++;
		++res;
	}
	while (len-- && *tmpShort);
	tmpWChar = 0;

	return res;
}
//---------------------------------------------------------------------------//
uint32_t getLenShortWcharStr(const WCHAR_T* Source)
{
	uint32_t res = 0;
	WCHAR_T *tmpShort = (WCHAR_T*)Source;

	while (*tmpShort++)
		++res;

	return res;
}
//---------------------------------------------------------------------------//
uint32_t GetUtf8Size(const WCHAR_T *W)
{
	const WCHAR_T *c = W;
	size_t prelen = 0;
	for (; *c; c++) {
		if (*c < 0x80)
			prelen += 1;
		else if (*c < 0x800)
			prelen += 2;
		else if (*c < 0x10000)
			prelen += 3;
		else if (*c < 0x20000)
			prelen += 4;
		else if (*c < 0x40000)
			prelen += 5;
		else
			prelen += 6;
	} // for i
	return prelen;
}


char* ConvToUtf8(const WCHAR_T *W)
{
	size_t prelen = GetUtf8Size(W);
	char *utf8 = new char[prelen + 1];

	char *res = utf8;
	const WCHAR_T *c = W;

	for (; *c; c++) {

		if (*c < 0x80) {

			*res = static_cast<const char>(*c);
			++res;

		} else if (*c < 0x800) {

			*res = ((*c >> 6) & 0x1F) | 0xC0;
			++res;

			*res = (*c & 0x3F) | 0x80;
			++res;

		} /* 0x10000 0x20000 0x40000*/
		/* TODO: [E8::Core::UTF-8] сделать обработку кодовых позиций более 2 байт*/
		else {

		}
	} // for c

	*res = 0;

	return utf8;
}
