// DBLog.cpp : Implementation of DLL Exports.


// Note: Proxy/Stub Information
//      To build a separate proxy/stub DLL, 
//      run nmake -f DBLogps.mk in the project directory.

#include "stdafx.h"
#include "resource.h"
#include <initguid.h>
#include "DBLog.h"

#include "DBLog_i.c"
#include "TSDBLog.h"
#include "TSDatalink.h"

// This line declares classes to access Engine API.
#include "tsapivc.cpp"
//#include "tsapivc.cpp"

// This line declares classes to access the NI Session Manager server, typically located at C:\Program Files\National Instruments\Shared\Session Manager\NISessionServer.dll"
//#import "progid:SessionMgr.InstrSessionMgr.1" implementation_only rename_namespace("SM") rename("GetObject", "TSRenamed_GetObject")
#import "C:\Program Files (x86)\National Instruments\Shared\Session Manager\NISessionServer.dll" implementation_only rename_namespace("SM") rename("GetObject", "TSRenamed_GetObject")

// This line declares classes to access Microsoft ADO
#import "C:\Program Files (x86)\Common Files\System\ado\msado15.dll "implementation_only inject_statement("#undef EOF") no_namespace

CComModule _Module;

BEGIN_OBJECT_MAP(ObjectMap)
OBJECT_ENTRY(CLSID_TSDBLog, CTSDBLog)
OBJECT_ENTRY(CLSID_TSDatalink, CTSDatalink)
END_OBJECT_MAP()

class CDBLogApp : public CWinApp
{
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CDBLogApp)
	public:
    virtual BOOL InitInstance();
    virtual int ExitInstance();
	//}}AFX_VIRTUAL

	//{{AFX_MSG(CDBLogApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CDBLogApp, CWinApp)
	//{{AFX_MSG_MAP(CDBLogApp)
		// NOTE - the ClassWizard will add and remove mapping macros here.
		//    DO NOT EDIT what you see in these blocks of generated code!
	//}}AFX_MSG_MAP
END_MESSAGE_MAP()

CDBLogApp theApp;
CRITICAL_SECTION sLoggingOrderCritSect;

BOOL CDBLogApp::InitInstance()
{
    _Module.Init(ObjectMap, m_hInstance, &LIBID_DBLOGLib);
	InitializeCriticalSection (&sLoggingOrderCritSect);
    return CWinApp::InitInstance();
}

int CDBLogApp::ExitInstance()
{
    _Module.Term();
	DeleteCriticalSection (&sLoggingOrderCritSect);
    return CWinApp::ExitInstance();
}

/////////////////////////////////////////////////////////////////////////////
// Used to determine whether the DLL can be unloaded by OLE

STDAPI DllCanUnloadNow(void)
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState());
    return (AfxDllCanUnloadNow()==S_OK && _Module.GetLockCount()==0) ? S_OK : S_FALSE;
}

/////////////////////////////////////////////////////////////////////////////
// Returns a class factory to create an object of the requested type

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _Module.GetClassObject(rclsid, riid, ppv);
}

/////////////////////////////////////////////////////////////////////////////
// DllRegisterServer - Adds entries to the system registry

STDAPI DllRegisterServer(void)
{
    // registers object, typelib and all interfaces in typelib
    return _Module.RegisterServer(TRUE);
}

/////////////////////////////////////////////////////////////////////////////
// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
    return _Module.UnregisterServer(TRUE);
}


