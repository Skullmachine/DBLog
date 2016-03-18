// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently

#if !defined(AFX_STDAFX_H__9D6338C4_57FB_11D4_8105_0090276F59E1__INCLUDED_)
#define AFX_STDAFX_H__9D6338C4_57FB_11D4_8105_0090276F59E1__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#ifndef _WIN32_WINNT
#define _WIN32_WINNT 0x0501 //0x0400
#endif
#define _ATL_APARTMENT_THREADED

#define _CRT_SECURE_NO_WARNINGS

#include <afxwin.h>
#include <afxdisp.h>
#include <afxtempl.h>

#include <atlbase.h>
//You may derive a class from CComModule and use it if you want to override
//something, but do not change the name of _Module
extern CComModule _Module;
#include <atlcom.h>

// This line declares classes to access Engine API.
// Note: this file is located under c:\Teststand\api\vc\tsapivc.h
#include "tsapivc.h"
using namespace TS;

// This line declares classes to access the NI Session Manager server, typically located at C:\Program Files\National Instruments\Shared\Session Manager\NISessionServer.dll"
//#import "progid:SessionMgr.InstrSessionMgr.1" no_implementation rename_namespace("SM") rename("GetObject", "TSRenamed_GetObject")
#import "C:\Program Files (x86)\National Instruments\Shared\Session Manager\NISessionServer.dll" no_implementation rename_namespace("SM") rename("GetObject", "TSRenamed_GetObject")

// This line declares classes to access Microsoft ADO
#undef EOF
#pragma warning( disable : 4146)
// #import "C:\Program Files\Common Files\System\ado\msado15.dll" no_implementation no_namespace 
#import "C:\Program Files (x86)\Common Files\System\ado\msado15.dll" no_implementation no_namespace
#pragma warning( default : 4146)
//rename("EOF", "EndOfFile")
#define EOF -1



//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__9D6338C4_57FB_11D4_8105_0090276F59E1__INCLUDED)
