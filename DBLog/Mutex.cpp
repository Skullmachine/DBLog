
// Mutex.cpp : Implementation of CMutex
#include "stdafx.h"
#include "Mutex.h"
#include <stdio.h>
#include <comdef.h>


/////////////////////////////////////////////////////////////////////////////
// CMutex
#define LEAVING_EXECUTION -1

static DWORD WaitInMessageLoop(HANDLE mutex, IDispatch *sequenceContextDisp, DWORD eventCheckingInterval, BOOL *leavingExecution)
	{
	MSG							msg;
	DWORD						result;
	SequenceContextPtr			seqContextPtr(sequenceContextDisp);
	ExecutionPtr				execution;
	PropertyObjectPtr			termMonitorData;

	execution = seqContextPtr->GetExecution();

	termMonitorData = execution->InitTerminationMonitor();

	*leavingExecution = FALSE;

	while (true)
		{
		result = ::MsgWaitForMultipleObjects(1, &mutex, false, eventCheckingInterval, QS_ALLINPUT);

		switch (result)
			{
			case WAIT_OBJECT_0:
			case WAIT_ABANDONED_0:
			case WAIT_FAILED:
				return result;
				break;
			case WAIT_TIMEOUT:
			case WAIT_OBJECT_0 + 1: // @@@ PATCH 1.0.1 if patch made. window msg occurred
			default: // @@@ PATCH 1.0.1 was incorrectly at top
				break;
			}

			// must peek so GetMessage doesn't block while the event has been signalled
		while (::PeekMessage(&msg, 0, 0, 0, PM_NOREMOVE))
			{
			::GetMessage(&msg, NULL, 0, 0);
			::TranslateMessage(&msg);
			::DispatchMessage(&msg);
			}

			// Check termination state via the sequence context
		if (execution->GetTerminationMonitorStatus(termMonitorData, seqContextPtr.GetInterfacePtr()))
			{	
				*leavingExecution = TRUE;
				return WAIT_ABANDONED_0; 
			}
		}
	}

//CMutex::Initialize(BSTR Name, VARIANT_BOOL MakeProcessUnique, IDispatch *sequenceContextDisp, BSTR *errMsgOut, long *errorCode)
void CMutex::Initialize (_bstr_t Name, bool MakeProcessUnique, IDispatch &sequenceContextDisp, _bstr_t &errMsg, long *errorCode)
{
	_bstr_t nameStr = Name;

	BOOL leavingExecution = FALSE;

	if (errorCode)
		errorCode = 0;

	if (MakeProcessUnique == true) 
	{
		char processId[64];
		sprintf(processId, "%x", (int) GetCurrentProcessId());
		nameStr += _bstr_t(processId);
	}

	mMutexHandle = CreateMutex(NULL, FALSE, nameStr);
	if(!mMutexHandle)
	{
		if (errorCode)
			*errorCode = TS_Err_UnableToAllocateSystemResource; // -17006

		errMsg = _bstr_t("Failed to create mutex object.");
	}
	else if(WAIT_OBJECT_0 != WaitInMessageLoop(mMutexHandle, &sequenceContextDisp, 50, &leavingExecution))
	{
		if (leavingExecution == TRUE) 
		{
			if (errorCode)
				*errorCode = TS_Err_OperationCanceled; // -17604
			errMsg = _bstr_t("Execution is terminating or aborting.");
		}
		else 
		{
			if (errorCode)
				*errorCode = TS_Err_UnableToAllocateSystemResource; // -17006
			errMsg = _bstr_t("Failed to acquire mutex object.");
		}
	}
}


CMutex::CMutex():mMutexHandle(NULL)
{
}

CMutex::~CMutex()
{
	if(mMutexHandle) {
		ReleaseMutex(mMutexHandle);
		CloseHandle(mMutexHandle);
	}
}

