// TSDBLog.cpp : Implementation of CTSDBLog
#include "stdafx.h"
#include "DBLog.h"
#include "TSDBLog.h"
#include "Mutex.h"


#define EXECUTION_DATALINK_PROPERTY		"TSDatabaseLoggingDatalink"
#define DATALINK_ATTACH_NAME			"Datalink"
#define SESSION_MGR_CATEGORY			"CUSTOM::"


/////////////////////////////////////////////////////////////////////////////
// CTSDBLog

STDMETHODIMP CTSDBLog::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ITSDBLog
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

static void MakeValidName(LPWSTR name)
{
	int		index, size;
	WCHAR  *character = &name[0];

	if (!_ismbcalpha(*character) && (*character != '_')) 
		*character = (unsigned short) '_';

	size = wcslen(name);
	for (index=1; index<size; index++)
	{
		character = &name[index]; 
		if (!_ismbcalnum(*character) && (*character != '_'))
		*character = (unsigned short) '_';
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Post UUT Logging: Log UUT and step results
STDMETHODIMP CTSDBLog::LogResults(IDispatch *seqContextDisp,
                               IDispatch *optionsDisp,
                               IDispatch *mainSeqResultsDisp)
{
	TS::SequenceContextPtr pContext(seqContextDisp);

	PRE_API_FUNC_CODE

	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	HRESULT result = S_OK;


	// Get Database Schema Name to create the unique Datalink property name
	TS::PropertyObjectPtr pOptions(optionsDisp);
	_bstr_t propName(EXECUTION_DATALINK_PROPERTY);
	_bstr_t schemaName = pOptions->GetValString("DatabaseSchema.Name", 0);
	if (!(!schemaName))
		propName += schemaName;

	// FindAndReplace '.' with '_' and any other odd characters
	MakeValidName(propName);

	// Obtain a pointer to the Datalink object
	result = GetDatalink (pContext, pOptions, propName);
	if (S_OK != result)
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailureToGetDataLink, result, "");

	try
	{
		// Setup Result Filter Expression Logging for Sequence Generator
		TS::PropertyObjectPtr pExpr = pContext->AsPropertyObject()->Evaluate("FindAndReplace(Parameters.DatabaseOptions.ResultFilterExpression, \"%Result\", \"Logging.Result\", 0, True)");

		// Process results
		result = m_pDatalink->LogResults(seqContextDisp, mainSeqResultsDisp);
		if (result != S_OK)
			RAISE_ERROR_WITH_DESC(TSDBLogErr_FailureLoggingResults, result, "");
	}
	catch (...)
	{
		//release the custom session associated with the execution
		//this will release the Datalink reference as a logging error has occured using the Datalink
		try
		{
			pContext->GetExecution()->AsPropertyObject()->SetValIDispatch(propName, 0, NULL);
		}
		catch (...) {} //ignore if propName does not exist exception
		
		throw;
	}

	POST_API_FUNC_CODE(IID_ITSDBLog, pContext->Engine)
}


//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Connect to the database, and if sharing handles, look for an existing handle
HRESULT CTSDBLog::GetDatalink(TS::SequenceContextPtr pContext, TS::PropertyObjectPtr& pOptions, _bstr_t &propName)
{
    HRESULT result = S_OK;
	long error = 0;
	TS::PropertyObjectPtr pEPO;
	bool bReInitDatalink = false;
	bool shareHandles = (pOptions->GetValBoolean("ShareHandles", 0) != 0);
	double dChangeCount = pOptions->GetValNumber("ChangeCount",0);
	double dLastChangeCount;
	_bstr_t strLastChange = "LastChangeCount";

	#ifdef _DEBUG
		DebugPrintf("GetDatalink:  Name=%s, ThreadId=%d\n", (char*)propName, GetCurrentThreadId());
	#endif

	// Synchronize function across processes
    _bstr_t errMsg;
	CMutex mutex;
    mutex.Initialize((_bstr_t)propName, true, pContext, errMsg);
	
	// Get execution
	pEPO = pContext->GetExecution()->AsPropertyObject();

	// We need to track if the dbOptions have changed since the execution was started, attach counter value to execution
	if (!pEPO->Exists(strLastChange,0)) 
	{
		pEPO->NewSubProperty(strLastChange, TS::PropValType_Number, FALSE, "", 0);
		pEPO->SetValNumber(strLastChange, 0, dChangeCount);
	}

	dLastChangeCount = pEPO->GetValNumber(strLastChange, 0);

	//check if dbOptions have changed during this execution
	if ( dLastChangeCount != dChangeCount )
	{
		bReInitDatalink = true;
		m_pDatalink = 0;  //release reference
	}
	
	pEPO->SetValNumber(strLastChange, 0, dChangeCount);

	// if we are sharing datalink handle between executions, we will use the session manager
	if (shareHandles == true) 
	{
		SM::IInstrSessionMgrPtr	pSMgr;
		SM::IInstrSessionPtr pSession;
		IUnknownPtr pUnknown;
		_bstr_t sessionName(SESSION_MGR_CATEGORY);
		sessionName += propName;

		// Open session manager
		result = pSMgr.CreateInstance(__uuidof(SM::InstrSessionMgr));
		if (result) 
			RAISE_ERROR_WITH_DESC(TSDBLogErr_SessionManagerError, result, "");

		//#ifdef _DEBUG
		//pSMgr->PutAudibleLifeTimeAlerts(true);
		//#endif

		// Get session if it exists, if not, create one
		pSession = pSMgr->GetInstrSession(sessionName, false);

		if ( bReInitDatalink )
		{
			//lose the reference to the datalink
			try {
				pSession->DetachData(DATALINK_ATTACH_NAME);
			}
			catch (...) { } // Ignore exceptions
		}

		// See if attached data link object exists
		pUnknown = 0;
		try {
			pUnknown = pSession->TSRenamed_GetObject(DATALINK_ATTACH_NAME);
			#ifdef _DEBUG
				DebugPrintf("GetDataLink:  Acquired, SessionName=%s, ThreadId=%d\n", DATALINK_ATTACH_NAME, GetCurrentThreadId());
			#endif
		}
		catch (...) 
		{
			#ifdef _DEBUG
				DebugPrintf("GetDataLink:  Failed, SessionName=%s, ThreadId=%d\n", DATALINK_ATTACH_NAME, GetCurrentThreadId());
			#endif
			pUnknown = 0;
		} //exception occurs if object does not exist, we will ignore this


		try {
			//Get existing data link object if it exists
			if (pUnknown != 0) {
				m_pDatalink = pUnknown;
			}
			else 
			{
				result = m_pDatalink.CreateInstance(CLSID_TSDatalink);
				if (result != S_OK)
					RAISE_ERROR_WITH_DESC(TSDBLogErr_FailureToGetDataLink, result, "");

				// Initialize else if there was a problem creating the necessary object, bail
				result = m_pDatalink->Initialize(pContext.GetInterfacePtr(), pOptions);
				if (result != S_OK)
					RAISE_ERROR_WITH_DESC(TSDBLogErr_FailureToGetDataLink, result, "");

				//First detach to make sure an empty item does not already exist
				try {
					pSession->DetachData(DATALINK_ATTACH_NAME);
				}
				catch (...) { } // Ignore exceptions

				//Attach object
				pSession->AttachObject(DATALINK_ATTACH_NAME, m_pDatalink, (_bstr_t)"", true);
			}

			//update the user configurable booleans
			m_pDatalink->SetRuntimeBooleans(pContext.GetInterfacePtr(), pOptions);
			
			//Attach session to execution so the data link is automatically released when execution goes out of scope
			IDispatchPtr pI = pSession;
			pEPO->SetValIDispatch(propName,TS::PropOption_InsertIfMissing,pI);

			RELEASE_OBJECT_IF_EXISTS(pUnknown)
		}
		catch (...) {
			RELEASE_OBJECT_IF_EXISTS(pUnknown)
			throw;
		}
	}
	else // No sharing, so attach datalink to execution
	{
		if ( bReInitDatalink )
		{
			try 
			{
				pEPO->SetValIDispatch(propName, 0, NULL);  //detach datalink from execution
			}
			catch (...) {};  //catch propName not found
		}

		// Create property object if it does not exist
		if (!pEPO->Exists(propName,0)) 
		{
			pEPO->NewSubProperty(propName, TS::PropValType_Reference, FALSE, "", 0);
		}

		// Get the Dispatch pointer for the Datalink
		IDispatchPtr pI = pEPO->GetValIDispatch(propName,0);
		if (NULL == pI) 
		{
			result = m_pDatalink.CreateInstance(CLSID_TSDatalink);

			// Init data link
			if (S_OK == result) {
				result = m_pDatalink->Initialize(pContext, pOptions);
			}
			else {
				return result;
			}

			//update the user configurable booleans
			m_pDatalink->SetRuntimeBooleans(pContext.GetInterfacePtr(), pOptions);

			pI = m_pDatalink;
			pEPO->SetValIDispatch(propName, 0, pI);
		}
		else {
		   m_pDatalink = pI;
		}
	}

    return result;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// On-The-Fly Logging: Process model signaled that a new UUT is starting
STDMETHODIMP CTSDBLog::NewUUT(IDispatch *seqContextDisp, IDispatch *uutDisp, IDispatch *optionsDisp, 
							  IDispatch *startTimeDisp, IDispatch *startDateDisp, IDispatch *stationInfoDisp)
{

	TS::SequenceContextPtr pContext(seqContextDisp);
	PRE_API_FUNC_CODE
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
	HRESULT result = S_OK;


	if (m_pDatalink == 0 )
	{
		// Get Database Schema Name to create the unique Datalink property name
		TS::PropertyObjectPtr pOptions(optionsDisp);
		_bstr_t propName(EXECUTION_DATALINK_PROPERTY);
		_bstr_t schemaName = pOptions->GetValString("DatabaseSchema.Name", 0);
		if (!(!schemaName))
			propName += schemaName;

		// FindAndReplace '.' with '_' and any other odd characters
		MakeValidName(propName);

		// Obtain a pointer to the Datalink object
		result = GetDatalink (pContext, pOptions, propName);
		if (S_OK != result)
			RAISE_ERROR_WITH_DESC(TSDBLogErr_FailureToGetDataLink, result, "");
	}

	m_pDatalink->NewUUT(seqContextDisp, uutDisp, optionsDisp, startTimeDisp, startDateDisp, stationInfoDisp);
	
	
	POST_API_FUNC_CODE(IID_ITSDBLog, pContext->Engine)
}


//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// On-The-Fly Logging: Log a UUT result
STDMETHODIMP CTSDBLog::LogUUTResult(IDispatch *seqContextDisp, IDispatch *uutResultDisp)
{
	TS::SequenceContextPtr pContext(seqContextDisp);
	PRE_API_FUNC_CODE
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	if (m_pDatalink == 0 )
		RAISE_ERROR_WITH_DESC(TSDBLogErr_LoggingSequenceError, 0, "");
	
	retval = m_pDatalink->LogUUTResult(seqContextDisp, uutResultDisp);

	POST_API_FUNC_CODE(IID_ITSDBLog, pContext->Engine)
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// On-The-Fly Logging: Log a new step result
STDMETHODIMP CTSDBLog::LogOneResult(IDispatch *stepContextDisp, IDispatch *stepResultDisp)
{
	TS::SequenceContextPtr pContext(stepContextDisp);
	PRE_API_FUNC_CODE
	AFX_MANAGE_STATE(AfxGetStaticModuleState())

	if (m_pDatalink == 0 )
		RAISE_ERROR_WITH_DESC(TSDBLogErr_LoggingSequenceError, 0, "");

	retval = m_pDatalink->LogOneResult(stepContextDisp, stepResultDisp);

	POST_API_FUNC_CODE(IID_ITSDBLog, pContext->Engine)

}
