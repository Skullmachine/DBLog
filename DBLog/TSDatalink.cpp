
// TSDatalink.cpp : Implementation of CTSDatalink
#include "stdafx.h"
#include "DBLog.h"
#include "TSDatalink.h"
#include "Mutex.h"
#include <locale.h>


//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Create Datalink object
CTSDatalink::CTSDatalink()
{
	try {
        CREATEADOINSTANCE(m_spAdoCon, Connection, "Connection");
    }
	catch( _com_error &e){
		delete this;
        throw e;
	}

	m_bIncludeStepResults = 0;
	m_bIncludeOutputValues = 0;
	m_bIncludeLimits = 0;
	m_bIncludeTimes = 0;
	m_bUseTransactionProcessing = 0;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Cleanup Datalink object
CTSDatalink::~CTSDatalink()
{
	clearStatements(m_UUTStatements);
    clearStatements(m_stepStatements);
    clearStatements(m_propStatements);

	if (m_spAdoCon) 
	{	
		try {
			HRESULT hr = m_spAdoCon->Close();
		}
		catch(_com_error &){
		}
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
STDMETHODIMP CTSDatalink::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ITSDatalink
	};
	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Log a tree of results (This is not on-the-fly)
STDMETHODIMP CTSDatalink::LogResults(IDispatch *seqContextDisp, IDispatch *mainSeqResultsDisp)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
	TS::SequenceContextPtr pContext(seqContextDisp);
	TS::PropertyObjectPtr pMainResults(mainSeqResultsDisp);
	VARIANT_BOOL vBoolFalse = VARIANT_FALSE;
	LoggingInfo loggingInfo(pContext, false);
	int showProgress = 0;
	CMutex mutex;
	
	try {
		// Initialization work
		if (loggingInfo.GetDBOptionShowProgress())
			showProgress = 2;
		loggingInfo.SetUUTResult(pMainResults);
		loggingInfo.SetStepResult(pMainResults);
		
		if (showProgress)
			pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 1, "", vBoolFalse);  //show something on the progress bar
		 
		 // If true, only one execution can log at a time using the link so create mutex
		if (m_bDatalinkIsShared && m_bUseTransactionProcessing) 
		{
			_bstr_t errMsg;
			mutex.Initialize((_bstr_t)TSDBLog_Mutex, true, pContext, errMsg);
		}

		// UUT Results 
		BeginTransaction();
		try {
			StatementTraverseOptions traverseOptions = tsDBOptionsTraverseOptions_Continue;
			ProcessStatements(loggingInfo, false, m_UUTStatements, traverseOptions); // May throw exception
			CommitBatchStatement(m_UUTStatements);
			if (showProgress)
				 pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 2, "", vBoolFalse);

			if (m_bIncludeStepResults && !(traverseOptions & tsDBOptionsTraverseOptions_SkipSubResult)) 
			{
				loggingInfo.InitTerminationMonitor();
				RecurseResultsTree(loggingInfo, showProgress);
			}

			if (loggingInfo.GetLastTerminationStatus())
				RollbackTransaction();
			else
				CommitTransaction();
		}
		catch (...)	{
			try {
				RollbackTransaction();
			}
			catch (...) {}
			throw;
		}

		if (showProgress) {
			pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 100, "", vBoolFalse);
			pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 0,  "", vBoolFalse);
		}

		PopStatementKeyValues(loggingInfo, m_UUTStatements, true);
		PopStatementKeyValues(loggingInfo, m_stepStatements, true);
		PopStatementKeyValues(loggingInfo, m_propStatements, true);
	}
	catch (TSDBLogException &e) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForSchema");
		description += m_name;
		description += "\n";
		description += e.mDescription;
		e.mDescription = description;

		try {
			if (showProgress)
				pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 0, "", vBoolFalse);
		}
		catch (...){}

		throw e;
	}		
	catch (_com_error &ce) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForSchema");
		description += m_name;
		description += ".\n";
		description += ce.Description();

		try {
			if (showProgress)
				pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 0, "", vBoolFalse);
		}
		catch (...){}

		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToInitStatement, ce.Error(), description);
	}
	catch (...) {
		try {
			if (showProgress)
				pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 0,	 "", vBoolFalse);
		}
		catch (...){}

		RAISE_ERROR(TSDBLogErr_FailedToInitStatement);
	}
	
	return S_OK;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Store the logging options into the datalink object
STDMETHODIMP CTSDatalink::SetRuntimeBooleans(IDispatch* seqContextDisp, IDispatch *optionsDisp)
{
	HRESULT hr = S_OK;

	AFX_MANAGE_STATE(AfxGetStaticModuleState())

    // Get Connection String
	TS::PropertyObjectPtr pOptions (optionsDisp);
    // Get boolean options
    m_bIncludeStepResults   = pOptions->GetValBoolean("IncludeStepResults",0);
    m_bIncludeLimits        = pOptions->GetValBoolean("IncludeLimits",0);
    m_bIncludeOutputValues  = pOptions->GetValBoolean("IncludeOutputValues",0);
    m_bIncludeTimes         = pOptions->GetValBoolean("IncludeTimes",0);
    m_strDBMS               = pOptions->GetValString("DatabaseManagementSystem",0);

    m_bUseTransactionProcessing         = pOptions->GetValBoolean("UseTransactionProcessing",0);
	m_bDatalinkIsShared					= pOptions->GetValBoolean("ShareHandles",0);
	m_bOnTheFlyLogging					= pOptions->GetValBoolean("UseOnTheFlyLogging",0);

	//verify the provider supports transactions
	if ( m_bUseTransactionProcessing && m_spAdoCon->Properties->GetItem("Transaction DDL") != NULL )
		m_bUseTransactionProcessing = true;
	else 
		m_bUseTransactionProcessing = false;

	return hr;
}
  
//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Store the datalink and schema information into the datalink object
STDMETHODIMP CTSDatalink::Initialize(IDispatch* seqContextDisp, IDispatch *optionsDisp)
{
	HRESULT hr = S_OK;

	AFX_MANAGE_STATE(AfxGetStaticModuleState())
    TS::SequenceContextPtr pContext(seqContextDisp);

	// Test version of ADO
    char *oldLocale = setlocale (LC_ALL, NULL);
    setlocale (LC_ALL, "C");
	_bstr_t version = m_spAdoCon->Version;
	double value = atof((char*)version);
	bool wrongVersion = (value < TSDBLog_Min_ADO_Version);
    setlocale (LC_ALL, oldLocale);
	if (wrongVersion)
		RAISE_ERROR(TSDBLogErr_InvalidADOVersion);

    // Get Connection String
	_bstr_t connectionString;
	TS::PropertyObjectPtr pOptions (optionsDisp);

	try {
		TS::PropertyObjectPtr pCSPO = pContext->AsPropertyObject()->EvaluateEx(pOptions->GetValString("ConnectionString", 0), EvalOption_DoNotAlterValues);
		connectionString = pCSPO->GetValString("", 0);

		if (!(!connectionString)) 
		{
			int pos1, pos2;

			// "FILE NAME=TestStand Results.udl;"
			CString connection ((char*)connectionString);
			pos1 = connection.Find("FILE NAME=");
			if (pos1 >= 0) 
			{
				CString file (connection.Mid(pos1 + 10));
				pos2 = file.Find(";", 0);
				if (pos2 > 0) file.SetAt(pos2, 0);
				
				_bstr_t absPath;
				TS::IEnginePtr pEngine = pContext->Engine;
				VARIANT_BOOL cancelled;

				VARIANT_BOOL found = pEngine->FindFile((_bstr_t) file, absPath.GetAddress(), &cancelled, FindFile_PromptDisable, FindFile_AddDirToSrchList_No, 0, vtMissing);
				connectionString = connection.Left(pos1+10);
				connectionString += (found) ? absPath : _bstr_t(file);
				connectionString += ";";
				connectionString += (LPCTSTR) connection.Mid(pos1+10+pos2+1);
			}
		}
		
		if (!(!connectionString)) {
			m_spAdoCon->ConnectionString = connectionString;
		}

		m_spAdoCon->Mode = adModeReadWrite;
	    HRESULT hr = m_spAdoCon->Open("","","",-1);
		if (hr != S_OK)
			throw _com_error(hr);
	}
	catch (TSDBLogException &e) {
		throw e;
	}		
	catch (_com_error &ce) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForDataLink");
		description += connectionString;
		description += "\n";
		description += ::GetTextForAdoErrors(m_spAdoCon, ce.Description());

		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedInitConnection, ce.Error(), description);
	}
	catch (...) {
		RAISE_ERROR(TSDBLogErr_FailedInitConnection);
	}		


	// This could fail if the 1.0.x type is used.
	try {
	    m_name = pOptions->GetValString("DatabaseSchema.Name",0);
	}
	catch (_com_error &ce) {
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedNoStatementsDefined, ce.Error(), ce.Description());
	}

    if (pOptions->Exists("DatabaseSchema.Statements",0)) {
		Statements stmts;
        TS::PropertyObjectPtr pStatements = pOptions->GetPropertyObject("DatabaseSchema.Statements",0);
        long numElements = pStatements->GetNumElements();

		if (numElements < 1)
			RAISE_ERROR(TSDBLogErr_FailedNoStatementsDefined);

        for (long i=0; i<numElements; i++) {
            TS::PropertyObjectPtr pDBStatement = pStatements->GetPropertyObjectByOffset(i,0);
            Statement *pStatement = new Statement(pContext, pDBStatement, this); //May throw exception
			stmts.AddTail(pStatement);

            switch (pStatement->GetResultType()) {
				case tsDBOptionsResultType_UUT:
					m_UUTStatements.AddTail(pStatement);
					break;
				case tsDBOptionsResultType_Step:
					m_stepStatements.AddTail(pStatement);
					break;
				case tsDBOptionsResultType_Property:
					m_propStatements.AddTail(pStatement);
					break;
            }
        }

		// Update foreignKey references
		POSITION pos = stmts.GetHeadPosition();
		while (pos)
			stmts.GetNext(pos)->InitForeignKeyInformation(pContext);
    }
	return hr;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Determine whether the list of (property) statements will log a result.
bool CTSDatalink::WillStatementsLog(LoggingInfo &loggingInfo, Statements& stmts)
{
    POSITION pos = stmts.GetHeadPosition();
    while (pos) {
        if (stmts.GetNext(pos)->WillLog(loggingInfo))
			return true;
    }
	return false;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Log the result specified in the context using the list of statements
void CTSDatalink::ProcessStatements(LoggingInfo &loggingInfo, bool bOnTheFly, Statements& stmts, StatementTraverseOptions &traverseOptions)
{
    POSITION pos = stmts.GetHeadPosition();
    while (pos && !(traverseOptions & tsDBOptionsTraverseOptions_SkipResult))
    {
	    Statement *pStatement = stmts.GetNext(pos);
		if (bOnTheFly)
			pStatement->ExecuteForOnTheFly(loggingInfo, traverseOptions); // May throw exceptionLogOneResult
		else
			pStatement->Execute(loggingInfo, traverseOptions, Statement::eLogResult); // May throw exception
    }
}

StatementTraverseOptions CTSDatalink::ResetTraversingOptions(LoggingInfo &loggingInfo, StatementTraverseOptions currentOptions)
{
	StatementTraverseOptions newOptions = currentOptions & ~tsDBOptionsTraverseOptions_PropertyMask; // Turn off before going deeper
	if (!loggingInfo.GetDBOptionIncludeAdditionalResults())
		newOptions |= tsDBOptionsTraverseOptions_SkipAdditionalProperty;

	return newOptions;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Recursively process and log step results
void CTSDatalink::RecurseResultsTree(LoggingInfo &loggingInfo, int showProgress)
{
	try {
		TS::SequenceContextPtr& pContext = loggingInfo.GetContext();
		VARIANT_BOOL vBoolFalse = true;
		StatementTraverseOptions traverseOptions = tsDBOptionsTraverseOptions_Continue;
		TS::PropertyObjectPtr pStepResult = loggingInfo.GetStepResult();
		int depth = loggingInfo.IncLoggingDepth(); 
		if (showProgress > 0)
			showProgress--;

		// filter step results?
		VARIANT_BOOL logStepResult = loggingInfo.GetDBOptionResultFilter();
		if (logStepResult) 
		{
			#ifdef _DEBUG
				DebugPrintf("RecurseSteps:     Depth=%d, ThreadId=%d, %s\n", loggingInfo.GetLoggingDepth(), loggingInfo.GetContext()->Thread->Id, (char*)loggingInfo.GetDebugName().GetBuffer());
			#endif

			loggingInfo.LoggingNewStep();
			ProcessStatements(loggingInfo, false, m_stepStatements, traverseOptions); //May throw exception

			if (loggingInfo.GetTerminationMonitorStatus())
				goto Leave;
		}

		if (showProgress)
			 pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, 3, "", vBoolFalse);

		// SEQUENCE CALL - cycle through result list
		if (pStepResult->Exists(PROP_TS L"." PROP_SEQUENCE_CALL L"." PROP_RESULT_LIST, 0))
		{
			TS::PropertyObjectPtr pResultList = pStepResult->GetPropertyObject(PROP_TS L"." PROP_SEQUENCE_CALL L"." PROP_RESULT_LIST, 0);
			long numElements = pResultList->GetNumElements();
			
			int currPercent = 0;
			for (long i=0; i<numElements && !loggingInfo.GetLastTerminationStatus(); i++) 
			{
				// Increment the progress % based upon the completion of the results
				if (showProgress) {
					int percent = i * 100 / numElements;
					if (percent != currPercent && percent > 3) {
						pContext->GetThread()->PostUIMessage(UIMsg_ProgressPercent, (double)percent, "", vBoolFalse);
						currPercent = percent;
					}
				}

				// Recurse into sequence call results
				TS::PropertyObjectPtr pChildStepResult = pResultList->GetPropertyObjectByOffset(i, 0);
				if (pChildStepResult->Exists(PROP_TS, 0)) // Not sure if we still need this conditional (OLD text: Skip if not a step result - see old database.seq/Log Sequence Results)
				{
					loggingInfo.SetStepResult(pChildStepResult);
					RecurseResultsTree(loggingInfo, showProgress);  
				}
			}
		}

		// POST ACTION - cycle through result list
		if (pStepResult->Exists(PROP_TS L"." PROP_POST_ACTION L"." PROP_RESULT_LIST, 0))
		{
			TS::PropertyObjectPtr pResultList = pStepResult->GetPropertyObject(PROP_TS L"." PROP_POST_ACTION L"." PROP_RESULT_LIST, 0);
			long numElements = pResultList->GetNumElements();
			
			for (long i=0; i<numElements && !loggingInfo.GetLastTerminationStatus(); i++) 
			{
				// Recurse into post action result
				TS::PropertyObjectPtr pChildStepResult = pResultList->GetPropertyObjectByOffset(i, 0);
				if (pChildStepResult->Exists(PROP_TS, 0)) // Not sure if we still need this conditional (OLD text: Skip if not a step result - see old database.seq/Log Sequence Results)
				{
					loggingInfo.SetStepResult(pChildStepResult);
					RecurseResultsTree(loggingInfo, showProgress);  
				}
			}
		}

		if (logStepResult)
			CommitBatchStatement(m_stepStatements);

		if (loggingInfo.GetTerminationMonitorStatus())
			goto Leave;

		// Are we supposed to include values? Are there any statements to process?
		if (logStepResult && m_propStatements.GetCount() &&
			!(traverseOptions & tsDBOptionsTraverseOptions_SkipSubResult)) 
		{
			// check for step result properties to log
			PropertyInfo propInfo(pStepResult);
			RecursePropertyResults(loggingInfo, propInfo, ResetTraversingOptions(loggingInfo, traverseOptions)); // Recurse into result
			loggingInfo.ClearPropertyInfo();
			CommitBatchStatement(m_propStatements);
		}

Leave:
		// restore key values, logging depth and step result
		if (logStepResult)
			PopStatementKeyValues(loggingInfo, m_stepStatements);
		loggingInfo.DecLoggingDepth(); 
		loggingInfo.SetStepResult(pStepResult);
	}
	catch (...)	{
		throw;
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Delete a list of statements
void CTSDatalink::clearStatements(Statements& stmts)
{
    while (!stmts.IsEmpty())
    {
        Statement* pStatement = stmts.RemoveTail();
        delete pStatement;
    }
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Pops any primary key values in specified statements for the current thread and depth. This is called as we crawl back up the 
// recursive calls to log results and properties.
void CTSDatalink::PopStatementKeyValues(LoggingInfo &loggingInfo, Statements &stmts, bool clearAllKeys)
{
    POSITION pos = stmts.GetHeadPosition();
    while (pos)
        stmts.GetNext(pos)->PopKeyValue(loggingInfo, clearAllKeys);
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Recursively process and log the subproperties of a result
void CTSDatalink::RecursePropertyResults(LoggingInfo &loggingInfo, PropertyInfo &propInfo, StatementTraverseOptions traverseOptions)
{
	// Skip limit and measurement properties
	long flags = propInfo.GetProperty()->GetFlags("",0);
	if ((flags & TS::PropFlags_IsLimit && !loggingInfo.GetDBOptionIncludeLimits()) ||
		(flags & TS::PropFlags_IsMeasurementValue && !loggingInfo.GetDBOptionIncludeMeasurements()))
		return;

	// Increment the stack/property depth as we go
	loggingInfo.IncLoggingDepth(); 

	// Reset property order if we are just starting to recurse
	if (!propInfo.GetDepth())
		loggingInfo.ResetPropertyOrder();

	// Process statements
	if (!propInfo.SkipThisResult()) {
		propInfo.UpdateLoggingInfo(loggingInfo);
		#ifdef _DEBUG
			DebugPrintf("RecurseProperty:  Depth=%d, ThreadId=%d, %s\n", loggingInfo.GetLoggingDepth(), loggingInfo.GetContext()->Thread->Id, (char*)loggingInfo.GetDebugName().GetBuffer());
		#endif
		ProcessStatements(loggingInfo, false, m_propStatements, traverseOptions); // May throw exception
	}

	// check for need to go deeper in the property tree
	if (!(traverseOptions & tsDBOptionsTraverseOptions_SkipSubResult))
	{
		TS::PropertyObjectPtr pProperty = propInfo.GetProperty();
		PropertyValueTypes valueType = pProperty->Type->ValueType;
		long numElements;

		if (valueType == TS::PropValType_Array) 
		{
			// We do not recurse arrays of intrinsic elements, i.e. numeric array
			TS::PropertyValueTypes elemType;
			bstr_t lowerBounds;
			bstr_t upperBounds;
			pProperty->GetDimensions("", 0, lowerBounds.GetAddress(), upperBounds.GetAddress(), &numElements, &elemType);
			if (elemType == TS::PropValType_Container && numElements) 
			{
				for (long i=0; i<numElements; i++)
				{
					TS::PropertyObjectPtr arrayElement = pProperty->GetPropertyObjectByOffset(i, 0);
					_bstr_t elementName, actualName = arrayElement->Name;
					if (actualName.length()>0) {
						elementName = "[";
						elementName += actualName;
						elementName += "]";
					} else
						elementName = pProperty->GetArrayIndex("", 0, i);

					PropertyInfo subPropInfo(propInfo, arrayElement, elementName);
					if (subPropInfo.CanProcess(traverseOptions))
						RecursePropertyResults(loggingInfo, subPropInfo, ResetTraversingOptions(loggingInfo, traverseOptions)); // Recurse
				}
			}
		} else if (valueType == TS::PropValType_Container) {
			numElements = pProperty->GetNumSubProperties("");
			for (long i=0; i<numElements; i++) {

				// Don't mess with the TS or AdditionalParameterResults containers
				_bstr_t subPropName = pProperty->GetNthSubPropertyName("", i, 0);
				if (propInfo.GetDepth() == 0 && (_bstr_t(PROP_TS) == subPropName))
					continue;

				PropertyInfo subPropInfo(propInfo, pProperty->GetPropertyObject(subPropName, 0));
				if (subPropInfo.CanProcess(traverseOptions))
					RecursePropertyResults(loggingInfo, subPropInfo, ResetTraversingOptions(loggingInfo, traverseOptions)); // Recurse
			}
		}
	}
	PopStatementKeyValues(loggingInfo, m_propStatements);

	// Decrement the stack/property depth before we leave
	loggingInfo.DecLoggingDepth(); 
}

// Commits records for a recordsets with batch optimistic cursor
void CTSDatalink::CommitBatchStatement(Statements &stmts)
{
    POSITION pos = stmts.GetHeadPosition();
    while (pos)
	    stmts.GetNext(pos)->Commit();
}

void CTSDatalink::BeginTransaction()
{
	if (m_bUseTransactionProcessing)
		m_spAdoCon->BeginTrans();
}

/*
Normally, a provider will close the cursor on a recordset after a commit or
rollback. We need to close our statements so that a new cursor will be retrieved.
This is needed if the datalink is used to log multiple executions
*/
void CTSDatalink::RollbackTransaction()
{
	if (m_bUseTransactionProcessing)
	{
		m_spAdoCon->RollbackTrans();
		CloseAllStatements();
	}

}

void CTSDatalink::CommitTransaction()
{
	if (m_bUseTransactionProcessing)
	{
		m_spAdoCon->CommitTrans();
		CloseAllStatements();
	}
}

void CTSDatalink::CloseAllStatements()
{
	Statements *stmts;
	for (int i=0; i<3; i++)    
	{
		switch ( i)
		{
		case 0:
			stmts = &m_UUTStatements;
			break;
		case 1:
			stmts = &m_stepStatements;
			break;
		case 2:
			stmts = &m_propStatements;
			break;
		}
		
		POSITION pos = stmts->GetHeadPosition();
		while (pos)
			stmts->GetNext(pos)->Close();
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
STDMETHODIMP CTSDatalink::NewUUT(IDispatch *seqContextDisp, IDispatch *uutDisp, IDispatch *optionsDisp, 
							     IDispatch *startTimeDisp, IDispatch *startDateDisp, IDispatch *stationInfoDisp)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
	TS::SequenceContextPtr pContext(seqContextDisp);
	TS::PropertyObjectPtr pUUT(uutDisp);
	TS::PropertyObjectPtr pOptions(optionsDisp);
	TS::PropertyObjectPtr pStartTime(startTimeDisp);
	TS::PropertyObjectPtr pStartDate(startDateDisp);
	TS::PropertyObjectPtr pStationInfo(stationInfoDisp);
	TS::SequenceContextPtr pRootContext(pContext->Root);

	//create the template logging context that on-the-fly will use
	bool loggingCreated = false;
	LoggingInfo loggingInfo(pContext, LoggingInfo::LoggingInfoType_Template, false, loggingCreated, true);
	if (loggingCreated)
	{
		loggingInfo.SetUUT(pUUT);
		loggingInfo.SetDBOptions(pOptions);
		loggingInfo.SetStartTime(pStartTime);
		loggingInfo.SetStartDate(pStartDate);
		loggingInfo.SetStationInfo(pStationInfo);
	}

    // Remove any dangling key values
	PopStatementKeyValues(loggingInfo, m_UUTStatements, true); 

	#ifdef _DEBUG
		DebugPrintf("NewUUT:           Stack=%s, Depth=%d, Rppt order set to %d\n", (char*)pRootContext->CallStackName, pRootContext->CallStackDepth, 0);
	#endif

	return S_OK;
}

// Fix for when Access swaps month and day in adDBTimeStamp column for parameterized insert statements
// logging in "European" local that uses dd/mm/yyyy format
DataTypeEnum CTSDatalink::GetDateDataType(void)
{
	DataTypeEnum dataType = adDBTimeStamp;
	if (m_spAdoCon)
	{
		CString connectionString ((char*)m_spAdoCon->ConnectionString);
		if (connectionString.Find("Provider=Microsoft.Jet.OLEDB")>=0)
			dataType = adDate;
	}

	return dataType;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5
// On-The-Fly Logging: Log a new step result
STDMETHODIMP CTSDatalink::LogOneResult(IDispatch *stepContextDisp, IDispatch *stepResultDisp)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
	TS::SequenceContextPtr pContext(stepContextDisp);
	TS::PropertyObjectPtr pStepResult(stepResultDisp);
	Statements stepStatements;

	#ifdef _DEBUG
		DebugPrintf("=============================================================================================================\n");
		DebugPrintf("LogOneResult:     Step=%s, Depth=%d, ThreadId=%d\n", (char*)pStepResult->GetValString(PROP_TS_STEPNAME,0), pContext->CallStackDepth, pContext->Thread->Id);
	#endif

	// Check if we are logging the steps in the current context
	if (pContext->CallerDiscardsResults)
		return S_OK;	 

	try 
	{
		bool loggingCreated = false;
		LoggingInfo loggingInfo(pContext, LoggingInfo::LoggingInfoType_Copy, true, loggingCreated, true);

		// Add StepResult to logging container
		loggingInfo.SetStepResult(pStepResult);

		// Filter results for non-sequence call steps upfront.  Sequence call steps will be filtered later in Statement::Execute if no child result was logged.
		if (!loggingInfo.GetStepResult()->Exists(PROP_TS_SEQUENCECALL_STATUS, 0) && !loggingInfo.GetDBOptionResultFilter())
			return S_OK;

		// Build list of statements that will log this result
		POSITION pos = m_stepStatements.GetHeadPosition();
		while (pos) {
			Statement *pStatement = m_stepStatements.GetNext(pos);
			if (pStatement->WillLog(loggingInfo))
				stepStatements.AddTail(pStatement);
		}

		if (!stepStatements.IsEmpty()) {
			// Generate stub SequenceCall/Wait result for foreign key statements that this result references
			POSITION pos = stepStatements.GetHeadPosition();
			while (pos)
				stepStatements.GetNext(pos)->CreateParentStubs(loggingInfo);

			// Loop thru the statements and log results
			StatementTraverseOptions traverseOptions = tsDBOptionsTraverseOptions_Continue;
			ProcessStatements(loggingInfo, true, stepStatements, traverseOptions);
			CommitBatchStatement(stepStatements);

			// Are we supposed to include values? Are there any statements to process?
			if (0 != m_propStatements.GetCount() &&
				!(traverseOptions & tsDBOptionsTraverseOptions_SkipSubResult)) 
			{
				// check for step result properties to log
				PropertyInfo propInfo(pStepResult);
				RecursePropertyResults(loggingInfo, propInfo, ResetTraversingOptions(loggingInfo, traverseOptions)); // Recurse into result
				CommitBatchStatement(m_propStatements);
			}

			PopStatementKeyValues(loggingInfo, stepStatements);
			stepStatements.RemoveAll();
		}
 	}
	catch (_com_error &ce) {
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailureLoggingResults, ce.Error(), ce.Description());
	}
	catch (TSDBLogException &e) {
		throw e;
	}		
	
	return S_OK;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// On-The-Fly Logging: Log a UUT result
STDMETHODIMP CTSDatalink::LogUUTResult(IDispatch *seqContextDisp, IDispatch *resultDisp)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState())
	TS::SequenceContextPtr pContext(seqContextDisp);
	TS::PropertyObjectPtr pUUTResult(resultDisp);

	try 
	{
	#ifdef _DEBUG
		DebugPrintf("=============================================================================================================\n");
		DebugPrintf("LogUUTResult:     ThreadId=%d\n", pContext->Thread->Id);
	#endif
		bool loggingCreated = false;
		LoggingInfo loggingInfo(pContext, LoggingInfo::LoggingInfoType_Copy, true, loggingCreated, true);

		// Add UUTResult to logging container
		loggingInfo.SetUUTResult(pUUTResult);

		// Log the UUT
		StatementTraverseOptions traverseOptions = tsDBOptionsTraverseOptions_Continue;
		ProcessStatements(loggingInfo, true, m_UUTStatements, traverseOptions);
		
		// Cleanup keys
		PopStatementKeyValues(loggingInfo, m_UUTStatements);
		CommitBatchStatement(m_UUTStatements);
	}
	catch (_com_error &ce) {
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailureLoggingResults, ce.Error(), ce.Description());
	}
	catch (TSDBLogException &e) {
		throw e;
	}		
	
	return S_OK;
}
