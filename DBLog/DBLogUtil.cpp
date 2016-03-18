
//PropInfo.cpp
#include "stdafx.h"
#include "DBLogUtil.h"

#define LOGGING							L"Logging"					// Name of property under context that contains the logging context

#define PROP_UUT						L"UUT"
#define PROP_START_TIME					L"StartTime"
#define PROP_START_DATE					L"StartDate"
#define PROP_STATION_INFO				L"StationInfo"
#define PROP_EXECUTION_ORDER			L"ExecutionOrder"
#define PROP_UUT_RESULT					L"UUTResult"
#define PROP_STEP_RESULT				L"StepResult"
#define PROP_STEP_RESULT_PROPERTY		L"StepResultProperty"

#define PROP_PROPERTY_RESULT			L"PropertyResult"
#define PROP_PROPERTY_DETAILS			L"PropertyResultDetails"
#define PROP_NAME						L"Name"
#define PROP_PATH						L"Path"
#define PROP_NUMERIC_FORMAT				L"NumericFormat"
#define PROP_CATEGORY					L"Category"
#define PROP_PROPERTY_ORDER				L"Order"

#define PROP_TYPE						L"Type"
#define PROP_TYPE_NAME					L"TypeName"
#define PROP_VALUE_TYPE					L"ValueType"
#define PROP_DISPLAY_STRING				L"DisplayString"
#define PROP_ELEMENT_TYPE				L"ElementType"
#define PROP_LOWER_BOUNDS				L"LowerBounds"
#define PROP_UPPER_BOUNDS				L"UpperBounds"

#define LOGGING_TEMPLATE				L"Locals.TSDBLoggingTemplate"			// OnTheFly Logging: Name of property under root context that contains the template logging property which is cloned to the root context of each thread
#define LOCALS_LOGGING					L"Locals.TSDBLogging"					// OnTheFly Logging: Name of property in under Locals of first context of a thread which holds the context.  Each context in the thread creates alias object to this.

extern CRITICAL_SECTION sLoggingOrderCritSect;

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
LoggingInfo::LoggingInfo(TS::SequenceContextPtr& pContext, bool isOnTheFly)
{
	bool loggingCreated;
	Initialize(pContext, LoggingInfoType_Context, false, loggingCreated, isOnTheFly);
}

LoggingInfo::LoggingInfo(TS::SequenceContextPtr& pContext, LoggingInfoTypes type, bool autoDelete, bool &loggingCreated, bool isOnTheFly)
{
	Initialize(pContext, type, autoDelete, loggingCreated, isOnTheFly);
}

void LoggingInfo::Clear()
{
	m_pContext = NULL;
	m_pRootContext = NULL;
	m_pLogging = NULL;
	m_pUUT = NULL;
	m_pDBOptions = NULL;
	m_pStartTime = NULL;
	m_pStartDate = NULL;
	m_pStationInfo = NULL;
	m_pExecutionOrder = NULL;
	m_pRootExecutionOrder = NULL;
	m_pUUTResult = NULL;
	m_pStepResult = NULL;
	m_pResultProperty = NULL;
	m_pPath = NULL;
	m_path = "";
	m_pType = NULL;
	m_pCategory = NULL;
	m_category = tsDBOptionsPropertyCategory_Unknown;

	m_pResultFilter = NULL;

	m_pShowProgress = NULL;
	m_showProgress = true;
	m_pIncludeAdditionalResults = NULL;
	m_includeAdditionalResults = true;
	m_pIncludeLimits = NULL;
	m_includeLimits = true;
	m_pIncludeMeasurements = NULL;
	m_includeMeasurements = true;
	
	m_depth = 0;

	m_pPropertyOrder = NULL;
	m_propertyOrder = 0;
		
	m_pTermination = NULL;
	m_termination = false;

	m_autoDelete = false;
	m_type = LoggingInfoType_Context;

	m_isOnTheFly = false;
	m_creatingStubStepResult = false;
	m_stepExecutionOrderChanged = false;
}

void LoggingInfo::Initialize(TS::SequenceContextPtr& pContext, LoggingInfoTypes type, bool autoDelete, bool &loggingCreated, bool isOnTheFly)
{
	Clear();

	m_isOnTheFly = isOnTheFly;
	m_autoDelete = autoDelete;
	loggingCreated = false;
	m_pContext = pContext;
	m_pRootContext = pContext->Root;
	m_type = type;
	m_depth = pContext->GetCallStackDepth();

	switch (type)
	{
		case LoggingInfoType_Context:

			if (!pContext->AsPropertyObject()->Exists(LOGGING, 0)) {
				pContext->AsPropertyObject()->NewSubProperty(LOGGING, TS::PropValType_Container, VARIANT_FALSE, _bstr_t(""), SET_ALIAS_OPTIONS);
				loggingCreated = true;
			}
			m_pLogging = pContext->AsPropertyObject()->GetPropertyObject(LOGGING, 0);
			break;
		case LoggingInfoType_Template:
			if (!m_pRootContext->AsPropertyObject()->Exists(LOGGING_TEMPLATE, 0)) {
				m_pRootContext->AsPropertyObject()->NewSubProperty(LOGGING_TEMPLATE, TS::PropValType_Container, VARIANT_FALSE, _bstr_t(""), SET_ALIAS_OPTIONS);
				loggingCreated = true;
			}
			m_pLogging = m_pRootContext->AsPropertyObject()->GetPropertyObject(LOGGING_TEMPLATE, 0);
			break;
		case LoggingInfoType_Copy:
			TS::PropertyObjectPtr pLoggingToAttach;
			TS::ThreadPtr pThread = pContext->Thread;
			long notused;

			//get the first sequence context of this thread
			TS::PropertyObjectPtr threadsFirstContextPropObj = pThread->GetSequenceContext(pThread->GetCallStackSize() - 1, &notused);

			//attach a ref to our Logging object to a local variable so that it will be released when the thread 
			//goes out of scope. Then for each LogResult call, attach the ref to the seq context then release after logging
			//This way the property object will always be removed from the seq context and will not remain if the seq context
			//object is cached and reused
			if (threadsFirstContextPropObj->Exists(LOCALS_LOGGING, 0) ){
				// Use logging stored on thread
				pLoggingToAttach = threadsFirstContextPropObj->GetPropertyObject( LOCALS_LOGGING, 0);
			}
			else {
				// Clone logging template
				pLoggingToAttach = m_pRootContext->AsPropertyObject()->GetPropertyObject(LOGGING_TEMPLATE, 0)->Clone("", 0);

				// Create alias to clone logging on thread
				threadsFirstContextPropObj->SetPropertyObject( LOCALS_LOGGING, SET_ALIAS_OPTIONS, pLoggingToAttach);
			}

			// Create alias to clone logging on context
			if (!pContext->AsPropertyObject()->Exists(LOGGING, 0)) {
				pContext->AsPropertyObject()->SetPropertyObject(LOGGING, SET_ALIAS_OPTIONS, pLoggingToAttach);
				loggingCreated = true;
			}
			m_pLogging = pContext->AsPropertyObject()->GetPropertyObject(LOGGING, 0);
			break;
	}

}

LoggingInfo::~LoggingInfo()
{
	if (m_pContext != NULL && m_autoDelete)
	{
		m_pContext->AsPropertyObject()->DeleteSubProperty((m_type==LoggingInfoType_Template)?LOGGING_TEMPLATE:LOGGING, 
			TS::PropOption_DeleteIfExists | TS::PropOption_ReferToAlias);
	}
}

void LoggingInfo::ClearPropertyInfo()
{
	SetResultProperty(NULL);
	SetPropPath("");
	SetPropCategory(tsDBOptionsPropertyCategory_Unknown);
}

//------------------------------------------------------------------------------------------------------
TS::SequenceContextPtr& LoggingInfo::GetContext(void)			
{
	return m_pContext;
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::SetUUT(TS::PropertyObjectPtr& propPtr)			
{
	if(m_pUUT == NULL || m_pUUT != propPtr)
	{
		m_pUUT = propPtr;
		m_pLogging->SetPropertyObject(PROP_UUT, SET_ALIAS_OPTIONS, propPtr);
	}
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::SetDBOptions(TS::PropertyObjectPtr& propPtr)	
{
	if(m_pDBOptions == NULL || m_pDBOptions != propPtr)
	{
		m_pDBOptions = propPtr;
		m_pLogging->SetPropertyObject(PROP_DATABASE_OPTIONS, SET_ALIAS_OPTIONS, propPtr);
	}
	m_pResultFilter = NULL;
	m_pShowProgress = NULL;
	m_pIncludeAdditionalResults = NULL;
	m_pIncludeLimits = NULL;
	m_pIncludeMeasurements = NULL;
}

//------------------------------------------------------------------------------------------------------
bool LoggingInfo::GetDBOptionResultFilter(void)			
{
	bool filter = true;

	// Get expression
	if (m_pResultFilter == NULL)
	{
		if (!m_pLogging->Exists(PROP_DATABASE_OPTIONS L"." PROP_RESULT_FILTER, 0))
			return true;
		else
			m_pResultFilter = m_pLogging->GetPropertyObject(PROP_DATABASE_OPTIONS L"." PROP_RESULT_FILTER, 0);
	}
	
	// Evaluate
	try {
		filter = ((m_pContext->AsPropertyObject()->EvaluateEx(m_pResultFilter->GetValString("", 0), 0))->GetValBoolean("",0)==VARIANT_TRUE);
	}
	catch (_com_error &ce) {
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedFilterExpression, ce.Error(), ce.Description());
	}
	
	return filter;
}

//------------------------------------------------------------------------------------------------------
bool LoggingInfo::GetDBOptionShowProgress(void)			
{
	if (m_pShowProgress == NULL)
	{
		if (m_pLogging->Exists(PROP_DATABASE_OPTIONS L"." PROP_SHOW_LOGGING_STATUS, 0))
			m_showProgress = (m_pLogging->GetValBoolean(PROP_DATABASE_OPTIONS L"." PROP_SHOW_LOGGING_STATUS, 0)==VARIANT_TRUE);
	}
	else
		m_showProgress = (m_pShowProgress->GetValBoolean("", 0)==VARIANT_TRUE);

	return m_showProgress;
}

//------------------------------------------------------------------------------------------------------
bool LoggingInfo::GetDBOptionIncludeAdditionalResults(void)			
{
	if (m_pIncludeAdditionalResults == NULL)
	{
		if (m_pLogging->Exists(PROP_DATABASE_OPTIONS L"." PROP_INCLUDE_ADDITIONAL_RESULTS, 0))
			m_includeAdditionalResults = (m_pLogging->GetValBoolean(PROP_DATABASE_OPTIONS L"." PROP_INCLUDE_ADDITIONAL_RESULTS, 0)==VARIANT_TRUE);
	}
	else
		m_includeAdditionalResults = (m_pIncludeAdditionalResults->GetValBoolean("", 0)==VARIANT_TRUE);

	return m_includeAdditionalResults;
}

//------------------------------------------------------------------------------------------------------
bool LoggingInfo::GetDBOptionIncludeMeasurements(void)			
{
	if (m_pIncludeMeasurements == NULL)
	{
		if (m_pLogging->Exists(PROP_DATABASE_OPTIONS L"." PROP_INCLUDE_MEASUREMENTS, 0))
			m_includeMeasurements = (m_pLogging->GetValBoolean(PROP_DATABASE_OPTIONS L"." PROP_INCLUDE_MEASUREMENTS, 0)==VARIANT_TRUE);
	}
	else
		m_includeMeasurements = (m_pIncludeMeasurements->GetValBoolean("", 0)==VARIANT_TRUE);

	return m_includeMeasurements;
}

//------------------------------------------------------------------------------------------------------
bool LoggingInfo::GetDBOptionIncludeLimits(void)			
{
	if (m_pIncludeLimits == NULL)
	{
		if (m_pLogging->Exists(PROP_DATABASE_OPTIONS L"." PROP_INCLUDE_LIMITS, 0))
			m_includeLimits = (m_pLogging->GetValBoolean(PROP_DATABASE_OPTIONS L"." PROP_INCLUDE_LIMITS, 0)==VARIANT_TRUE);
	}
	else
		m_includeLimits = (m_pIncludeLimits->GetValBoolean("", 0)==VARIANT_TRUE);

	return m_includeLimits;
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::SetStartTime(TS::PropertyObjectPtr& propPtr)	
{
	if(m_pStartTime == NULL || m_pStartTime != propPtr)
	{
		m_pStartTime = propPtr;
		m_pLogging->SetPropertyObject(PROP_START_TIME, SET_ALIAS_OPTIONS, propPtr);
	}
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::SetStartDate(TS::PropertyObjectPtr& propPtr)	
{
	if(m_pStartDate == NULL || m_pStartDate != propPtr)
	{
		m_pStartDate = propPtr;
		m_pLogging->SetPropertyObject(PROP_START_DATE, SET_ALIAS_OPTIONS, propPtr);
	}
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::SetStationInfo(TS::PropertyObjectPtr& propPtr)	
{
	if(m_pStationInfo == NULL || m_pStationInfo != propPtr)
	{
		m_pStationInfo = propPtr;
		m_pLogging->SetPropertyObject(PROP_STATION_INFO, SET_ALIAS_OPTIONS, propPtr);
	}
}

//------------------------------------------------------------------------------------------------------
// We do not cache the execution order value
int LoggingInfo::GetExecutionOrder(void)			
{
	if (m_pExecutionOrder == NULL)
	{
		if(!m_pLogging->Exists(PROP_EXECUTION_ORDER, 0))
			return 0;
		else
			m_pExecutionOrder = m_pLogging->GetPropertyObject(PROP_EXECUTION_ORDER, 0);
	}

	return (int) m_pExecutionOrder->GetValNumber("", 0);
}

void LoggingInfo::SetExecutionOrder(int value)
{
	if (m_pExecutionOrder == NULL)
	{
		m_pLogging->SetValNumber(PROP_EXECUTION_ORDER, TS::PropOption_InsertIfMissing, (double) value);
		m_pExecutionOrder = m_pLogging->GetPropertyObject(PROP_EXECUTION_ORDER, 0);
	}
	else
		m_pExecutionOrder->SetValNumber("", 0, (double) value);

	if (!m_creatingStubStepResult)
		m_stepExecutionOrderChanged = true;
}
void LoggingInfo::LoggingNewStep(void)
{
	m_stepExecutionOrderChanged = false;
}

// We only want to increment the step execution order count once for each new step, unless we are creating stub parent step results
void LoggingInfo::IncExecutionOrder(void)			
{
	if (!m_stepExecutionOrderChanged || m_creatingStubStepResult)
	{
		CCritSectLock cs(&sLoggingOrderCritSect);
		long order;
		if (m_isOnTheFly)
			order = GetRootExecutionOrder() + 1; //get logging order from root context and set in current context
		else
			order = GetExecutionOrder() + 1;
		
		SetExecutionOrder(order);
		if (m_isOnTheFly)
			SetRootExecutionOrder(order);

		#ifdef _DEBUG
			//DebugPrintf("SetLoggingOrder:  StackDepth=%d, Order Set to %d\n", loggingInfo.GetContext()->CallStackDepth, order);
			//DebugPrintf("                  RootDepth =%d, Order Changed from %d to %d\n", loggingInfo.GetContext()->Root->CallStackDepth, order, order + 1);
		#endif
	}
}

void LoggingInfo::CreatingStubStepResult(bool value)
{
	m_creatingStubStepResult = value;
}

//------------------------------------------------------------------------------------------------------
// We do not cache the execution order value
int LoggingInfo::GetRootExecutionOrder(void)			
{
	if (m_pRootExecutionOrder == NULL)
	{
		if(!m_pRootContext->AsPropertyObject()->Exists(LOGGING_TEMPLATE L"." PROP_EXECUTION_ORDER, 0))
			return 0;
		else
			m_pRootExecutionOrder = m_pRootContext->AsPropertyObject()->GetPropertyObject(LOGGING_TEMPLATE L"." PROP_EXECUTION_ORDER, 0);
	}

	return (int) m_pRootExecutionOrder->GetValNumber("", 0);
}

void LoggingInfo::SetRootExecutionOrder(int value)
{
	if (m_pRootExecutionOrder == NULL)
	{
		m_pRootContext->AsPropertyObject()->SetValNumber(LOGGING_TEMPLATE L"." PROP_EXECUTION_ORDER, TS::PropOption_InsertIfMissing, (double) value);
		m_pRootExecutionOrder = m_pRootContext->AsPropertyObject()->GetPropertyObject(LOGGING_TEMPLATE L"." PROP_EXECUTION_ORDER, 0);
	}	
	else
		m_pRootExecutionOrder->SetValNumber("", 0, (double) value);
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::SetUUTResult(TS::PropertyObjectPtr& propPtr)	
{
	if(m_pUUTResult == NULL || m_pUUTResult != propPtr)
	{
		m_pUUTResult = propPtr;
		m_pLogging->SetPropertyObject(PROP_UUT_RESULT, SET_ALIAS_OPTIONS, propPtr);
	}
}

//------------------------------------------------------------------------------------------------------
TS::PropertyObjectPtr& LoggingInfo::GetStepResult(void)			
{
	if (m_pStepResult == NULL)
		m_pStepResult = m_pLogging->GetPropertyObject(PROP_STEP_RESULT, 0);
	return m_pStepResult;
}

void LoggingInfo::SetStepResult(TS::PropertyObjectPtr propPtr)	
{
	if(m_pStepResult == NULL || m_pStepResult != propPtr)
	{
		m_pStepResult = propPtr;
		if (propPtr == NULL)
			m_pLogging->DeleteSubProperty(PROP_STEP_RESULT, PropOption_ReferToAlias | TS::PropOption_DeleteIfExists);
		else
			m_pLogging->SetPropertyObject(PROP_STEP_RESULT, SET_ALIAS_OPTIONS, propPtr);
	}
}

//------------------------------------------------------------------------------------------------------
TS::PropertyObjectPtr& LoggingInfo::GetResultProperty(void)			
{
	if (m_pResultProperty == NULL)
		m_pResultProperty = m_pLogging->GetPropertyObject(PROP_PROPERTY_RESULT, 0);
	return m_pResultProperty;
}

void LoggingInfo::SetResultProperty(TS::PropertyObjectPtr dataPtr)	
{
	if(m_pResultProperty != NULL && m_pResultProperty == dataPtr)
		return;

	m_pResultProperty = dataPtr;
	if (dataPtr == NULL) {
		m_pLogging->DeleteSubProperty(PROP_PROPERTY_RESULT, PropOption_ReferToAlias | TS::PropOption_DeleteIfExists);
		SetPropPath("");
		SetPropCategory(tsDBOptionsPropertyCategory_Unknown);
	}
	else {
		TS::PropertyObjectTypePtr typePtr = dataPtr->Type;
		m_pLogging->SetPropertyObject(PROP_PROPERTY_RESULT, SET_ALIAS_OPTIONS, dataPtr);
		TS::PropertyObjectPtr detailsPtr = m_pLogging->GetPropertyObject(PROP_PROPERTY_DETAILS, TS::PropOption_InsertIfMissing);
		detailsPtr->SetValString(PROP_NAME, TS::PropOption_InsertIfMissing, dataPtr->Name);
		detailsPtr->SetValString(PROP_NUMERIC_FORMAT, TS::PropOption_InsertIfMissing, dataPtr->NumericFormat);
		
		m_pType = typePtr;
		TS::PropertyObjectPtr propTypePtr = detailsPtr->GetPropertyObject(PROP_TYPE, TS::PropOption_InsertIfMissing);
		propTypePtr->SetValString(PROP_TYPE_NAME, TS::PropOption_InsertIfMissing, typePtr->DisplayString);
		propTypePtr->SetValNumber(PROP_VALUE_TYPE, TS::PropOption_InsertIfMissing, typePtr->ValueType);
		propTypePtr->SetValString(PROP_DISPLAY_STRING, TS::PropOption_InsertIfMissing, typePtr->DisplayString);

		bool isArray = (typePtr->ValueType == TS::PropValType_Array);
		detailsPtr->SetValString(PROP_LOWER_BOUNDS, TS::PropOption_InsertIfMissing, (isArray) ? typePtr->ArrayDimensions->GetLowerBoundsString() : "");
		detailsPtr->SetValString(PROP_UPPER_BOUNDS, TS::PropOption_InsertIfMissing, (isArray) ? typePtr->ArrayDimensions->GetUpperBoundsString() : "");

		TS::PropertyObjectTypePtr elemPtr = typePtr->ElementType;
		TS::PropertyObjectPtr propElemPtr = propTypePtr->GetPropertyObject(PROP_ELEMENT_TYPE, TS::PropOption_InsertIfMissing);
		propElemPtr->SetValString(PROP_TYPE_NAME, TS::PropOption_InsertIfMissing, (isArray) ? elemPtr->DisplayString : "");
		propElemPtr->SetValNumber(PROP_VALUE_TYPE, TS::PropOption_InsertIfMissing, (isArray) ? elemPtr->ValueType : TS::PropValType_Container);
		propElemPtr->SetValString(PROP_DISPLAY_STRING, TS::PropOption_InsertIfMissing, (isArray) ? elemPtr->DisplayString : "");
	}
}

//------------------------------------------------------------------------------------------------------
TS::PropertyObjectPtr& LoggingInfo::GetPropData(void)			
{
	if (m_pResultProperty == NULL)
		m_pResultProperty = m_pLogging->GetPropertyObject(PROP_PROPERTY_RESULT, 0);
	return m_pResultProperty;
}

//------------------------------------------------------------------------------------------------------
_bstr_t LoggingInfo::GetPropPath(void)			
{
	if (m_pPath == NULL)
	{
		if (!m_pLogging->Exists(PROP_PROPERTY_DETAILS L"." PROP_PATH, 0))
			m_pLogging->SetValString(PROP_PROPERTY_DETAILS L"." PROP_PATH, TS::PropOption_InsertIfMissing, m_path);
		else
			m_path = m_pLogging->GetValString(PROP_PROPERTY_DETAILS L"." PROP_PATH, 0);

		m_pPath = m_pLogging->GetPropertyObject(PROP_PROPERTY_DETAILS L"." PROP_PATH, 0);
	}
	else
		m_path = m_pPath->GetValString("", 0);

	return m_path;
}

void LoggingInfo::SetPropPath(_bstr_t value)	
{
	if(m_pPath == NULL || m_path != value)
	{
		m_path.Attach(value.copy());
		if(m_pPath == NULL) 
		{
			m_pLogging->SetValString(PROP_PROPERTY_DETAILS L"." PROP_PATH, TS::PropOption_InsertIfMissing, m_path);
			m_pPath = m_pLogging->GetPropertyObject(PROP_PROPERTY_DETAILS L"." PROP_PATH, 0);
		}
		else
			m_pPath->SetValString("", 0, m_path);
	}
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::SetPropCategory(PropertyCategory value)
{
	if (m_pCategory == NULL || m_category != value)
	{
		m_category = value;
		if (m_pCategory == NULL)
		{
			m_pLogging->SetValNumber(PROP_PROPERTY_DETAILS L"." PROP_CATEGORY, TS::PropOption_InsertIfMissing, (double)(int)m_category);
			m_pCategory = m_pLogging->GetPropertyObject(PROP_PROPERTY_DETAILS L"." PROP_CATEGORY, 0);
		}
		else
			m_pCategory->SetValNumber("", 0, (double)(int)m_category);
	}
}

void LoggingInfo::SetPropertyOrder(int value)
{
	if (m_pPropertyOrder == NULL || m_propertyOrder != value)
	{
		m_propertyOrder = value;
		if (m_pPropertyOrder == NULL)
		{
			m_pLogging->SetValNumber(PROP_PROPERTY_DETAILS L"." PROP_PROPERTY_ORDER, TS::PropOption_InsertIfMissing, m_propertyOrder);
			m_pPropertyOrder = m_pLogging->GetPropertyObject(PROP_PROPERTY_DETAILS L"." PROP_PROPERTY_ORDER, 0);
		}
		else
			m_pPropertyOrder->SetValNumber("", 0, m_propertyOrder);
	}
}

void LoggingInfo::ResetPropertyOrder(void)
{
	SetPropertyOrder(0);
}

void LoggingInfo::IncPropertyOrder(void)
{
	SetPropertyOrder(m_propertyOrder + 1);
}

//------------------------------------------------------------------------------------------------------
void LoggingInfo::InitTerminationMonitor(void)
{
	if (m_pTermination == NULL) {
		m_pTermination = m_pContext->Execution->InitTerminationMonitor();
		m_termination = false;
	}
}

bool LoggingInfo::GetTerminationMonitorStatus(void)
{
	if (m_termination)
		return true;

	if (m_pTermination != NULL) {
		m_termination = (m_pContext->Execution->GetTerminationMonitorStatus(m_pTermination)==VARIANT_TRUE);
		#ifdef _DEBUG
			if (m_termination)
				DebugPrintf("Terminating:\n");
		#endif
	}

	return m_termination;
}

bool LoggingInfo::GetLastTerminationStatus(void)
{
		return m_termination;
}

//--------------------------------------------------------------------------------------
// Misc functions
//--------------------------------------------------------------------------------------
int __cdecl DebugPrintf(const char *buf, ...)
{
	int error = 0;
	char msvcbuf[8192];
	va_list argList;
	DWORD thread;
	char threadString[16];

	thread = GetCurrentThreadId();
	sprintf(threadString, "DBLog[0x%04x]", thread);

	va_start(argList, buf);
	error = vsprintf(msvcbuf, buf, argList);
	
	OutputDebugString(threadString);
	OutputDebugString(msvcbuf);

	va_end(argList);

	return error;
}

//--------------------------------------------------------------------------------------
_bstr_t GetTSDBLogResString(TS::IEnginePtr enginePtr, const char *symbol)
{
	VARIANT optionalParamNotSpecified;
	optionalParamNotSpecified.vt = VT_ERROR;
	optionalParamNotSpecified.scode = DISP_E_PARAMNOTFOUND;

	return enginePtr->GetResourceString("TSDBLOGGING", symbol, optionalParamNotSpecified, &optionalParamNotSpecified);
}

//--------------------------------------------------------------------------------------
_bstr_t PrefixTSDBLogErrorString(TS::IEnginePtr enginePtr, _bstr_t errorString, int mTSDBLogErrorCode) 
{
	if (!enginePtr) 
		return errorString;

	_bstr_t symbol;
	switch(mTSDBLogErrorCode) 
	{
		case TSDBLogErr_FailureLoggingResults:
			symbol = "Err_FailureLoggingResults";
			break;
		case TSDBLogErr_FailureToGetDataLink:
			symbol = "Err_FailureToGetDataLink";
			break;
		case TSDBLogErr_FailedInitConnection:
			symbol = "Err_FailedInitConnection";
			break;
		case TSDBLogErr_FailedToOpenStatement:
			symbol = "Err_FailedToOpenStatement";
			break;
		case TSDBLogErr_FailedToExecuteStatement:
			symbol = "Err_FailedToExecuteStatement";
			break;
		case TSDBLogErr_FailedToInitStatement:
			symbol = "Err_FailedToInitStatement";
			break;
		case TSDBLogErr_FailedToSetColumn:
			symbol = "Err_FailedToSetColumn";
			break;
		case TSDBLogErr_FailedFilterExpression:
			symbol = "Err_FailedFilterExpression";
			break;
		case TSDBLogErr_FailedPrecondition:
			symbol = "Err_FailedPrecondition";
			break;
		case TSDBLogErr_FailedNoStatementsDefined:
			symbol = "Err_FailedNoStatementsDefined";
			break;
		case TSDBLogErr_SessionManagerError:
			symbol = "Err_SessionManagerError";
			break;
		case TSDBLogErr_InvalidADOVersion:
			symbol = "Err_InvalidADOVersion";
			break;
		case TSDBLogErr_LoggingSequenceError:
			symbol = "Err_LoggingSequenceError";
			break;
		case TSDBLogErr_ContainerTypeHeterogeneous:
			symbol = "Err_ContainerTypeHeterogeneous";
			break;
		case TSDBLogErr_ReferenceTypeNotSupported:
			symbol = "Err_ReferenceTypeNotSupported";
			break;
		case TSDBLogErr_ContainerTypeNotSupported:
			symbol = "Err_ContainerTypeNotSupported";
			break;
		case TSDBLogErr_InvalidFormatValue:
			symbol = "Err_InvalidFormatValue";
			break;
		case TSDBLogErr_InvalidColumnType:
			symbol = "Err_InvalidColumnType";
			break;
		case TSDBLogErr_FailedGetPrimaryKey:
			symbol = "Err_FailedGetPrimaryKey";
			break;
		case TSDBLogErr_InvalidTraversingOption:
			symbol = "Err_InvalidTraversingOption";
			break;
		case TSDBLogErr_InvalidResultType:
			symbol = "Err_InvalidResultType";
			break;
		case TSDBLogErr_InvalidForeignKeyName:
			symbol = "Err_InvalidForeignKeyName";
			break;
		case TSDBLogErr_ErrUnknown:
		default:
			symbol = "ErrUnknown";
			break;
	}

	bstr_t buffer = GetTSDBLogResString(enginePtr, symbol);
	if (errorString.length()) {
		buffer += "\n";
		buffer += errorString;
	}
	return buffer;
}

//--------------------------------------------------------------------------------------
HRESULT SetErrorInfoForException(REFIID riidSource, HRESULT errorCode, const char *descriptionStr, const char *sourceStr)
{
	HRESULT retval = S_OK;
	ICreateErrorInfo *perrinfo;
	IErrorInfo *pIErr;
	_bstr_t description = descriptionStr;
	_bstr_t source = sourceStr;

	// If the error info has already been set by a nested function call,
	// reallySetErrorInfo will be set to false and the old error info will
	// be maintained.

	if((retval = CreateErrorInfo(&perrinfo)) == S_OK) 
	{
		perrinfo->SetGUID(riidSource);
		if(description.length() > 0)
			perrinfo->SetDescription(description);
		if(source.length() > 0)
			perrinfo->SetSource(source);
		perrinfo->SetHelpContext(0);
		perrinfo->SetHelpFile(NULL);
		if((retval = perrinfo->QueryInterface(IID_IErrorInfo, (void**)&pIErr)) == S_OK) 
		{
			SetErrorInfo(0L, pIErr);
			pIErr->Release();
		}
		perrinfo->Release();
	}

	if(retval == S_OK)
		retval = errorCode;

	return retval;
}

//--------------------------------------------------------------------------------------
_bstr_t GetTextForAdoErrors(_ConnectionPtr spAdoCon, _bstr_t raisedText)
{
	_bstr_t finalDescription;
	_bstr_t description;
	bool found = false;
	char buffer[1024];

    for (long i = 0; i < spAdoCon->Errors->Count; i++) 
	{
        HRESULT error		= spAdoCon->Errors->Item[i]->Number;
        _bstr_t errorText	= spAdoCon->Errors->Item[i]->Description;
		long nativeError	= spAdoCon->Errors->Item[i]->NativeError;
		_bstr_t stateText	= spAdoCon->Errors->Item[i]->SQLState;
		_bstr_t sourceText	= spAdoCon->Errors->Item[i]->Source;
		
		if (errorText == raisedText)
			found = true;

		description += "Description: ";
		description += errorText;

		description += "\nNumber: ";
		sprintf(buffer, "%d", error);
		description += buffer;

        description += "\nNativeError: ";
		sprintf(buffer, "%d", nativeError);
		description += buffer;

        description += "\nSQLState: ";
		description += stateText;

        description += "\nReported by: ";
		description += sourceText;
    }

	if (!found)
		finalDescription = raisedText;

	finalDescription += description;
	finalDescription += "\n";
	return finalDescription;
}