#pragma once
// PropInfo.h

#ifndef __PROPINFO_H_
#define __PROPINFO_H_

#include "enums.h"
#include <assert.h>

#define TSDBLog_SRC_STR						"TSDBLog"
#define TSDBLog_Mutex						"TSDBLog_Transaction_Mutex"

#define PROP_TS								L"TS"
#define PROP_TS_STEPNAME					PROP_TS L"." L"StepName"
#define PROP_TS_STEPTYPE					PROP_TS L"." L"StepType"
#define PROP_TS_SEQUENCECALL_STATUS			PROP_TS L"." L"SequenceCall.Status"
#define PROP_ASYNCID						L"AsyncId"		/* This property exists for both the Caller and the Wait step */

#define PROP_DATABASE_OPTIONS				L"DatabaseOptions"
#define PROP_SHOW_LOGGING_STATUS			L"ShowLoggingStatus"
#define PROP_RESULT_FILTER					L"ResultFilterExpression"
#define PROP_INCLUDE_ADDITIONAL_RESULTS		L"IncludeAdditionalResults"
#define PROP_INCLUDE_LIMITS					L"IncludeLimits"
#define PROP_INCLUDE_MEASUREMENTS			L"IncludeOutputValues"

#define PROP_ADDITIONAL_RESULT				L"AdditionalResults"
#define PROP_PARAMETERS						L"[Parameters]"
#define PROP_SEQUENCE_CALL					L"SequenceCall"
#define PROP_POST_ACTION					L"PostAction"
#define PROP_RESULT_LIST					L"ResultList"
#define PROP_ERROR							L"Error"
#define PROP_STATUS							L"Status"
#define PROP_COMMON							L"Common"
#define PROP_REPORTTEXT						L"ReportText"

#define SET_ALIAS_OPTIONS				TS::PropOption_InsertIfMissing | TS::PropOption_NotOwning | TS::PropOption_ReferToAlias

#define RELEASE_OBJECTPTR_IF_EXISTS(pObject) {	\
		if (pObject != 0) {						\
			try {								\
				pObject->Release();				\
				pObject = 0;					\
			}									\
			catch (...) {assert(0);}			\
		}										\
	}

#define RELEASE_OBJECT_IF_EXISTS(pObject) {		\
		if (pObject != 0) {						\
			try {								\
				pObject.Release();				\
			}									\
			catch (...) {assert(0);}			\
		}										\
	}


#define GetEngineFromContext(seqContextPtr) seqContextPtr->AsPropertyObject()->GetPropertyObject("Engine",0)


//--------------------------------------------------------------------------------------
//TSDBLogException
//--------------------------------------------------------------------------------------
class TSDBLogException
{
public:
	TSDBLogException(int TSDBLogErrorCode, HRESULT errorCode, const char *descriptionStr)
	{
		mTSDBLogErrorCode = TSDBLogErrorCode;
		mErrorCode = errorCode;
		mDescription = descriptionStr;
		mHandled = false;
	}

	int mTSDBLogErrorCode;
	HRESULT mErrorCode;
	_bstr_t mDescription;
	bool mHandled; //used to handle exceptions thrown in a recursive rtn

};


#ifdef _DEBUG
	#define DO_DEBUG_PRINTF 1
#else
	#define DO_DEBUG_PRINTF 0
#endif

#define RAISE_ERROR(TSDBLogErrorCode)	{ \
		if (DO_DEBUG_PRINTF) DebugPrintf("TSDBLog: Exception thown at %d in %s, TSDBLogError=%d\n", __LINE__, __FILE__, TSDBLogErrorCode);	\
		throw TSDBLogException((int)TSDBLogErrorCode, 0, "");                                                                    \
	}

#define RAISE_ERROR_WITH_DESC(TSDBLogErrorCode, errorCode, description)	{ \
		if (DO_DEBUG_PRINTF) DebugPrintf("TSDBLog: Exception thown at %d in %s, TSDBLogError=%d, errorCode=%x, description=%s\n", __LINE__, __FILE__, TSDBLogErrorCode, errorCode, (char*)description);	\
		throw TSDBLogException((int)TSDBLogErrorCode, (HRESULT)errorCode, description); \
	}

#define CATCH_ALL_AND_SET_ERROR_INFO(myIID, IEnginePtr) \
		catch (TSDBLogException &e) {											\
			if (DO_DEBUG_PRINTF) DebugPrintf("DBLOG:  CATCH_ALL_AND_SET_ERROR_INFO: {TSDBLogException}[%d]\n", GetCurrentThreadId()); \
			retval = ::SetErrorInfoForException(myIID, (e.mErrorCode)?(e.mErrorCode):(DISP_E_EXCEPTION), PrefixTSDBLogErrorString(IEnginePtr, e.mDescription, e.mTSDBLogErrorCode), TSDBLog_SRC_STR);	\
		}																	\
		catch (_com_error &com_except) {									\
			if (DO_DEBUG_PRINTF) DebugPrintf("DBLOG:  CATCH_ALL_AND_SET_ERROR_INFO: {_com_error &com_except}[%d]\n", GetCurrentThreadId()); \
			retval = ::SetErrorInfoForException(myIID, com_except.Error(), com_except.Description(), TSDBLog_SRC_STR);	\
		}																	\
		catch(HRESULT errorCode) {											\
			if (DO_DEBUG_PRINTF) DebugPrintf("DBLOG:  CATCH_ALL_AND_SET_ERROR_INFO: {HRESULT errorCode}[%d]\n", GetCurrentThreadId()); \
			retval = ::SetErrorInfoForException(myIID, errorCode, "", TSDBLog_SRC_STR);	\
		}																	\
		catch(int errorCode) {												\
			if (DO_DEBUG_PRINTF) DebugPrintf("DBLOG:  CATCH_ALL_AND_SET_ERROR_INFO: {int errorCode}[%d]\n", GetCurrentThreadId()); \
			retval = ::SetErrorInfoForException(myIID, (HRESULT)errorCode, "", TSDBLog_SRC_STR);	\
		}																	\
		catch (...)															\
		{																	\
			if (DO_DEBUG_PRINTF) DebugPrintf("DBLOG:  CATCH_ALL_AND_SET_ERROR_INFO: {...}[%d]\n", GetCurrentThreadId()); \
			retval = TS::TS_Err_UnexpectedSystemError;						\
		}

#define PRE_API_FUNC_CODE		HRESULT retval = S_OK;					\
								try {
#define POST_API_FUNC_CODE(myIID, IEnginePtr) }										\
								CATCH_ALL_AND_SET_ERROR_INFO(myIID, IEnginePtr)		\
								return retval;
#define POST_API_FUNC_RETHROW   }										\
								throw;


class LoggingInfo 
{
public:
	enum LoggingInfoTypes {
		LoggingInfoType_Context = 0,
		LoggingInfoType_Template = 1,
		LoggingInfoType_Copy = 2
	};
	LoggingInfo(TS::SequenceContextPtr& pContext, bool isOnTheFly);
	LoggingInfo(TS::SequenceContextPtr& pContext, LoggingInfoTypes type, bool autoDelete, bool &loggingCreated, bool isOnTheFly);
	~LoggingInfo();

	void Clear();
	void Initialize(TS::SequenceContextPtr& pContext, LoggingInfoTypes type, bool autoDelete, bool &loggingCreated, bool isOnTheFly);
	void ClearPropertyInfo();

	TS::SequenceContextPtr& GetContext(void);

	void SetUUT(TS::PropertyObjectPtr& propPtr);
	void SetDBOptions(TS::PropertyObjectPtr& propPtr);
	bool GetDBOptionShowProgress(void);
	bool GetDBOptionResultFilter(void);
	bool GetDBOptionIncludeAdditionalResults(void);
	bool GetDBOptionIncludeMeasurements(void);
	bool GetDBOptionIncludeLimits(void);

	void SetStartTime(TS::PropertyObjectPtr& propPtr);
	void SetStartDate(TS::PropertyObjectPtr& propPtr);
	void SetStationInfo(TS::PropertyObjectPtr& propPtr);

	int	GetExecutionOrder(void);
	void SetExecutionOrder(int value);
	void IncExecutionOrder(void);
	void CreatingStubStepResult(bool value);
	void LoggingNewStep(void);

	int	GetRootExecutionOrder(void);
	void SetRootExecutionOrder(int value);

	void SetUUTResult(TS::PropertyObjectPtr& propPtr);
	
	TS::PropertyObjectPtr&	GetStepResult(void);
	void SetStepResult(TS::PropertyObjectPtr propPtr);
	
	TS::PropertyObjectPtr& GetResultProperty(void);
	void SetResultProperty(TS::PropertyObjectPtr dataPtr);
	TS::PropertyObjectPtr&	GetPropData(void);

	_bstr_t	GetPropPath(void);
	void SetPropPath(_bstr_t value);

	void SetPropCategory(PropertyCategory value);

	int	GetLoggingDepth(void)			{return m_depth;}
	void SetLoggingDepth(int value)		{m_depth = value;}
	int IncLoggingDepth(void)			{return ++m_depth;}
	int DecLoggingDepth(void)			{return --m_depth;}

	#ifdef _DEBUG
	CString GetDebugName(void)		{
										CString temp;
										_bstr_t stepName;
										if (m_pLogging->Exists(L"StepResult", 0)) {
											TS::PropertyObjectPtr pStepResult = m_pLogging->GetPropertyObject(L"StepResult", 0);
											if (pStepResult != NULL && pStepResult->Exists(PROP_TS_STEPNAME, 0)) 
												stepName = pStepResult->GetValString(PROP_TS_STEPNAME,0);
										}

										if (m_path.length())
											temp.Format("StepResult='%s', Prop='%s'", (char *)stepName, (char *)m_path);
										else
											temp.Format("StepResult='%s'", (char *)stepName);
										return temp;
									}
	#endif

	void SetPropertyOrder(int value);
	void IncPropertyOrder(void);
	void ResetPropertyOrder(void);

	void InitTerminationMonitor(void);
	bool GetTerminationMonitorStatus(void); // Asks TestStand to determine if termination has occurred
	bool GetLastTerminationStatus(void);	// Returns status from last call to GetTerminationMonitorStatus

private:
	_bstr_t						m_path;
	int							m_depth;
	TS::SequenceContextPtr		m_pContext;
	TS::SequenceContextPtr		m_pRootContext;
	TS::PropertyObjectPtr		m_pLogging;
	TS::PropertyObjectPtr		m_pUUT;
	TS::PropertyObjectPtr		m_pDBOptions;
	TS::PropertyObjectPtr		m_pStartTime;
	TS::PropertyObjectPtr		m_pStartDate;
	TS::PropertyObjectPtr		m_pStationInfo;
	TS::PropertyObjectPtr		m_pExecutionOrder;
	TS::PropertyObjectPtr		m_pRootExecutionOrder;
	TS::PropertyObjectPtr		m_pUUTResult;
	TS::PropertyObjectPtr		m_pStepResult;
	TS::PropertyObjectPtr		m_pResultProperty;
	TS::PropertyObjectPtr		m_pPath;
	TS::PropertyObjectTypePtr	m_pType;
	TS::PropertyObjectPtr		m_pCategory;
	PropertyCategory			m_category;

	TS::PropertyObjectPtr		m_pResultFilter;

	TS::PropertyObjectPtr		m_pShowProgress;
	bool						m_showProgress;
	TS::PropertyObjectPtr		m_pIncludeAdditionalResults;
	bool						m_includeAdditionalResults;
	TS::PropertyObjectPtr		m_pIncludeLimits;
	bool						m_includeLimits;
	TS::PropertyObjectPtr		m_pIncludeMeasurements;
	bool						m_includeMeasurements;

	TS::PropertyObjectPtr		m_pPropertyOrder;
	int							m_propertyOrder;

	TS::PropertyObjectPtr		m_pTermination;
	bool						m_termination;

	bool						m_autoDelete;
	LoggingInfoTypes			m_type;

	bool						m_isOnTheFly;
	bool						m_creatingStubStepResult;
	bool						m_stepExecutionOrderChanged;

};

class PropertyInfo 
{
public:
	PropertyInfo(TS::PropertyObjectPtr &pStepResultProperty)
														{
															m_depth = 0; 
															m_category = tsDBOptionsPropertyCategory_Builtin; 
															m_property = pStepResultProperty;
															m_skipThisResult = true;
															m_path = L"";
														}
	PropertyInfo(PropertyInfo &parentPropertyInfo, TS::PropertyObjectPtr &pSubProperty, _bstr_t altPropertyName = "")
														{
															m_property = pSubProperty;
															m_depth = parentPropertyInfo.m_depth + 1;
															m_path.Attach(parentPropertyInfo.m_path.copy(true));
															m_category = parentPropertyInfo.m_category;
															_bstr_t name = (altPropertyName.length() == 0) ? pSubProperty->Name : altPropertyName;
															m_skipThisResult = false;

															if (m_depth == 0) {
																m_category = tsDBOptionsPropertyCategory_Builtin;
																m_skipThisResult = true;
															}
															else if (m_depth == 1) {
																if (!_wcsicmp(name, PROP_ERROR) || 
																	!_wcsicmp(name, PROP_STATUS) ||
																	!_wcsicmp(name, PROP_REPORTTEXT) ||
																	!_wcsicmp(name, PROP_COMMON) ||
																	!_wcsicmp(name, PROP_TS))
																{
																	m_category = tsDBOptionsPropertyCategory_Builtin;
																} 
																else if (!_wcsicmp(name, PROP_ADDITIONAL_RESULT))
																{
																	m_category = tsDBOptionsPropertyCategory_Additional;
																	m_skipThisResult = true;
																}
																else
																	m_category = tsDBOptionsPropertyCategory_Custom;
															} else {
																if (m_depth == 2 && 
																	m_category == tsDBOptionsPropertyCategory_Additional &&
																	!_wcsicmp(name, PROP_PARAMETERS)) 
																{
																	m_category = tsDBOptionsPropertyCategory_Parameter;
																	m_skipThisResult = true;
																}
															}
															if (!m_skipThisResult) {
																if (m_path.length() > 0)
																	m_path += ".";
																m_path += name;
															}
														}
	bool SkipThisResult(void)							{return m_skipThisResult;}
	TS::PropertyObjectPtr&	GetProperty(void)			{return m_property;}
	int	GetDepth(void)									{return m_depth;}
	void UpdateLoggingInfo(LoggingInfo &loggingInfo)	{
															loggingInfo.SetResultProperty(m_property);
															loggingInfo.SetPropPath(m_path);
															loggingInfo.SetPropCategory(m_category);
															loggingInfo.IncPropertyOrder();
														}
	bool CanProcess(StatementTraverseOptions traverseOptions)
														{
															if ((traverseOptions & tsDBOptionsTraverseOptions_SkipSubResult)!=0)
																return false;
															if (m_category == tsDBOptionsPropertyCategory_Builtin && 
																((traverseOptions & tsDBOptionsTraverseOptions_SkipBuiltinProperty)!=0))
																return false;
															if (m_category == tsDBOptionsPropertyCategory_Additional && (
																(traverseOptions & tsDBOptionsTraverseOptions_SkipAdditionalProperty)!=0))
																return false;
															if (m_category != tsDBOptionsPropertyCategory_Builtin && 
																m_category != tsDBOptionsPropertyCategory_Additional && 
																m_category != tsDBOptionsPropertyCategory_Parameter && /* CAR 141297 */
																((traverseOptions & tsDBOptionsTraverseOptions_SkipCustomProperty)!=0))
																return false;
															return true;
														}

private:
	TS::PropertyObjectPtr	m_property;
	int						m_depth;
	_bstr_t					m_path;
	PropertyCategory		m_category;
	bool					m_skipThisResult;
};

class ADOLoggingObject
{
public:
	virtual HRESULT AppendChunk(_bstr_t &name, variant_t &value) PURE;
	virtual HRESULT AppendChunk(_bstr_t &name, bstr_t &value) PURE;
	virtual void Assign(_bstr_t &name, variant_t &value) PURE;
	virtual void Assign(_bstr_t &name, bstr_t &value) PURE;
	virtual bool IsRecordSet(void) PURE;
	virtual bool IsCommand(void) PURE;
	virtual int GetStatus(_bstr_t &name) { return 0;}
};

class ADOLoggingObject_Command : public ADOLoggingObject
{
public:
	ADOLoggingObject_Command(Command25Ptr pCmd)
	{
		m_pCmd = pCmd;
	}

	virtual HRESULT AppendChunk(_bstr_t &name, variant_t &value)
	{
		return m_pCmd->Parameters->Item[name]->AppendChunk(value);
	}

	virtual HRESULT AppendChunk(_bstr_t &name, bstr_t &value)
	{
		return m_pCmd->Parameters->Item[name]->AppendChunk(value);
	}

	virtual void Assign(_bstr_t &name, variant_t &value)
	{
		m_pCmd->Parameters->Item[name]->Value = value;
	}

	virtual void Assign(_bstr_t &name, bstr_t &value)
	{
		m_pCmd->Parameters->Item[name]->Value = value;
	}

	virtual bool IsRecordSet(void) {return false;}
	virtual bool IsCommand(void) {return true;}

private:
	Command25Ptr m_pCmd;
};

class ADOLoggingObject_Recordset : public ADOLoggingObject
{
public:
	ADOLoggingObject_Recordset(_RecordsetPtr pRS)
	{
		m_pRS = pRS;
	}

	virtual HRESULT AppendChunk(_bstr_t &name, variant_t &value)
	{
		return m_pRS->Fields->Item[name]->AppendChunk(value);
	}

	virtual HRESULT AppendChunk(_bstr_t &name, bstr_t &value)
	{
		return m_pRS->Fields->Item[name]->AppendChunk(value);
	}

	virtual void Assign(_bstr_t &name, variant_t &value)
	{
		m_pRS->Fields->Item[name]->Value = value;
	}

	virtual void Assign(_bstr_t &name, bstr_t &value)
	{
		m_pRS->Fields->Item[name]->Value = value;
	}

	virtual bool IsRecordSet(void) {return true;}
	virtual bool IsCommand(void) {return false;}

	virtual int GetStatus(_bstr_t &name) 
	{ 
		return m_pRS->Fields->Item[name]->Status;
	}

private:
	_RecordsetPtr m_pRS;
};

//--------------------------------------------------------------------------------------
//CCritSectLock
//--------------------------------------------------------------------------------------
class CCritSectLock
{
public:
	// Constructor
	CCritSectLock(CRITICAL_SECTION *cs, BOOL aquireLockImmediately = true)
	{
		m_cs = cs;

		if (aquireLockImmediately) 
		{
			::EnterCriticalSection(m_cs);
			m_haveLock = true;
		}
	};

	// Destructor
	~CCritSectLock()
	{
		if (m_haveLock) 
		{
			::LeaveCriticalSection(m_cs);
			m_haveLock = false;
		}
		m_cs = 0;
	};

	// Operations
	void Lock()
	{
		::EnterCriticalSection(m_cs);
		m_haveLock = true;
		return;
	};

	void Unlock()
	{
		::LeaveCriticalSection(m_cs);
		m_haveLock = false;
		return;
	};
private:
	CRITICAL_SECTION *m_cs;
	bool m_haveLock;
};

//--------------------------------------------------------------------------------------
//Misc Utility functions
//--------------------------------------------------------------------------------------
int __cdecl DebugPrintf(const char *buf, ...);

HRESULT SetErrorInfoForException(REFIID riidSource, HRESULT errorCode, const char *descriptionStr, const char *sourceStr);
_bstr_t GetTSDBLogResString(TS::IEnginePtr enginePtr, const char *symbol);
_bstr_t PrefixTSDBLogErrorString(TS::IEnginePtr enginePtr, _bstr_t errorString, int mTSDBLogErrorCode);
_bstr_t GetTextForAdoErrors(_ConnectionPtr spAdoCon, _bstr_t raisedText);


#endif //__PROPINFO_H_