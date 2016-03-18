#pragma once
// TSDatalink.h : Declaration of the CTSDatalink

#ifndef __TSDATALINK_H_
#define __TSDATALINK_H_

#include "resource.h"       // main symbols
#include "enums.h"
#include "DBLogUtil.h"
#include "Statement.h"
#include <assert.h>

#define TSDBLog_Min_ADO_Version			2.5

/////////////////////////////////////////////////////////////////////////////
// CTSDatalink
class ATL_NO_VTABLE CTSDatalink : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CTSDatalink, &CLSID_TSDatalink>,
	public ISupportErrorInfo,
	public IDispatchImpl<ITSDatalink, &IID_ITSDatalink, &LIBID_DBLOGLib>
{
public:
	CTSDatalink();
	~CTSDatalink();
	DataTypeEnum GetDateDataType(void);

DECLARE_REGISTRY_RESOURCEID(IDR_TSDATALINK)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CTSDatalink)
	COM_INTERFACE_ENTRY(ITSDatalink)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ITSDatalink
public: //methods
	STDMETHOD(Initialize)(IDispatch *seqContextDisp, IDispatch *optionsDisp);
	STDMETHOD(LogResults)(IDispatch *seqContextDisp, IDispatch *mainSeqResultsDisp);

public: // properties
	STDMETHOD(LogUUTResult)(IDispatch *seqContextDisp, IDispatch *uutResultDisp);
	STDMETHOD(LogOneResult)(IDispatch *stepContextDisp, IDispatch *stepResultDisp);
	STDMETHOD(NewUUT)(IDispatch *seqContextDisp, IDispatch *uutDisp, IDispatch *optionsDisp, IDispatch *startTimeDisp, IDispatch *startDateDisp, IDispatch *stationInfoDisp);
	STDMETHOD(SetRuntimeBooleans)(IDispatch *seqContextDisp, IDispatch *optionsDisp);
	_ConnectionPtr m_spAdoCon;
	_bstr_t      m_strDBMS;
	_bstr_t      m_name;
	VARIANT_BOOL m_bIncludeStepResults;
	VARIANT_BOOL m_bIncludeOutputValues;
	VARIANT_BOOL m_bIncludeLimits;
	VARIANT_BOOL m_bIncludeTimes;
	VARIANT_BOOL m_bUseTransactionProcessing;
	VARIANT_BOOL m_bDatalinkIsShared;
	VARIANT_BOOL m_bOnTheFlyLogging;

    // collections of types of statements
    // only run a particular type of statement over a particular type of result
	Statements m_UUTStatements;
	Statements m_stepStatements;
	Statements m_propStatements;

private:
	void CloseAllStatements();
	void CommitTransaction();
	void RollbackTransaction(void);
	void BeginTransaction(void);
	void CommitBatchStatement(Statements &stmt);
    // log info from this property if a property result
    void RecursePropertyResults(LoggingInfo &loggingInfo, PropertyInfo &propInfo, StatementTraverseOptions traverseOptions = tsDBOptionsTraverseOptions_Continue);

    // convenience routine to pop key values from a colletion of statements
	void PopStatementKeyValues(LoggingInfo &loggingInfo, Statements& stmt, bool clearAllKeys = false);

    // convenience routine to empty a collection of statements
	void clearStatements(Statements& stmts);

	bool WillStatementsLog(LoggingInfo &loggingInfo, Statements& stmts);

    // run a collection of statements over a particular result
	void ProcessStatements(LoggingInfo &loggingInfo, bool bOnTheFly, Statements& stmtsStatement, StatementTraverseOptions &traverseOptions);

	StatementTraverseOptions ResetTraversingOptions(LoggingInfo &loggingInfo, StatementTraverseOptions currentOptions);

    // run step and property result statements over a particular result
	void RecurseResultsTree(LoggingInfo &loggingInfo, int showProgress = 0);
};

    _COM_SMARTPTR_TYPEDEF(ITSDatalink, __uuidof(ITSDatalink));

#endif //__TSDATALINK_H_
