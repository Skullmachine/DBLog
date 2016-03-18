// TSDBLog.h : Declaration of the CTSDBLog

#ifndef __TSDBLOG_H_
#define __TSDBLOG_H_

#include "resource.h"       // main symbols
#include "TSDatalink.h"

/////////////////////////////////////////////////////////////////////////////
// CTSDBLog
class ATL_NO_VTABLE CTSDBLog : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CTSDBLog, &CLSID_TSDBLog>,
	public ISupportErrorInfo,
	public IDispatchImpl<ITSDBLog, &IID_ITSDBLog, &LIBID_DBLOGLib>
{
public:
	CTSDBLog()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_TSDBLOG)

DECLARE_PROTECT_FINAL_CONSTRUCT()

BEGIN_COM_MAP(CTSDBLog)
	COM_INTERFACE_ENTRY(ITSDBLog)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(ISupportErrorInfo)
END_COM_MAP()

// ISupportsErrorInfo
	STDMETHOD(InterfaceSupportsErrorInfo)(REFIID riid);

// ITSDBLog
public:
	STDMETHOD(LogOneResult)(IDispatch *stepContextDisp, IDispatch *stepResultDisp);
	STDMETHOD(LogUUTResult)(IDispatch *seqContextDisp, IDispatch *stepResultDisp);
	STDMETHOD(NewUUT)(IDispatch *seqContextDisp, IDispatch *uutDisp, IDispatch *optionsDisp, IDispatch *startTimeDisp, IDispatch *startDateDisp, IDispatch *stationInfoDisp);
	STDMETHOD(LogResults)(/*[in]*/IDispatch *seqContextDisp, /*[in]*/ IDispatch *optionsDisp, /*[in]*/ IDispatch *mainSeqResultsDisp);
protected:
	ITSDatalinkPtr m_pDatalink;
	double  m_dLastChangeCount;
	HRESULT GetDatalink(TS::SequenceContextPtr pContext, TS::PropertyObjectPtr& pOptions, _bstr_t &propName);
};

#endif //__TSDBLOG_H_
