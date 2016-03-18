

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 7.00.0555 */
/* at Fri Mar 18 10:02:34 2016
 */
/* Compiler settings for DBLog.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 7.00.0555 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__


#ifndef __DBLog_h__
#define __DBLog_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __ITSDBLog_FWD_DEFINED__
#define __ITSDBLog_FWD_DEFINED__
typedef interface ITSDBLog ITSDBLog;
#endif 	/* __ITSDBLog_FWD_DEFINED__ */


#ifndef __ITSDatalink_FWD_DEFINED__
#define __ITSDatalink_FWD_DEFINED__
typedef interface ITSDatalink ITSDatalink;
#endif 	/* __ITSDatalink_FWD_DEFINED__ */


#ifndef __TSDBLog_FWD_DEFINED__
#define __TSDBLog_FWD_DEFINED__

#ifdef __cplusplus
typedef class TSDBLog TSDBLog;
#else
typedef struct TSDBLog TSDBLog;
#endif /* __cplusplus */

#endif 	/* __TSDBLog_FWD_DEFINED__ */


#ifndef __TSDatalink_FWD_DEFINED__
#define __TSDatalink_FWD_DEFINED__

#ifdef __cplusplus
typedef class TSDatalink TSDatalink;
#else
typedef struct TSDatalink TSDatalink;
#endif /* __cplusplus */

#endif 	/* __TSDatalink_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 



#ifndef __DBLOGLib_LIBRARY_DEFINED__
#define __DBLOGLib_LIBRARY_DEFINED__

/* library DBLOGLib */
/* [helpstring][version][uuid] */ 


EXTERN_C const IID LIBID_DBLOGLib;

#ifndef __ITSDBLog_INTERFACE_DEFINED__
#define __ITSDBLog_INTERFACE_DEFINED__

/* interface ITSDBLog */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ITSDBLog;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9D6338CE-57FB-11D4-8105-0090276F59E1")
    ITSDBLog : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LogResults( 
            /* [in] */ IDispatch *seqContextDisp,
            /* [in] */ IDispatch *optionsDisp,
            /* [in] */ IDispatch *mainSeqResultsDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LogUUTResult( 
            IDispatch *seqContextDisp,
            IDispatch *uutResultDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LogOneResult( 
            IDispatch *stepContextDisp,
            IDispatch *stepResultDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NewUUT( 
            IDispatch *seqContextDisp,
            IDispatch *uutDisp,
            IDispatch *optionsDisp,
            IDispatch *startTimeDisp,
            IDispatch *startDateDisp,
            IDispatch *stationInfoDisp) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ITSDBLogVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITSDBLog * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITSDBLog * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITSDBLog * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITSDBLog * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITSDBLog * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITSDBLog * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITSDBLog * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LogResults )( 
            ITSDBLog * This,
            /* [in] */ IDispatch *seqContextDisp,
            /* [in] */ IDispatch *optionsDisp,
            /* [in] */ IDispatch *mainSeqResultsDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LogUUTResult )( 
            ITSDBLog * This,
            IDispatch *seqContextDisp,
            IDispatch *uutResultDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LogOneResult )( 
            ITSDBLog * This,
            IDispatch *stepContextDisp,
            IDispatch *stepResultDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *NewUUT )( 
            ITSDBLog * This,
            IDispatch *seqContextDisp,
            IDispatch *uutDisp,
            IDispatch *optionsDisp,
            IDispatch *startTimeDisp,
            IDispatch *startDateDisp,
            IDispatch *stationInfoDisp);
        
        END_INTERFACE
    } ITSDBLogVtbl;

    interface ITSDBLog
    {
        CONST_VTBL struct ITSDBLogVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITSDBLog_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITSDBLog_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITSDBLog_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITSDBLog_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITSDBLog_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITSDBLog_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITSDBLog_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITSDBLog_LogResults(This,seqContextDisp,optionsDisp,mainSeqResultsDisp)	\
    ( (This)->lpVtbl -> LogResults(This,seqContextDisp,optionsDisp,mainSeqResultsDisp) ) 

#define ITSDBLog_LogUUTResult(This,seqContextDisp,uutResultDisp)	\
    ( (This)->lpVtbl -> LogUUTResult(This,seqContextDisp,uutResultDisp) ) 

#define ITSDBLog_LogOneResult(This,stepContextDisp,stepResultDisp)	\
    ( (This)->lpVtbl -> LogOneResult(This,stepContextDisp,stepResultDisp) ) 

#define ITSDBLog_NewUUT(This,seqContextDisp,uutDisp,optionsDisp,startTimeDisp,startDateDisp,stationInfoDisp)	\
    ( (This)->lpVtbl -> NewUUT(This,seqContextDisp,uutDisp,optionsDisp,startTimeDisp,startDateDisp,stationInfoDisp) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITSDBLog_INTERFACE_DEFINED__ */


#ifndef __ITSDatalink_INTERFACE_DEFINED__
#define __ITSDatalink_INTERFACE_DEFINED__

/* interface ITSDatalink */
/* [unique][helpstring][dual][uuid][object] */ 


EXTERN_C const IID IID_ITSDatalink;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9D6338D0-57FB-11D4-8105-0090276F59E1")
    ITSDatalink : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LogResults( 
            /* [in] */ IDispatch *seqContextDisp,
            /* [in] */ IDispatch *mainSeqResultsDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE Initialize( 
            /* [in] */ IDispatch *seqContextDisp,
            /* [in] */ IDispatch *optionsDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE SetRuntimeBooleans( 
            IDispatch *seqContextDisp,
            IDispatch *optionsDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE NewUUT( 
            IDispatch *seqContextDisp,
            IDispatch *uutDisp,
            IDispatch *optionsDisp,
            IDispatch *startTimeDisp,
            IDispatch *startDateDisp,
            IDispatch *stationInfoDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LogOneResult( 
            IDispatch *stepContextDisp,
            IDispatch *stepResultDisp) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE LogUUTResult( 
            IDispatch *seqContextDisp,
            IDispatch *uutResultDisp) = 0;
        
    };
    
#else 	/* C style interface */

    typedef struct ITSDatalinkVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ITSDatalink * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            __RPC__deref_out  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ITSDatalink * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ITSDatalink * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ITSDatalink * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ITSDatalink * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ITSDatalink * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ITSDatalink * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LogResults )( 
            ITSDatalink * This,
            /* [in] */ IDispatch *seqContextDisp,
            /* [in] */ IDispatch *mainSeqResultsDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *Initialize )( 
            ITSDatalink * This,
            /* [in] */ IDispatch *seqContextDisp,
            /* [in] */ IDispatch *optionsDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *SetRuntimeBooleans )( 
            ITSDatalink * This,
            IDispatch *seqContextDisp,
            IDispatch *optionsDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *NewUUT )( 
            ITSDatalink * This,
            IDispatch *seqContextDisp,
            IDispatch *uutDisp,
            IDispatch *optionsDisp,
            IDispatch *startTimeDisp,
            IDispatch *startDateDisp,
            IDispatch *stationInfoDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LogOneResult )( 
            ITSDatalink * This,
            IDispatch *stepContextDisp,
            IDispatch *stepResultDisp);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *LogUUTResult )( 
            ITSDatalink * This,
            IDispatch *seqContextDisp,
            IDispatch *uutResultDisp);
        
        END_INTERFACE
    } ITSDatalinkVtbl;

    interface ITSDatalink
    {
        CONST_VTBL struct ITSDatalinkVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ITSDatalink_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ITSDatalink_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ITSDatalink_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ITSDatalink_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ITSDatalink_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ITSDatalink_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ITSDatalink_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define ITSDatalink_LogResults(This,seqContextDisp,mainSeqResultsDisp)	\
    ( (This)->lpVtbl -> LogResults(This,seqContextDisp,mainSeqResultsDisp) ) 

#define ITSDatalink_Initialize(This,seqContextDisp,optionsDisp)	\
    ( (This)->lpVtbl -> Initialize(This,seqContextDisp,optionsDisp) ) 

#define ITSDatalink_SetRuntimeBooleans(This,seqContextDisp,optionsDisp)	\
    ( (This)->lpVtbl -> SetRuntimeBooleans(This,seqContextDisp,optionsDisp) ) 

#define ITSDatalink_NewUUT(This,seqContextDisp,uutDisp,optionsDisp,startTimeDisp,startDateDisp,stationInfoDisp)	\
    ( (This)->lpVtbl -> NewUUT(This,seqContextDisp,uutDisp,optionsDisp,startTimeDisp,startDateDisp,stationInfoDisp) ) 

#define ITSDatalink_LogOneResult(This,stepContextDisp,stepResultDisp)	\
    ( (This)->lpVtbl -> LogOneResult(This,stepContextDisp,stepResultDisp) ) 

#define ITSDatalink_LogUUTResult(This,seqContextDisp,uutResultDisp)	\
    ( (This)->lpVtbl -> LogUUTResult(This,seqContextDisp,uutResultDisp) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ITSDatalink_INTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_TSDBLog;

#ifdef __cplusplus

class DECLSPEC_UUID("9D6338CF-57FB-11D4-8105-0090276F59E1")
TSDBLog;
#endif

EXTERN_C const CLSID CLSID_TSDatalink;

#ifdef __cplusplus

class DECLSPEC_UUID("9D6338D1-57FB-11D4-8105-0090276F59E1")
TSDatalink;
#endif
#endif /* __DBLOGLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


