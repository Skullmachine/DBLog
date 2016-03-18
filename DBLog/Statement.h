#pragma once

#ifndef __STATEMENT_H_
#define __STATEMENT_H_

#include "stdafx.h"
#include "DBLogUtil.h"
#include "Columns.h"
#include "enums.h"
#include <vector>

class CTSDatalink;
class CDbColumn;

class CKeyItem {
public:
#ifdef DEBUG
	~CKeyItem() 
	{
		keyDepth;
	}
#endif

	int				keyDepth;
	unsigned long	executionOrder;
	_variant_t		keyValue;
};

class CLastKeyList {
public:
	~CLastKeyList() 
	{
		while (!keyItems.IsEmpty())
			delete keyItems.RemoveTail();
	}
	
    DWORD							threadId;
    CList<CKeyItem*, CKeyItem*>		keyItems;
};
typedef CList<CLastKeyList*, CLastKeyList*> CLastKeyLists;


class Statement { 

public:
	enum eProcessType{
	eLogResult,
	eCreateStub,
	eUpdateStub
	};

    Statement(TS::SequenceContextPtr& pContext, TS::PropertyObjectPtr& pDBStatement, CTSDatalink *pDatalink);
    ~Statement();

    // should be called immediately after construction
	void InitForeignKeyInformation(TS::SequenceContextPtr &pContext);

	void ExecuteForOnTheFly(LoggingInfo &loggingInfo, StatementTraverseOptions &traverseOptions);
	void CreateStubResult(LoggingInfo &loggingInfo);
	void CreateParentStubs(LoggingInfo &loggingInfo);
	void CreateOnTheFlyCommands();

    // Actions needed to open a Statement object for processing
    void Open(TS::SequenceContextPtr& pContext, _ConnectionPtr pCon);
	void Commit();
	void Close();

	bool IsOnTheFlyLogging();

	bstr_t Statement::GetIntrinsicTypeName(TS::PropertyValueTypes propType);
	bool WillLog(LoggingInfo &loggingInfo);
    // main purpose of this object, processing the SQL statement
    // and recording the results into a database
    void Execute(LoggingInfo &loggingInfo, StatementTraverseOptions &traverseOptions, eProcessType eType=eLogResult);

    // convenience accessor
    CTSDatalink*    GetDatalink()					{return m_pDatalink;}
    ResultTypes     GetResultType()					{return m_resultType;}
	StatementType	GetStmtType()					{return m_type;}
    _bstr_t			GetName()						{return m_name;}
	bool			HasSameName(_bstr_t bstrName)	{return m_name == bstrName;}

    // returns the appropriate key value for foreign key references
	_variant_t		GetLastKeyValue(LoggingInfo &loggingInfo);
	CKeyItem *		GetLastKeyItem(LoggingInfo &loggingInfo); 
	_variant_t		GetKeyValueForDepth(LoggingInfo &loggingInfo, int depth);

	CLastKeyList*	GetLastKeyList(LoggingInfo &loggingInfo);
	CLastKeyList*	GetLastKeyListFromId(long currentId);

    void			PushKeyValue (LoggingInfo &loggingInfo, _variant_t keyValue);	// pushes key value on the stack as we go deeper in the results tree
    void			PopKeyValue (LoggingInfo &loggingInfo, bool clearAllKeys = false);	// pops key value from the stack as we come up the results tree

private:
    _bstr_t						m_name;					// Statement name
    StatementType				m_type;					// Record set or command
    TS::ExpressionPtr			m_textExpr;				// SQL statement
	_bstr_t						m_stubText;				//SQL statement for creating stubs
	_bstr_t						m_updateText;			//SQL statement for updating stubs
	_bstr_t						m_updateWaitText;		//SQL statement for updating Wait step stubs
	std::vector<_bstr_t>		m_typesToLog;			// types to look for before logging
	std::vector<_bstr_t>		m_expectedProperties;	// properties to look for before logging

	TS::ExpressionPtr			m_preconditionExpr;		// expression to evaluate before logging
    CursorLocationEnum			m_cursorLocation;   
    CursorTypeEnum				m_cursorType;
    LockTypeEnum				m_lockType;
    ResultTypes					m_resultType;			// UUT, step or property
	StatementTraverseOptions	m_logOptions;

    Command25Ptr				m_spAdoCommand;
    Command25Ptr				m_spStubCommand;
    Command25Ptr				m_spUpdateCommand;
    Command25Ptr				m_spUpdateWaitCommand;
    _RecordsetPtr				m_spAdoRecordSet;

    // storage for foreign key references
    // when we go deeper in the results tree, m_currentKeyValue
    // is pushed onto this stack
    CLastKeyLists				m_lastKeyThreadList;
    CTSDatalink*				m_pDatalink;        
    Columns						m_columns;
    Columns						m_primaryKeyColumns;			// columns which make up the primary key
	Columns						m_foreignKeyColumns;
	Columns						m_recursiveKeyColumns;
														// currently ASSUMING only one
	CRITICAL_SECTION			m_critSect;
	CRITICAL_SECTION			m_keyCritSect;

	bool	m_opened;
};

// definition for a collection of statements
typedef CList<Statement*, Statement*> Statements;

#endif