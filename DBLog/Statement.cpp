
#include "stdafx.h"
#include "DBLog.h"
#include "Statement.h"
#include "Columns.h"
#include "TSDatalink.h"
#include "tsapivc.h"	// Header file found in <TestStand>/api/vc/
#include "tsvcutil.h"   // Header file found in <TestStand>/api/vc/

#define GET_ENUM(et,str,mv) {tempVal = pDBStatement->GetValNumber(str, 0); \
                             tempLong = (long)tempVal; \
                             mv = (et)tempLong; }

#define MAX_RETRY 5;

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Extract the settings for the statement, setup the keylists, and populate the column information for the statement
Statement::Statement(TS::SequenceContextPtr& pContext, TS::PropertyObjectPtr& pDBStatement, CTSDatalink *pDatalink) 
{
	double tempVal;
    long tempLong;
    CursorTypes tempCursorType;
    LockTypes tempLockType;
    CursorLocations tempCursorLocation;
	_variant_t avtString;
	VARIANT    avString;
	SAFEARRAY *asfString;
	TSUTIL::SafeArray<BSTR, VT_BSTR> stringSafeArray;

	InitializeCriticalSection (&m_critSect);
	InitializeCriticalSection (&m_keyCritSect);

	// Process statement settings
    try {
		// Get statement name
		m_name = pDBStatement->GetValString("Name",0);
		m_pDatalink = pDatalink;

		// Get statement type
		GET_ENUM(StatementType,"StatementType",m_type);
		m_textExpr = pContext->Engine->NewExpression();
		m_textExpr->Text = pDBStatement->GetValString("StatementText",0);
   		m_textExpr->Tokenize(0, 0);

		// Get result type
		GET_ENUM(ResultTypes, "ResultType", m_resultType);
		
		// Get types to log list
		avtString = pDBStatement->GetValVariant("TypesToLog",0);
 		avString = avtString.Detach();
		asfString = V_ARRAY(&avString);
		stringSafeArray.Set(asfString, true);
		stringSafeArray.GetVector(m_typesToLog);

		// Get expected property list
		avtString = pDBStatement->GetValVariant("ExpectedProperties",0);
 		avString = avtString.Detach();
		asfString = V_ARRAY(&avString);
		stringSafeArray.Set(asfString, true);
		stringSafeArray.GetVector(m_expectedProperties);

		// Get precondition
		m_preconditionExpr = pContext->Engine->NewExpression();
		m_preconditionExpr->Text = pDBStatement->GetValString("Precondition",0);
		m_preconditionExpr->Tokenize(0, 0);
    
		// Get cursor location
		//adUseNone is obsolete, use server is default
		GET_ENUM(CursorLocations, "CursorLocation", tempCursorLocation);
		m_cursorLocation = (tsDBOptionsCursorLocation_Client == tempCursorLocation) ? adUseClient : adUseServer;
    
   		// Get cursor type
		GET_ENUM(CursorTypes, "CursorType", tempCursorType);
		switch (tempCursorType) {
			case tsDBOptionsCursorType_Keyset:
				m_cursorType = adOpenKeyset;
				break;

			case tsDBOptionsCursorType_Dynamic:
				m_cursorType = adOpenDynamic;
				break;

			case tsDBOptionsCursorType_Static:
				m_cursorType = adOpenStatic;
				break;
			case tsDBOptionsCursorType_ForwardOnly:
				m_cursorType = adOpenForwardOnly;
				break;
			case tsDBOptionsCursorType_Unspecified:
			default:
				m_cursorType = adOpenUnspecified;
				break;
		}

   		// Get lock type
		GET_ENUM(LockTypes, "LockType", tempLockType);
		switch (tempLockType) {
			case tsDBOptionsLockType_Pessimistic:
				m_lockType = adLockPessimistic;
				break;
			case tsDBOptionsLockType_Optimistic:
				m_lockType = adLockOptimistic;
				break;
			case tsDBOptionsLockType_BatchOptimistic:
				m_lockType = adLockBatchOptimistic;
				break;
			case tsDBOptionsLockType_ReadOnly:
				m_lockType = adLockReadOnly;
				break;
			case tsDBOptionsLockType_Unspecified:
			default:
				m_lockType = adLockUnspecified;
		}

		m_logOptions = (int) pDBStatement->GetValNumber("TraverseOptions",0);

		// Make sure we ignore child property result options for UUT and Property results
		switch (m_resultType) {
			case tsDBOptionsResultType_UUT:
				m_logOptions &= tsDBOptionsTraverseOptions_UUTMask;
				break;
			case tsDBOptionsResultType_Step:
				m_logOptions &= tsDBOptionsTraverseOptions_StepMask;
				break;
			case tsDBOptionsResultType_Property:
				m_logOptions &= tsDBOptionsTraverseOptions_PropertyMask;
				break;
			default:
				RAISE_ERROR_WITH_DESC(TSDBLogErr_InvalidTraversingOption, TS::TS_Err_ValueIsInvalidOrOutOfRange, "");
		}

		//Get Columns
		if (pDBStatement->Exists("Columns",0)) {
			TS::PropertyObjectPtr pColumns = pDBStatement->GetPropertyObject("Columns",0);
			long numElements = pColumns->GetNumElements();

			for (long i=0; i<numElements; i++) {
				TS::PropertyObjectPtr pDBColumn = pColumns->GetPropertyObjectByOffset(i,0);
				CDbColumn *pColumn = new CDbColumn(pContext->Engine, pDBColumn, this); // May throw exception
				m_columns.AddTail(pColumn); 

				// Determine key columns
				if (pColumn->IsPrimaryKey())
					m_primaryKeyColumns.AddTail(pColumn);
				else if (pColumn->IsForeignKey())
					m_foreignKeyColumns.AddTail(pColumn);
			}
		}

		m_opened = false;
	}
	catch (TSDBLogException &e) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForStatement");
		description += m_name;
		description += ".\n";
		description += e.mDescription;
		e.mDescription = description;
		throw e;
	}		
	catch (_com_error &ce) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForStatement");
		description += m_name;
		description += ".\n";
		description += ce.Description();
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToInitStatement, ce.Error(), description);
	}
	catch (...) {
		RAISE_ERROR(TSDBLogErr_FailedToInitStatement);
	}		
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Cleanup the statement's recordset and command handles, and discard the key lists
Statement::~Statement()
{
    if (m_spAdoRecordSet)
	{
		try {
			m_spAdoRecordSet->Close();
		}
		catch( _com_error &e){
			e;
		}
		m_spAdoRecordSet.Release();
	}

    if (m_spAdoCommand) 
		m_spAdoCommand.Release();

    if (m_spStubCommand) 
		m_spStubCommand.Release();

    if (m_spUpdateCommand) 
		m_spUpdateCommand.Release();

    while (!m_lastKeyThreadList.IsEmpty())
        delete m_lastKeyThreadList.RemoveTail();
	
	// Cleanup columns lists (don't delete because they're also held in m_columns)
    m_primaryKeyColumns.RemoveAll();
    m_foreignKeyColumns.RemoveAll();
    m_recursiveKeyColumns.RemoveAll();

    while (!m_columns.IsEmpty())
        delete m_columns.RemoveTail();

	DeleteCriticalSection (&m_critSect);
	DeleteCriticalSection (&m_keyCritSect);
}

void Statement::InitForeignKeyInformation(TS::SequenceContextPtr &pContext)
{
	POSITION pos = m_foreignKeyColumns.GetHeadPosition();
	while (pos) {
		CDbColumn* pColumn = m_foreignKeyColumns.GetNext(pos);
		pColumn->InitForeignKeyInformation();

		// Determine which foreign keys of step statements reference themselves
		Statement *pStmt = pColumn->GetForeignKeyStmt();
		if (pStmt) {
			if (pStmt->GetResultType() == tsDBOptionsResultType_Step && pStmt->m_name == m_name)
				m_recursiveKeyColumns.AddTail(pColumn);
		} else {
			_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForStatement");
			description += m_name;
			description += ".\n";
			description += GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForColumn");
			description += pColumn->GetName();
			description += ".\n";
			description += GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForRefStatement");
			description += pColumn->GetForeignKeyName();
			description += ".\n";
			
			RAISE_ERROR_WITH_DESC(TSDBLogErr_InvalidForeignKeyName, 0, description);
		}
	}
}

bool Statement::IsOnTheFlyLogging()
{ 
	return (m_pDatalink->m_bOnTheFlyLogging == VARIANT_TRUE);
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Attempt to create the required INSERT and UPDATE statemtents needed for command on the fly logging.
void Statement::CreateOnTheFlyCommands()
{
	char *tempString = 0, *inputString = 0, *fieldsString = 0;

	try 
	{
		char *pCStr = 0;
		int index, count, length;

		// Do simple checks to verify the format of the command
		length = m_spAdoCommand->CommandText.length() + 1;
		tempString = new char[length];
		strcpy(tempString, m_spAdoCommand->CommandText);
		_strupr(tempString);
		pCStr = strstr(tempString, "INSERT"); 
		if (!pCStr) goto Leave;
		pCStr = strstr(pCStr, "("); 
		if (!pCStr) goto Leave;
		pCStr = strstr(pCStr, ")");
		if (!pCStr) goto Leave;
		pCStr = strstr(pCStr, "VALUES");
		if (!pCStr) goto Leave;
		pCStr = strstr(pCStr, "(");
		if (!pCStr) goto Leave;
		pCStr = strstr(pCStr, "?");
		if (!pCStr) goto Leave;
		pCStr = strstr(pCStr, ")");
		if (!pCStr) goto Leave;

#ifdef _DEBUG
		static bool bUseSingleCommand = false;
		// For Testing: This will prevent the server from using INSERT/UPDATE commands
		if (bUseSingleCommand)
			goto Leave; 
#endif

		//process INSERT command
		inputString = new char[length];
		fieldsString = new char[length];

		strcpy(tempString, m_spAdoCommand->CommandText);
		pCStr = strtok(tempString, "("); //parse til after table name
		strcpy(inputString, pCStr);
		pCStr = strtok(NULL, ")"); //parse out column names
		strcpy(fieldsString, pCStr);

		//build list of column names
		CList<CString, CString&> columnList;
		pCStr = strtok(fieldsString, ",");
		do {
			if ( pCStr )
				columnList.AddTail(CString(pCStr));
			pCStr = strtok(NULL, ",");
		} while (pCStr != NULL);

		//build insert statement
		CDbColumn* pColumn, *pKeyColumn;
		CString string;
		POSITION pos, pos2;

		//walk down list to find the key columns
		strcat(inputString, " (");
		count = 0;
		pos = m_columns.GetHeadPosition();
		pos2 = columnList.GetHeadPosition();
		while (pos) {
			pColumn = m_columns.GetNext(pos);
			string = columnList.GetNext(pos2);
			if ( pColumn->IsPrimaryKey() || pColumn->IsForeignKey() )
			{
				strcat(inputString, string);
				strcat(inputString, ",");
				count++;
			}
		}

		if ( count == 0 ) //no keys defined
			goto Leave;

		inputString[strlen(inputString)-1] = '\0'; //remove last comma
		strcat(inputString, ") VALUES (");
		for (index=0; index<count; index++) {
			strcat(inputString, "?,");
		}
		inputString[strlen(inputString)-1] = '\0'; //remove last comma
		strcat(inputString, ")");

		m_stubText = inputString;  //that's the stub statement

		//build update statements
		//a seperate update cmd is needed for Wait step stubs because they need to update the foreign key values

		//parse out tablename from Insert statement
		strcpy(tempString, inputString);
		strcpy(inputString, "UPDATE ");

		
		pCStr = strtok(tempString, " "); // Skip INSERT keyword
		pCStr = strtok(NULL, " "); // Get INTO or table name

		if (!_stricmp(pCStr, "INTO")) // Skip optional INTO keyword
			pCStr = strtok(NULL, " "); // Get table name

		strcat(inputString, pCStr);  // Add table name to input string
		strcat(inputString, " SET "); // Add SET keyword

		strcpy(tempString, inputString);  //use tempString as Wait stub update cmd

		pos = m_columns.GetHeadPosition();
		pos2 = columnList.GetHeadPosition();
		while (pos) {
			char tt[1024];
			pColumn = m_columns.GetNext(pos);
			string = columnList.GetNext(pos2);
			if ( !pColumn->IsPrimaryKey() )
			{
				sprintf(tt, "%s=?,", string);
				strcat( tempString, tt);		//append all columns expect primary key to Wait step update
				if ( !pColumn->IsForeignKey() )
					strcat(inputString, tt);	//only not key columns in regular update
			}
		}

		inputString[strlen(inputString)-1] = '\0'; //remove last comma
		tempString[strlen(tempString)-1] = '\0'; //remove last comma

		strcat( inputString, " WHERE ");
		strcat( tempString, " WHERE ");

		//find primary key column
		pos = m_columns.GetHeadPosition();
		pos2 = columnList.GetHeadPosition();
		while (pos) {
			pColumn = m_columns.GetNext(pos);
			string = columnList.GetNext(pos2);
			if (pColumn->IsPrimaryKey())
			{
				char tt[1024];
				sprintf( tt, "%s=?", string); 
				strcat(inputString, tt);
				strcat(tempString, tt);
				break;
			}
		}
		pKeyColumn = pColumn;

		m_updateText = inputString;
		m_updateWaitText = tempString;

		//#ifdef _DEBUG
		//	DebugPrintf("CreateOnTheFlyCommands:  UpdateText=%s\n", (char*)m_stubText);
		//	DebugPrintf("CreateOnTheFlyCommands:  UpdateText=%s\n", (char*)m_updateText);
		//	DebugPrintf("CreateOnTheFlyCommands:  UpdateWaitText=%s\n", (char*)m_updateWaitText);
		//#endif

		//create command objects
		CREATEADOINSTANCE(m_spStubCommand, Command, "Command"); 
		CREATEADOINSTANCE(m_spUpdateCommand, Command, "Command"); 
		CREATEADOINSTANCE(m_spUpdateWaitCommand, Command, "Command"); 

		m_spStubCommand->ActiveConnection = m_spAdoCommand->ActiveConnection; 
		m_spUpdateCommand->ActiveConnection = m_spAdoCommand->ActiveConnection;
		m_spUpdateWaitCommand->ActiveConnection = m_spAdoCommand->ActiveConnection;
		m_spStubCommand->CommandText      = m_stubText;
		m_spUpdateCommand->CommandText      = m_updateText;
		m_spUpdateWaitCommand->CommandText      = m_updateWaitText;
		m_spStubCommand->CommandType  = adCmdText;
		m_spUpdateCommand->CommandType  = adCmdText;
		m_spUpdateWaitCommand->CommandType  = adCmdText;

		m_spStubCommand->put_Prepared(true);
		m_spUpdateCommand->put_Prepared(true);
		m_spUpdateWaitCommand->put_Prepared(true);

		Command25Ptr cmdPtr;
		pos = m_columns.GetHeadPosition();
		while (pos) {
			pColumn = m_columns.GetNext(pos);
			//primary and foreign keys get set with the stub cmd, all other columns by the update cmd
			if ( pColumn->IsPrimaryKey() || pColumn->IsForeignKey() )
				cmdPtr = m_spStubCommand;
			else
				cmdPtr = m_spUpdateCommand;
			
			_ParameterPtr pprmCol = cmdPtr->CreateParameter(pColumn->GetName(), pColumn->GetAdoType(), 
				pColumn->GetDirection(), pColumn->GetSize(), vtMissing);
			pprmCol->Attributes = adParamNullable;
			if (pColumn->GetAdoType() == adVarBinary || pColumn->GetAdoType() == adLongVarBinary)
				pprmCol->Attributes = adParamNullable | adParamLong;

			cmdPtr->Parameters->Append(pprmCol);

			//all columns except primary key are updated by the Wait stub cmd
			if ( !pColumn->IsPrimaryKey() )
				m_spUpdateWaitCommand->Parameters->Append(pprmCol);
		}

		//Add parameter for the WHERE clause of update stmts, we will use primary key column name
		{
			_ParameterPtr pprmCol = m_spUpdateCommand->CreateParameter(m_primaryKeyColumns.GetHead()->GetName(), pKeyColumn->GetAdoType(), 
				adParamInput, pKeyColumn->GetSize(), vtMissing);
			pprmCol->Attributes = adParamNullable;
			if (pKeyColumn->GetAdoType() == adVarBinary || pKeyColumn->GetAdoType() == adLongVarBinary)
				pprmCol->Attributes = adParamNullable | adParamLong;

			m_spUpdateCommand->Parameters->Append(pprmCol);
			m_spUpdateWaitCommand->Parameters->Append(pprmCol);
		}
	}
	catch (...)
	{
		if (m_spStubCommand) 
			m_spStubCommand.Release();

		if (m_spUpdateCommand) 
			m_spUpdateCommand.Release();

		m_spStubCommand = NULL;
		m_spUpdateCommand = NULL;

		if ( tempString)
			delete tempString;
		if (inputString)
			delete inputString;
		if (fieldsString)
			delete fieldsString;
		throw;
	}
Leave:
	if (tempString)
		delete tempString;
	if (inputString)
		delete inputString;
	if (fieldsString)
		delete fieldsString;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Open the statement's recordset or command objects
void Statement::Open(TS::SequenceContextPtr& pContext, _ConnectionPtr pCon)
{
	CDbColumn* pColumn;
    _variant_t vRows((short)0);

	CREATEADOINSTANCE(m_spAdoCommand, Command, "Command"); 
	m_spAdoCommand->ActiveConnection = pCon;
	m_spAdoCommand->CommandText      = (m_textExpr->Evaluate(pContext->AsPropertyObject(), 0))->GetValString("",0);
    m_spAdoCommand->CommandType      = adCmdText;

	// If on the fly and INSERT statement used, determine if we need to define create and update commands
	if ((m_type == tsDBOptionsStatementType_Parameterized) && IsOnTheFlyLogging())
	{
		// we must have at least one primary key so WHERE clause can 		
	    if (m_primaryKeyColumns.GetHead())
		{
			bool bCreateCommands = false;
			// UUT result should be created before step result
			if (m_resultType == tsDBOptionsResultType_UUT)
				bCreateCommands = true;
			else 
			{
				// Determine if a column references itself, i.e. parent foreign key relationship
				POSITION pos = m_columns.GetHeadPosition();
				while (pos) {
					pColumn = m_columns.GetNext(pos);
					if ( pColumn->GetForeignKeyName() == m_name ) //this is a parent reference
					{
						bCreateCommands = true;
						break;
					}
				}
			}
			// Attempt to parse INSERT command to create stub commands, if fails, we will just use original single command
			if (bCreateCommands)
				CreateOnTheFlyCommands();
		}
	}

	// Prepare SQL statement
	if ((m_type == tsDBOptionsStatementType_Parameterized) || (m_type == tsDBOptionsStatementType_Procedure))
	{
		try {
			m_spAdoCommand->put_Prepared(true);

			//walk column list and add each as a statement parameter
			//may not be columns per say, but actually parameters
			POSITION pos = m_columns.GetHeadPosition();
			while (pos) {
				pColumn = m_columns.GetNext(pos);
				_ParameterPtr pprmCol = m_spAdoCommand->CreateParameter(pColumn->GetName(), pColumn->GetAdoType(), 
					pColumn->GetDirection(), pColumn->GetSize(), vtMissing);

				pprmCol->Attributes = adParamNullable;

				if (pColumn->GetAdoType() == adVarBinary || pColumn->GetAdoType() == adLongVarBinary)
					pprmCol->Attributes = adParamNullable | adParamLong;

				m_spAdoCommand->Parameters->Append(pprmCol);
			}
		}
		catch( _com_error &ce) {
			_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForColumn");
			description += pColumn->GetName();
			description += "\n";
			description += ::GetTextForAdoErrors(pCon, ce.Description());

			RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToInitStatement, ce.Error(), description);
		}
	}

	// Open recordset
	if (m_type == tsDBOptionsStatementType_RecordSet)
	{
		CREATEADOINSTANCE(m_spAdoRecordSet, Recordset, "Recordset"); 
		m_spAdoRecordSet->PutRefActiveConnection( pCon );
		m_spAdoRecordSet->PutRefSource(m_spAdoCommand);
		m_spAdoRecordSet->CursorLocation = m_cursorLocation;
		
		HRESULT hr = m_spAdoRecordSet->Open(vtMissing, vtMissing, m_cursorType, m_lockType,-1);
		if (hr != S_OK)
			throw _com_error(hr);
	}
	m_opened = true;
}

//Determine intrinsic type name. Note that "Container" is returned for container and array
bstr_t Statement::GetIntrinsicTypeName(TS::PropertyValueTypes propType)
{
	bstr_t intrinsicName;

	switch (propType)
	{
		case PropValType_Boolean:
			intrinsicName = ksTypeName_Boolean;
			break;
		case PropValType_String:
			intrinsicName = ksTypeName_String;
			break;
		case PropValType_Number:
			intrinsicName = ksTypeName_Number;
			break;
		case PropValType_Reference:
			intrinsicName = ksTypeName_Reference;
			break;
		case PropValType_Container:
			intrinsicName = ksTypeName_Container;
			break;
		case PropValType_Array:
			intrinsicName = ksTypeName_Container;
			break;
	}
	return intrinsicName;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5		 
// Evaluate the statement's preconditions and types to determine whether to log the result
bool Statement::WillLog(LoggingInfo &loggingInfo)
{
	TS::SequenceContextPtr& pContext = loggingInfo.GetContext();
	bool bWillLog = false;
		
	try {
		// See if steptype or property type is in type list
		if (m_typesToLog.size() > 0)
		{
			bstr_t intrinsicName, intrinsicArrayName, customName, customArrayName;

			switch (m_resultType) 
			{
				case tsDBOptionsResultType_Step:
					customName = loggingInfo.GetStepResult()->GetValString(PROP_TS_STEPTYPE, 0);
					break;

				case tsDBOptionsResultType_Property:
				{
					// Intrinsic and custom name
					TS::PropertyObjectTypePtr propObjType = loggingInfo.GetPropData()->Type;
					TS::PropertyValueTypes propType = propObjType->ValueType;
					intrinsicName = GetIntrinsicTypeName(propType);
					customName = propObjType->TypeName;

					// Intrinsic array name
					if (propType == PropValType_Array) 
					{
						TS::PropertyObjectTypePtr prototypeObjType = propObjType->GetElementType();
						TS::PropertyValueTypes prototypeType = prototypeObjType->ValueType;

						intrinsicArrayName = "ArrayOf(";
						intrinsicArrayName += GetIntrinsicTypeName(prototypeType);
						intrinsicArrayName += ")";

						if (customName.length())
						{
						customArrayName = prototypeObjType->GetTypeName();
						if (customArrayName.length() > 0)
							{
							customArrayName = "ArrayOf(";
							customArrayName += customArrayName;
							customArrayName += ")";
							}
						}
					}
					break;
				}
				default:
					RAISE_ERROR_WITH_DESC(TSDBLogErr_InvalidResultType, TS::TS_Err_ValueIsInvalidOrOutOfRange, "");
					break;
			}

			// Is there a match
			bool found = false;
			for (unsigned int i=0; ((!found) && (i < m_typesToLog.size())); i++) 
			{
				if (m_typesToLog[i].length()) 
				{
					if ((customName.length() > 0) && (_wcsicmp(m_typesToLog[i], customName) == 0))
						found = true;
					if (m_resultType == tsDBOptionsResultType_Property) 
					{
						if ((intrinsicName.length() > 0) && (_wcsicmp(m_typesToLog[i], intrinsicName) == 0) &&
							((intrinsicArrayName.length() == 0) || _wcsicmp(intrinsicName, ksTypeName_Container)))
							found = true;
						else if ((intrinsicArrayName.length() > 0) && (_wcsicmp(m_typesToLog[i], intrinsicArrayName) == 0))
							found = true;
						else if ((customArrayName.length() > 0) && (_wcsicmp(m_typesToLog[i], customArrayName) == 0))
							found = true;
					}
				}
			}

			if (!found)
				goto Leave;
		}

		// See if expected properites exist
		if (m_expectedProperties.size() > 0)
		{
			TS::PropertyObjectPtr pProperty = pContext->AsPropertyObject();
			VARIANT_BOOL exists;
			for (unsigned int i=0; (i < m_expectedProperties.size()); i++)
			{
				exists = pProperty->Exists(m_expectedProperties[i], 0);
				if (!exists)
					goto Leave;
			}
		}

		// See if precondition is true
		if (m_preconditionExpr->NumTokens)
		{		
			bool precondition = true;
			try {
				if (!(m_preconditionExpr->Evaluate(pContext->AsPropertyObject(), 0))->GetValBoolean("", 0))
			        goto Leave;
			}
			catch (_com_error &ce) {
				RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedPrecondition, ce.Error(), ce.Description());
			}
		}
		bWillLog = true;

Leave:
		return bWillLog;
	}
	catch (...)
	{
		throw;
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Execute SQL statement which logs a result
void Statement::Execute(LoggingInfo &loggingInfo, StatementTraverseOptions &traverseOptions, eProcessType eType)
{
	TS::SequenceContextPtr& pContext = loggingInfo.GetContext();
	// The lock is needed so that we can reliably access the created record primary key
	CCritSectLock cs(&m_critSect);
    try {
		int maxRetry;
		bool attemptRetry;
        POSITION pos;
		CDbColumn* pColumn;
		_variant_t vKeyVal;
		Command25Ptr cmdPtr;
		bool statmentLogged = false;
		bool bIsOnTheFlyLogging = IsOnTheFlyLogging();

		bool waitIdExists = false;
		bool bSeqCallStatusExists = false;
		if (eType != eCreateStub) {
			waitIdExists = (loggingInfo.GetStepResult()->Exists(PROP_ASYNCID, 0) == VARIANT_TRUE);
			bSeqCallStatusExists = (loggingInfo.GetStepResult()->Exists(PROP_TS_SEQUENCECALL_STATUS, 0) == VARIANT_TRUE);
		}
		bool bWaitStep = waitIdExists && bSeqCallStatusExists;
		bool bAsyncCallStep = waitIdExists && !bSeqCallStatusExists;
		bool bSeqCallStep = !waitIdExists && bSeqCallStatusExists;

		TS::SequenceContextPtr& pCallerContext = pContext->Caller;
		bool bFirstInThreadsStack = (((pCallerContext==NULL) && !pContext->CallStackDepth) || 
									 ((pCallerContext!=NULL) && (pContext->Thread->Id != pCallerContext->Thread->Id)));

		// See if this statement will log a new result.  The creating and updating of stubs are spawned from new results that are already being logged.
		//   OnTheFlyLogging can only create and update stubs, so will just check results of type eLogResult.
		//   OnTheFlyLogging only processes step statements that already returned true for WillLog in CTSDatalink::LogOneResult, so exclude them
		//   OnTheFlyLogging step statements can only get here if an update to a stub fails to find its key and instead calls back with eLogResult. In this case we want to log the record anyway.
		if ((eType == eLogResult) && 
			((m_resultType != tsDBOptionsResultType_Step) || ((m_resultType == tsDBOptionsResultType_Step) && !bIsOnTheFlyLogging)) &&
			!WillLog(loggingInfo) ) 
		{
			#ifdef _DEBUG
				//DebugPrintf("!Execute:         Stmt=%s, Type=%s, %s, Wait=%s, Traverse=%d\n", (char*)m_name, (eType==eLogResult)?"eLogResult":((eType==eCreateStub)?"eCreateStub":"eUpdateStub"), (char*)loggingInfo.GetDebugName().GetBuffer(), (bWaitStep)? "True":"False", traverseOptions);
			#endif
			// We do not need to push a key onto the stack
			goto Leave; // We should update traverseOptions on exit
		}
		#ifdef _DEBUG
			DebugPrintf("Execute:          ------------------------------------------------------------------------------------------\n");
			DebugPrintf("Execute:          Stmt=%s, Type=%s, %s, Wait=%s, Traverse=%d\n", (char*)m_name, (eType==eLogResult)?"eLogResult":((eType==eCreateStub)?"eCreateStub":"eUpdateStub"), (char*)loggingInfo.GetDebugName().GetBuffer(), (bWaitStep)? "True":"False", traverseOptions);
		#endif

		// For OnTheFlyLogging Only
		// Filter out sequence call steps that do not have a stub. Non-sequence call steps were filtered in CTSDatalink::LogOneResult.
		// But if we have gotten to this point, CTSDatalink::LogOneResult will expect to pop a key off the stack, so push dummy value.
		if ((eType == eLogResult) && (m_resultType == tsDBOptionsResultType_Step) && (bSeqCallStep || bWaitStep) && bIsOnTheFlyLogging)
		{
			try 
			{
				if (!loggingInfo.GetDBOptionResultFilter()) 
				{
					_variant_t vt;
					vt.vt = VT_NULL;
					PushKeyValue(loggingInfo, vt); // Push dummy key value
					return; // We should not update traverseOptions on exit
				}
			}
			catch (_com_error &ce) 
			{
				RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedFilterExpression, ce.Error(), ce.Description());
			}
		}

		// Open the statement if it is not already, this is done here so unused recordsets are not opened
		if ( m_opened == false)
				Open(pContext, m_pDatalink->m_spAdoCon);	// May throw exception

		// For OnTheFlyLogging Only
		if (eType == eUpdateStub) 
		{
		    if (m_primaryKeyColumns.GetHead()) // Must have a primary key
			{
				//See if this is a Wait step
				if (bWaitStep)
				{
					// This is a Wait step that was waiting upon a thread or execution.  We need to update the async seq's parent stub 
					// with the wait step's result.  We must search for the key value using the thread id of the async thread
					CLastKeyList* lastKeyList;
					lastKeyList = GetLastKeyListFromId((long) loggingInfo.GetStepResult()->GetValNumber(PROP_ASYNCID, 0));
					if ( lastKeyList && !lastKeyList->keyItems.IsEmpty() )
					{
						//Pop off entry from list once we get the key
						CKeyItem *pKey = lastKeyList->keyItems.GetTail();
						vKeyVal = pKey->keyValue;
						lastKeyList->keyItems.RemoveTail();				
					}
					else
						vKeyVal.vt = VT_NULL; // No key found
				}
				else
					vKeyVal = GetLastKeyValue(loggingInfo); // Just get the last key
			}
			if (vKeyVal.vt == VT_NULL)
			{
  				Execute(loggingInfo, traverseOptions, eLogResult); //no matching key found, so create new record
				// We do not need to push a key onto the stack
  				return; // We should not update traverseOptions on exit
			}
		}

		// Adjust SQL cursor for recordset statements, and select command objects for other statements
		if ( m_type == tsDBOptionsStatementType_RecordSet )
		{
			if (eType != eUpdateStub)
			{
				// When setting and putting data to an Access database from two separate
				// threads that each use a unique database connection, an error can occur,
				// "Could not update; currently locked", stating that the Table is locked
				// by another user.  This method will retry when this occurs.
				maxRetry = MAX_RETRY;
				attemptRetry = true;
				while (attemptRetry == true)
				{
					try {
						m_spAdoRecordSet->AddNew();
						attemptRetry = false;
					}
					catch( _com_error &e)
					{
						if (maxRetry > 0)	
							maxRetry--;	
						else 
						{
							_bstr_t eDescrip = e.Description();
							throw e;
						}
					}
				}
			}
			else // (eType == eUpdateStub)
			{
				//set the cursor to the record to update if it has a primary key 
			    if (pColumn = m_primaryKeyColumns.GetHead())
				{
					_bstr_t sqlText = pColumn->GetName() + "=";
					if (vKeyVal.vt == VT_BSTR )
						sqlText += "\'" + _bstr_t(vKeyVal) + "\'";
					else
						sqlText += _bstr_t(vKeyVal);
					
					m_spAdoRecordSet->MoveFirst();
					m_spAdoRecordSet->Find(sqlText,0,adSearchForward);
					if (m_spAdoRecordSet->GetEOF() ) //record not found
					{
						Execute(loggingInfo, traverseOptions, eLogResult); //no stub found, so create new record
						// We do not need to push a key onto the stack
						return; // We should not update traverseOptions on exit
					}
				}
			}
		}
		else // tsDBOptionsStatementType_Parameterized || tsDBOptionsStatementType_Procedure
		{
			// Select which command object to use for OnTheFly logging
		    if (bIsOnTheFlyLogging && m_primaryKeyColumns.GetHead()) // Must have a primary key
			{
				switch (eType)
				{
				case eLogResult:
					cmdPtr =  m_spAdoCommand;
					break;

				case eCreateStub:
					cmdPtr = m_spStubCommand;
					break;

				case eUpdateStub:
					if (bWaitStep)
						cmdPtr = m_spUpdateWaitCommand;
					else
						cmdPtr = m_spUpdateCommand;
					break;

				}
			}
			if (cmdPtr == NULL)
				cmdPtr = m_spAdoCommand;
		}

		// For OnTheFlyLogging only, set the step execution order to the value of the stub's order
		// Otherwise increment the step execution order unless we creating a stub result for a wait step
		CKeyItem *pKey;
		if (bIsOnTheFlyLogging && (eType == eUpdateStub) && !bWaitStep && (pKey = GetLastKeyItem(loggingInfo))) //wait step logging order gets new value
			loggingInfo.SetExecutionOrder(pKey->executionOrder);
		else if (m_resultType == tsDBOptionsResultType_Step && (eType != eCreateStub || (!bWaitStep && !bFirstInThreadsStack)))
			loggingInfo.IncExecutionOrder();
		
		// For OnTheFlyLogging Only
		bool bUsesSingleCommand = true;
		if (bIsOnTheFlyLogging) 
		{
			bUsesSingleCommand =  	((m_type == tsDBOptionsStatementType_Procedure) || 
									 ((m_type == tsDBOptionsStatementType_Parameterized) && !((m_spStubCommand!=NULL) && (m_spUpdateCommand!=NULL) && (m_spUpdateWaitCommand!=NULL))));
		}

		// Loop through all the columns and set the new input column values associate with this statement
        pos = m_columns.GetHeadPosition();
        while (pos) {
            pColumn = m_columns.GetNext(pos);
			if ( pColumn->GetDirection() == adParamInput || pColumn->GetDirection() == adParamInputOutput)
			{
				if ((eType == eLogResult) || 
					(eType == eCreateStub &&  (bUsesSingleCommand) && (pColumn->IsPrimaryKey()) ) ||
					(eType == eCreateStub && !(bUsesSingleCommand) && (pColumn->IsPrimaryKey() || pColumn->IsForeignKey()) ) ||
					(eType == eUpdateStub && !(pColumn->IsPrimaryKey() || pColumn->IsForeignKey()) ) ||
					(eType == eUpdateStub && bUsesSingleCommand ) ||
					(eType == eUpdateStub && bWaitStep && pColumn->IsForeignKey() ) // If this is a wait step, we need to fill in the foreign key value because the async thread did not have a parent when we created the stub
				   )
				{
					if ( m_type == tsDBOptionsStatementType_RecordSet ) 
					{
						// For a recordset we just need to set the column value
						pColumn->SetColumn(loggingInfo, ADOLoggingObject_Recordset(m_spAdoRecordSet)); // May throw exception
					}
					else { 
						// tsDBOptionsStatementType_Parameterized || tsDBOptionsStatementType_Procedure
						if (bUsesSingleCommand && (eType == eUpdateStub)) 
						{
							// For OnTheFlyLogging Only
							if (pColumn->IsPrimaryKey())
								CmdITEM(cmdPtr, pColumn->GetName()) = vKeyVal; // Reset primary key back to value from earlier create 
							else if (pColumn->IsForeignKey() && (HasSameName(pColumn->GetForeignKeyName()))) // Is step parent foreign key
								pColumn->SetColumn(loggingInfo, ADOLoggingObject_Command(cmdPtr), pContext->CallStackDepth); // Get the previous level foreign key // May throw exception
							else
								pColumn->SetColumn(loggingInfo, ADOLoggingObject_Command(cmdPtr)); // May throw exception
						}
						else
							pColumn->SetColumn(loggingInfo, ADOLoggingObject_Command(cmdPtr)); // May throw exception
					}
				}
			}
        }

		// For OnTheFlyLogging Only
		// Need to include the primary key parameter value for the WHERE clause
		if ((eType == eUpdateStub) && (m_type == tsDBOptionsStatementType_Parameterized)  && !bUsesSingleCommand)
			CmdITEM(cmdPtr, m_primaryKeyColumns.GetHead()->GetName()) = vKeyVal;

		// When setting and putting data to an Access database from two separate
		// threads that each use a unique database connection, an error can occur,
		// "Could not update; currently locked", stating that the Table is locked
		// by another user.  This method will retry when this occurs.
		maxRetry = MAX_RETRY;
		attemptRetry = true;
		while (attemptRetry == true)
		{
			try {
				switch ( m_type )
				{
					case tsDBOptionsStatementType_RecordSet:
						#ifdef _DEBUG
							DebugPrintf("Execute:          Update %s\n", (char*)m_name);
						#endif
						m_spAdoRecordSet->Update(vtMissing, vtMissing);
						break;

					case tsDBOptionsStatementType_Procedure:
					{
						_variant_t vColVal;
						if (eType == eLogResult || eType == eUpdateStub) 
						{
							#ifdef _DEBUG
								DebugPrintf("Execute:          %s\n", (char*)cmdPtr->CommandText);
							#endif
							cmdPtr->Execute(NULL, NULL, adCmdStoredProc);
							
							//update output column expression
							pos = m_columns.GetHeadPosition();
							while (pos) {
								pColumn = m_columns.GetNext(pos);
								if ( pColumn->GetDirection() != adParamInput )
								{
									vColVal =  CmdITEM(cmdPtr,pColumn->GetName());
									pColumn->UpdateOutputValue(loggingInfo, vColVal );
								}
							}
						}
						break;
					}
					case tsDBOptionsStatementType_Parameterized:
						if (!bUsesSingleCommand || (bUsesSingleCommand && ((eType == eLogResult) || (eType == eUpdateStub))) ) 
						{
							#ifdef _DEBUG
								bstr_t text(cmdPtr->CommandText);
								DebugPrintf("Execute:          cmdPtr->%s...\n", (char*)text);
							#endif
							cmdPtr->Execute(NULL, NULL, adCmdText);
						}		
						break;
				}
				statmentLogged = true;
				attemptRetry = false;
			}
			catch( _com_error &e)
			{
				if (maxRetry > 0)	
					maxRetry--;	
				else 
				{
					_bstr_t eDescrip = e.Description();
					throw e;
				}
			}
		}
		// Push primary key value
		// Note: In general when updating stub there is no need to push key because it was pushed when the stub was created and will be popped off.
		//       If a wait step stub was created, it was in a different context, so we need to push a dummy key in this context
		if (eType != eUpdateStub || (eType == eUpdateStub && bWaitStep) )
		{
			bool bKeyValuePushed = false;
	        CDbColumn *pColumn = NULL;

			if (m_primaryKeyColumns.GetCount())
				pColumn = m_primaryKeyColumns.GetHead(); //assuming only one key column in list

			if (m_type == tsDBOptionsStatementType_RecordSet)
			{
				// Get primary key column value and push onto the stack
				if (pColumn)
				{
					PushKeyValue(loggingInfo, RsITEM(m_spAdoRecordSet,pColumn->GetName()));
					bKeyValuePushed = true;
				}
				else 
				{
					// Find the primary key column if the user did not specify a primary key column in the column settings, 
					for (short i=0; (!bKeyValuePushed) && (i < m_spAdoRecordSet->Fields->Count); i++)
					{
						_variant_t vIndex = i;
						if (adFldKeyColumn & m_spAdoRecordSet->Fields->Item[vIndex]->Attributes) //found key column
						{
							PushKeyValue(loggingInfo, RsITEM(m_spAdoRecordSet,vIndex));
							bKeyValuePushed = true;
						}
					}
				}
			}
			else  ///prepared statement
			{	
				if (pColumn)
				{
					try
					{
						PushKeyValue(loggingInfo, CmdITEM(cmdPtr,pColumn->GetName()));
						bKeyValuePushed = true;
					}
					catch( _com_error &e)
					{
						// LOOSE END: Alex added this try/catch here, but I am not sure why?
						_bstr_t eDescrip = e.Description();
					}
				}
			}

			// The caller will pop a key off the stack, so push a dummy key
			if (!bKeyValuePushed)
			{
				_variant_t vt;
				vt.vt = VT_NULL;
				PushKeyValue(loggingInfo, vt);
			}
		}
Leave:
		// Update logging options based on whether this statement logged
		if (statmentLogged)
			traverseOptions |= m_logOptions;
    }

	catch (TSDBLogException &e) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForStatement");
		description += m_name;
		description += ".\n";
		description += e.mDescription;
		e.mDescription = description;
		throw e;
	}		
	catch (_com_error &ce) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForStatement");
		description += m_name;
		description += ".\n";
		description += ::GetTextForAdoErrors(this->GetDatalink()->m_spAdoCon, ce.Description());

		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToExecuteStatement, ce.Error(), description);
	}
	catch (...) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForStatement");
		description += m_name;
		description += ".\n";
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToExecuteStatement,0 , description);
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Returns the key list associated with the thread or execution id
CLastKeyList* Statement::GetLastKeyListFromId(long currentId)
{
	POSITION pos = m_lastKeyThreadList.GetHeadPosition();
	while (pos)	{
	   CLastKeyList* lastKeyList = m_lastKeyThreadList.GetNext(pos);
	   if (lastKeyList->threadId == currentId) 
		   return lastKeyList;
	}
	return NULL;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Returns the key list associated with thread or execution id for the context we are logging.
CLastKeyList* Statement::GetLastKeyList(LoggingInfo &loggingInfo)
{
	//UUTs are execution specific, all others are thread specific
	long currentId;
	if (GetResultType() == tsDBOptionsResultType_UUT)
		currentId = loggingInfo.GetContext()->GetExecution()->GetId();
	else
		currentId = loggingInfo.GetContext()->GetThread()->GetId();

	POSITION pos = m_lastKeyThreadList.GetHeadPosition();
	while (pos)
	{
	   CLastKeyList* lastKeyList = m_lastKeyThreadList.GetNext(pos);
	   if (lastKeyList->threadId == currentId) 
		   return lastKeyList;
	}
	return NULL;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Returns last primary key value for the statement
_variant_t Statement::GetLastKeyValue(LoggingInfo &loggingInfo) 
{
	CCritSectLock cs(&m_keyCritSect);
	CKeyItem *lastKey = NULL;
	_variant_t vT;

	vT.vt = VT_NULL;
	lastKey = GetLastKeyItem(loggingInfo);
	if (lastKey != NULL) 
		vT = lastKey->keyValue;

	#ifdef _DEBUG
		DebugPrintf("GetLastKeyValue:  Key=%s, TsThreadId=%d \n", (char*)(vT.vt==1?"NULL":(_bstr_t)vT), loggingInfo.GetContext()->GetThread()->GetId());
	#endif

	return vT;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Returns primary key value at a specific depth
_variant_t Statement::GetKeyValueForDepth(LoggingInfo &loggingInfo, int depth)
{
	CCritSectLock cs(&m_keyCritSect);
	CLastKeyList* lastKeyList;
	CKeyItem *lastKey = NULL;
	_variant_t vT;
	vT.vt = VT_NULL;

	lastKeyList = GetLastKeyList(loggingInfo);
	if (lastKeyList && !lastKeyList->keyItems.IsEmpty()) 
	{
		POSITION pos = lastKeyList->keyItems.GetHeadPosition();
		while (pos) {
			lastKey = lastKeyList->keyItems.GetNext(pos);
			if (lastKey->keyDepth == depth) {
				vT = lastKey->keyValue;
				break;
			}
		}
	}
	#ifdef _DEBUG
		DebugPrintf("GetKeyForDepth:   Key=%s, Depth=%d\n", (char*)(vT.vt==1?"NULL":(_bstr_t)vT), depth);
	#endif
	return vT;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Returns the last primary key item logged for the statement.
CKeyItem *Statement::GetLastKeyItem(LoggingInfo &loggingInfo) 
{
	CCritSectLock cs(&m_keyCritSect);
	CLastKeyList* lastKeyList;

	lastKeyList = GetLastKeyList(loggingInfo); // Get key list for this context
	if (lastKeyList && !lastKeyList->keyItems.IsEmpty()) 
		return lastKeyList->keyItems.GetTail();
	else
		return NULL;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Pops primary key value for the current thread and depth when a statement is logged. This is called as 
// we crawl deeper into recursive calls to log results and properties.
void Statement::PushKeyValue(LoggingInfo &loggingInfo, _variant_t keyValue) 
{
	TS::SequenceContextPtr& pContext = loggingInfo.GetContext();
	CCritSectLock cs(&m_keyCritSect);
	CLastKeyList* lastKeyList;
	CKeyItem *lastKey;

	lastKeyList = GetLastKeyList(loggingInfo);

	lastKey = new CKeyItem();
	lastKey->keyValue = keyValue;
	lastKey->keyDepth = loggingInfo.GetLoggingDepth();
	lastKey->executionOrder = loggingInfo.GetExecutionOrder();

	#ifdef _DEBUG
		DebugPrintf("PushKeyValue:     Stmt=%s, Key=%s, StackDepth=%d, KeyDepth=%d, LogDepth=%d, KeyOrder is %d, TsThreadId=%d\n", (char*)m_name, (char*)(keyValue.vt==1?"NULL":(char*)(_bstr_t)keyValue), loggingInfo.GetContext()->CallStackDepth, lastKey->keyDepth, loggingInfo.GetLoggingDepth(), lastKey->executionOrder, loggingInfo.GetContext()->GetThread()->GetId());
	#endif
	// Create new threadlist and add keyvalue to it
	if (lastKeyList) 
   		lastKeyList->keyItems.AddTail(lastKey); 
	else 
	{
		lastKeyList = new CLastKeyList();
		//UUTs are execution specific, all others are thread specific
		if ( GetResultType() == tsDBOptionsResultType_UUT )
			lastKeyList->threadId = pContext->GetExecution()->GetId();
		else
			lastKeyList->threadId = pContext->GetThread()->GetId();

		lastKeyList->keyItems.AddTail(lastKey);
		m_lastKeyThreadList.AddTail(lastKeyList);
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Pops any primary key value for the current thread and depth. This is called as we crawl back up the 
// recursive calls to log results and properties.
void Statement::PopKeyValue(LoggingInfo &loggingInfo, bool clearAllKeys)
{
	CCritSectLock cs(&m_keyCritSect);
	CLastKeyList* lastKeyList = GetLastKeyList(loggingInfo);

	if (lastKeyList) 
	{
		bool stop = false;
	    while (!lastKeyList->keyItems.IsEmpty() && (!stop || clearAllKeys)) 
		{
			CKeyItem *lastKey = lastKeyList->keyItems.GetTail();
			if (clearAllKeys || lastKey->keyDepth >= loggingInfo.GetLoggingDepth())
			{
				lastKey = lastKeyList->keyItems.RemoveTail();
				#ifdef _DEBUG
					_variant_t vT;
					vT.vt = VT_NULL;
					vT = lastKey->keyValue;
					if (clearAllKeys && lastKey->keyDepth < loggingInfo.GetLoggingDepth())
						DebugPrintf("PopKeyValue:******Stmt=%s, Key=%s, StackDepth=%d, KeyDepth=%d, LogDepth=%d, KeyOrder is %d, TsThreadId=%d\n", (char*)m_name, (char*)(vT.vt==1?"NULL":(char*)(_bstr_t)vT), loggingInfo.GetContext()->CallStackDepth, lastKey->keyDepth, loggingInfo.GetLoggingDepth(), lastKey->executionOrder, loggingInfo.GetContext()->GetThread()->GetId());
					else
						DebugPrintf("PopKeyValue:      Stmt=%s, Key=%s, StackDepth=%d, KeyDepth=%d, LogDepth=%d, KeyOrder is %d, TsThreadId=%d\n", (char*)m_name, (char*)(vT.vt==1?"NULL":(char*)(_bstr_t)vT), loggingInfo.GetContext()->CallStackDepth, lastKey->keyDepth, loggingInfo.GetLoggingDepth(), lastKey->executionOrder, loggingInfo.GetContext()->GetThread()->GetId());
				#endif
				delete lastKey;
			} else {
				#ifdef _DEBUG
					_variant_t vT;
					vT.vt = VT_NULL;
					vT = lastKey->keyValue;
					// DebugPrintf("!PopKeyValue:     Stmt=%s, Key=%s, StackDepth=%d, KeyDepth=%d, LogDepth=%d, KeyOrder is %d, TsThreadId=%d\n", (char*)m_name, (char*)(vT.vt==1?"NULL":(char*)(_bstr_t)vT), loggingInfo.GetContext()->CallStackDepth, lastKey->keyDepth, loggingInfo.GetLoggingDepth(), lastKey->executionOrder, loggingInfo.GetContext()->GetThread()->GetId());
				#endif
				stop = true;
			}
		}
		
		// Be proactive and delete key list if empty
		POSITION pos;
		if (lastKeyList->keyItems.IsEmpty() && (pos = m_lastKeyThreadList.Find(lastKeyList))) 
		{
				m_lastKeyThreadList.RemoveAt(pos);
				delete lastKeyList;
		}
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Commits records for a recordset with batch optimistic cursor
void Statement::Commit()
{
	if ( m_type == tsDBOptionsStatementType_RecordSet && m_lockType == adLockBatchOptimistic)
		m_spAdoRecordSet->UpdateBatch(adAffectAll);
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Close recordset object
void Statement::Close()
{
	if (m_type == tsDBOptionsStatementType_RecordSet && m_opened)
	{
		try {
			m_spAdoRecordSet->Close();
		}
		catch (...) {}
	}
	m_opened = false;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Creates a stub result for a result that has a parent foreign key reference
void Statement::CreateStubResult(LoggingInfo &loggingInfo)
{
	StatementTraverseOptions ignoreOptions = tsDBOptionsTraverseOptions_Continue;
	Execute(loggingInfo, ignoreOptions, eCreateStub);
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Determine whether the stub results that this statement references as foreign keys are present, if not, create them
void Statement::CreateParentStubs(LoggingInfo &loggingInfo)
{
	_COM_ASSERT(GetResultType() == tsDBOptionsResultType_Step);
	TS::SequenceContextPtr& pContext = loggingInfo.GetContext();
	// lock the statement so we can protect column values and keys
	CCritSectLock cs(&m_critSect);
	Statement *pStmt;
	POSITION pos;
	CKeyItem *lastKey;
	bool parentStepResultExists = (pContext->Caller != NULL);

	#ifdef _DEBUG
		DebugPrintf("CreateParent1     Depth=%d, Step=%s, Stmt=%s\n", loggingInfo.GetLoggingDepth(), (char *)pContext->Step->Name, (char*)m_name);
	#endif

	// Loop thru the recursive key columns and see if we need to go deeper first before creating stub results at this level
	if (parentStepResultExists) {
		pos = m_recursiveKeyColumns.GetHeadPosition();
		while (pos) {
			// Get statement with primary key
			CDbColumn *pColumn = m_columns.GetNext(pos);
			if (pStmt = pColumn->GetForeignKeyStmt()) {
				// Check if stub exists for that primary key
				// If immediate parent stub is missing, consider going to its caller
				lastKey = pStmt->GetLastKeyItem(loggingInfo); 
				if (lastKey == NULL || lastKey->keyDepth < loggingInfo.GetLoggingDepth() - 1)
				{
					bool loggingCreated = false;
					LoggingInfo callerInfo(pContext->Caller, LoggingInfo::LoggingInfoType_Copy, false, loggingCreated, true);
					CreateParentStubs(callerInfo);
				}
			}
		}
	}

	// Loop thru all foreign key columns and determine if we need to create a stub result
	pos = m_foreignKeyColumns.GetHeadPosition();
	while (pos) {
		// Get statement with primary key
		CDbColumn *pColumn = m_columns.GetNext(pos);
		bool isStepRecursiveKey = pColumn->IsStepRecursiveKey();
		
		if (!isStepRecursiveKey || parentStepResultExists) 
		{
			if (pStmt = pColumn->GetForeignKeyStmt()) {
				// Check if stub exists for that primary key
				lastKey = pStmt->GetLastKeyItem(loggingInfo);

				// Create stub if no key exists for the statement and it is either a parent step result or a UUT result
				// Create stub if a key exists, but the referred step statement does not have a key for the current level
				if ((lastKey == NULL && (isStepRecursiveKey || pStmt->GetResultType() == tsDBOptionsResultType_UUT)) || 
					(lastKey && pStmt->GetResultType() == tsDBOptionsResultType_Step && lastKey->keyDepth < loggingInfo.GetLoggingDepth() - 1))
				{
					if (isStepRecursiveKey) {
						loggingInfo.CreatingStubStepResult(true);
						// For async calls, we want to create the stub result in the same thread as the called result. 
						// Later if and when the wait step accepts the results, the async step updates the stub. 
						loggingInfo.DecLoggingDepth(); // It is easier to decrease the logging depth than creating a new LoggingInfo.
					}
					#ifdef _DEBUG
						if (!lastKey)
							DebugPrintf("CreateStubResult: KeyDepth=%s, LoggingDepth=%d, Stmt=%s\n", "(empty)", loggingInfo.GetLoggingDepth(), (char*)m_name);
						else
							DebugPrintf("CreateStubResult: KeyDepth=%d, LoggingDepth=%d, Stmt=%s\n", lastKey->keyDepth, loggingInfo.GetLoggingDepth(), (char*)m_name);
					#endif
					pStmt->CreateStubResult(loggingInfo);
					if (isStepRecursiveKey) {
						loggingInfo.IncLoggingDepth();
						loggingInfo.CreatingStubStepResult(false);
					}
				}
			}
		}
	}

	#ifdef _DEBUG
		DebugPrintf("CreateParent2     Depth=%d, Step=%s, Stmt=%s\n", loggingInfo.GetLoggingDepth(), (char *)pContext->Step->Name, (char*)m_name);
	#endif
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Log the result with the statement
void Statement::ExecuteForOnTheFly(LoggingInfo &loggingInfo, StatementTraverseOptions &traverseOptions)
{
	TS::SequenceContextPtr& pContext = loggingInfo.GetContext();
	// Lock the statement so we can protect column values and keys
	CCritSectLock cs(&m_critSect);

	// Determine if there is a stub already created for this result
	CKeyItem *lastKey = GetLastKeyItem(loggingInfo); 
	#ifdef _DEBUG
		if (!lastKey)
			DebugPrintf("ExecuteOnTheFly:  KeyDepth=%s, LoggingDepth=%d, Stmt=%s\n", "(empty)", loggingInfo.GetLoggingDepth(), (char*)m_name);
		else
			DebugPrintf("ExecuteOnTheFly:  KeyDepth=%d, LoggingDepth=%d, Stmt=%s\n", lastKey->keyDepth, loggingInfo.GetLoggingDepth(), (char*)m_name);
	#endif
	_COM_ASSERT(!lastKey || lastKey->keyDepth <= loggingInfo.GetLoggingDepth()); 
	if (lastKey && lastKey->keyDepth >= loggingInfo.GetLoggingDepth()) 
	{ 
		//we are returning from deeper within the call stack, so update an existing stub result
		Execute(loggingInfo, traverseOptions, eUpdateStub); 
	} else {
		 bool bWaitStep = (loggingInfo.GetStepResult()->Exists(PROP_ASYNCID, 0) && loggingInfo.GetStepResult()->Exists(PROP_TS_SEQUENCECALL_STATUS, 0));

		// Create a new result unless we are a Wait step.  
		// For a Wait step, we ned to update the async seq's parent stub with the wait step's result. This will make the async seq's parent the wait step
		if (bWaitStep) 
			Execute(loggingInfo, traverseOptions, eUpdateStub);
		else
			Execute(loggingInfo, traverseOptions, eLogResult);
	}
}
