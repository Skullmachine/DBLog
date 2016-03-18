
#include "stdafx.h"
#include "Columns.h"
#include "Statement.h"
#include "DBLog.h"
#include "TSDatalink.h"
#include <rpc.h>
#include "tsapivc.h"	// Header file found in <TestStand>/api/vc/
#include "tsvcutil.h"	// Header file found in <TestStand>/api/vc/

#define STRING_DELIMITER L"\r\n"
#define STRING_DELIMITER_LEN 2
#define APPEND_CHUNK_BLOCK_SIZE 4096
#define BSTRSIZE sizeof(wchar_t)
#define GET_ENUM(et,str,mv) {tempVal = pDBColumn->GetValNumber(str,0); \
                             tempLong = (long)tempVal; \
                             mv = (et)tempLong; }

#define COERCE_FROM_OPTIONS PropOption_CoerceFromNumber | PropOption_CoerceFromString | PropOption_CoerceFromBoolean | PropOption_CoerceFromReference | PropOption_DecimalPoint_UsePeriod

extern  int FormatToVariant(
	LCID lcid,
	char* value,
	char* inputFmt,
	_variant_t &valueVariant);

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// DBMemoryBuilder class
// This class holds intrinsic and array values to be stored to Bstr, VarChar, GUID, and Binary column types.
DBMemoryBuilder::DBMemoryBuilder(ColDataType type) 
{
	_COM_ASSERT(type == tsDBOptionsColumnType_Binary || 
				type == tsDBOptionsColumnType_Bstr || 
				type == tsDBOptionsColumnType_VarChar ||
				type == tsDBOptionsColumnType_GUID);
	
	m_type = type;
	m_totalBytes = 0;
}

//-------------------------------------------------------------------------------------------------
// Converts property object value to list based on binary or string representation for writing to 
// database field
//-------------------------------------------------------------------------------------------------
// The following table summarizes how data is processed
//
// Value Type						Binary Field			VarChar/GUID Field
// ----------						------------			------------------
// String							BSTR					BSTR
// Number							R8 or format			BSTR (uses format)
// Boolean							I1						0 or 1
// Array(String)					BSTR (NUL delimeted)	BSTR (\r\n delimeted, escaped)
// Array(Number)					R8[] or format			BSTR (uses format, \r\n delimeted)
// Array(Boolean)					I1[]					0 or 1 (\n delimeted)
// Array of Container (String)		BSTR (NUL delimeted)	BSTR (\n delimeted, escaped, '[name]=' prefix)
// Array of Container (Numeric)		R8[] or format			BSTR (uses format, \r\n delimeted)
// Array of Container (Boolean)		I1[]					0 or 1 (\r\n delimeted)
// 
// Reference						error					error
// Container						error					error
// Array(Reference)					error					error
// Array(Container)					error					error
void DBMemoryBuilder::Append(TS::PropertyObjectPtr &pValue, _bstr_t &format, int depth, _bstr_t prefix)
{
	TS::PropertyObjectTypePtr pType = pValue->Type;
	
	switch (pType->ValueType)
	{
		case TS::PropValType_Number:
			AppendNumber(pValue, GetNumTypeFromFormat(format, VT_R8));
			break;
		case TS::PropValType_String:
			AppendString(pValue, prefix, (m_type != tsDBOptionsColumnType_Binary) && (depth > 1));
			break;
		case TS::PropValType_Boolean:
			AppendBoolean(pValue, GetNumTypeFromFormat(format, VT_UI1));
			break;
		case TS::PropValType_Reference:
			RAISE_ERROR_WITH_DESC(TSDBLogErr_ReferenceTypeNotSupported, TS::TS_Err_UnexpectedType, "");
			break;
		case TS::PropValType_Container:
			RAISE_ERROR_WITH_DESC(TSDBLogErr_ContainerTypeNotSupported, TS::TS_Err_UnexpectedType, "");
			break;
		case TS::PropValType_Array:
		{
			switch (pType->ElementType->ValueType)
			{
				case TS::PropValType_Number:
					AppendNumericArray(pValue, GetNumTypeFromFormat(format, VT_R8));
					break;
				case TS::PropValType_String:
					AppendStringArray(pValue);
					break;
				case TS::PropValType_Boolean:
					AppendBooleanArray(pValue, GetNumTypeFromFormat(format, VT_UI1));
					break;
				case TS::PropValType_Reference:
					RAISE_ERROR_WITH_DESC(TSDBLogErr_ReferenceTypeNotSupported, TS::TS_Err_UnexpectedType, "");
					break;
				case TS::PropValType_Container:
				case TS::PropValType_Array:
				{
					if (depth > 1)
						break;

					// Loop on container array elements
					long elements = pValue->GetNumElements();
					for (long i = 0; i < elements; i++) {
						TS::PropertyObjectPtr pElement = pValue->GetPropertyObjectByOffset(i, 0);

						if (pElement->Type->ValueType == TS::PropValType_String)	{
							CStringW prefix;
							bstr_t name = pElement->Name;
							if (name.length()) 
								prefix.AppendFormat(L"[%s]=", (wchar_t*)name);
							else
								prefix.AppendFormat(L"%s=", (wchar_t*)pValue->GetArrayIndex("", 0, i));
							Append(pElement, format, ++depth, (_bstr_t)prefix);
						} else
							Append(pElement, format, ++depth, "");
					}
					break;
				}
			}
		}
	}
}

// Adds numeric value to list as variant, or formatted BSTR value
void DBMemoryBuilder::AppendNumber(TS::PropertyObjectPtr &pValue, VARTYPE numType)
{
	variant_t vValue;
	if (m_type == tsDBOptionsColumnType_Binary) {
		vValue = pValue->GetValNumber("", 0);
		vValue.ChangeType(numType);
		m_totalBytes += GetNumTypeSizeFromFormat(numType);
	} else {
#if 1
		// TS conversion uses 13 significant digits
		vValue = pValue->GetFormattedValue("", COERCE_FROM_OPTIONS, "%$.13f", VARIANT_FALSE, "");
#else
		// ChangeType conversion uses 15 significant digits, but will store string value that is not rounded properly for whole numbers
		vValue = pValue->GetValNumber("", 0);
		vValue.ChangeType(VT_BSTR);
#endif
		m_totalBytes += SysStringLen(vValue.bstrVal) * BSTRSIZE;
	}
	m_Values.insert(m_Values.end(), vValue);
}

// Adds string to list as BSTR value, and escape '\', '\r', and '\n' for non-binary fields
void DBMemoryBuilder::AppendString(TS::PropertyObjectPtr &pValue, _bstr_t prefix, bool escapeCRLF)
{
	bstr_t value = pValue->GetValString("", 0);

	if (prefix.length()) {
		CStringW newValue((wchar_t *)prefix);
		newValue.Append(value);
		value = newValue;
	}
	if (escapeCRLF)
		EscapeCRLF(value);

	m_Values.insert(m_Values.end(), value);
	m_totalBytes += value.length() * BSTRSIZE;
}

// Adds Boolean value to list as variant, or BSTR value
void DBMemoryBuilder::AppendBoolean(TS::PropertyObjectPtr &pValue, VARTYPE numType)
{
	VARIANT_BOOL value = pValue->GetValBoolean("", 0);

	if (m_type == tsDBOptionsColumnType_Binary) {
		variant_t vValue((value)?1:0);
		vValue.ChangeType(numType);
		m_totalBytes += GetNumTypeSizeFromFormat(numType);
		m_Values.insert(m_Values.end(), vValue);
	} else {
		variant_t vValue((value)?L"1":L"0");
		m_totalBytes += SysStringLen(vValue.bstrVal) * BSTRSIZE;
		m_Values.insert(m_Values.end(), vValue);
	}
}

// Returns the number of elements in a variant safearray, also returns byte-size of an element
unsigned int DBMemoryBuilder::GetNumArrayElements(variant_t vArray, unsigned int *elementSize)
{
	COleSafeArray sa(vArray);
	_COM_ASSERT(((vArray.vt & VT_ARRAY) == VT_ARRAY) && (sa != NULL));

	unsigned int numElements = 1;
	if (elementSize)
		*elementSize = sa.GetElemSize();
	unsigned int dimensions = sa.GetDim();
	if (dimensions == 1)
        numElements = sa.GetOneDimSize();
	else {
		for (unsigned int i=1;i<=dimensions;i++) {
			long lBound, uBound;
			sa.GetLBound(i, &lBound);
			sa.GetUBound(i, &uBound);
			numElements *= (uBound - lBound + 1);
		}
	}
	return numElements;
}

// Adds numeric array as variant safearray, or string values
void DBMemoryBuilder::AppendNumericArray(TS::PropertyObjectPtr &pValue, VARTYPE numType /*, wchar_t *delimeter */)
{
	variant_t vArray = pValue->GetValVariant("", 0);
	_COM_ASSERT(vArray.vt == (VT_R8 | VT_ARRAY));
	SAFEARRAY *psa = vArray.parray;
	if (psa == NULL) 
		return;

	unsigned int numElements = GetNumArrayElements(vArray);

	if (m_type == tsDBOptionsColumnType_Binary) {
		ChangeArrayType(vArray, numType);
		m_totalBytes += GetNumTypeSizeFromFormat(numType) * numElements;
		m_Values.insert(m_Values.end(), vArray);
	} else {
		// Get a pointer to the elements of the array.
		double HUGEP *safeArray = NULL;
		HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&safeArray);
		if (FAILED(hr)) {
			_com_issue_error(hr);
		}

		m_Values.reserve(m_Values.size() + numElements);

		wchar_t buffer[64];
		for (unsigned int i = 0; i < numElements; i++) {
			swprintf_s(buffer, 64, L"%f", safeArray[i]);
			m_totalBytes += wcslen(buffer) * BSTRSIZE;
			m_Values.insert(m_Values.end(), buffer);
		}

		hr = SafeArrayUnaccessData(psa);
		if (FAILED(hr)) {
			_com_issue_error(hr);
		}
	}
}

// Adds string array elements as BSTR values to list, and escape '\', '\r', and '\n' for non-binary fields
void DBMemoryBuilder::AppendStringArray(TS::PropertyObjectPtr &pValue /*, wchar_t *delimeter, bool escapeCRLF*/)
{
	variant_t vArray = pValue->GetValVariant("", 0);
	_COM_ASSERT(vArray.vt == (VT_BSTR | VT_ARRAY));
	SAFEARRAY *psa = vArray.parray;
	if (psa == NULL) 
		return;	

	unsigned int numElements = GetNumArrayElements(vArray);

	// Get a pointer to the elements of the array.
	BSTR HUGEP *safeArray = NULL;
	HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&safeArray);
	if (FAILED(hr)) {
		_com_issue_error(hr);
	}

	m_Values.reserve(m_Values.size() + numElements);
	for (unsigned int i = 0; i < numElements; i++) 
	{
		bstr_t value(safeArray[i], true);
		if (m_type != tsDBOptionsColumnType_Binary)
			EscapeCRLF(value);

		m_totalBytes += value.length() * BSTRSIZE;
		m_Values.insert(m_Values.end(), value);
	}

	hr = SafeArrayUnaccessData(psa);
	if (FAILED(hr)) {
		_com_issue_error(hr);
	}
}

// Adds Boolean array as variant safearray, or single string value of 0's and 1's delimited by \r\n
void DBMemoryBuilder::AppendBooleanArray(TS::PropertyObjectPtr &pValue, VARTYPE numType /*, wchar_t *delimeter*/ )
{
	variant_t vArray = pValue->GetValVariant("", 0);
	_COM_ASSERT(vArray.vt == (VT_BOOL | VT_ARRAY));
	SAFEARRAY *psa = vArray.parray;
	if (psa == NULL) 
		return;

	unsigned int numElements = GetNumArrayElements(vArray);

	// If binary, just add array to list
	if (m_type == tsDBOptionsColumnType_Binary)
	{
		ChangeArrayType(vArray, numType);
		m_totalBytes += GetNumTypeSizeFromFormat(numType) * numElements;
		m_Values.insert(m_Values.end(), vArray);
	}
	else // Add booleans as strings to list
	{
		// Get a pointer to the elements of the array.
		VARIANT_BOOL HUGEP *safeArray = NULL;
		HRESULT hr = SafeArrayAccessData(psa, (void HUGEP**)&safeArray);
		if (FAILED(hr)) {
			_com_issue_error(hr);
		}

		CByteArray bytes;
		unsigned int numChars = numElements + (numElements - 1) * STRING_DELIMITER_LEN;
		bytes.SetSize(numChars * BSTRSIZE);
		wchar_t* pValue = (wchar_t*) bytes.GetData();

		for (unsigned int i = 0; i < numElements; i++) 
		{
			if (i) {
				wcsncpy(pValue, STRING_DELIMITER, STRING_DELIMITER_LEN);
				pValue += STRING_DELIMITER_LEN;
			}
			*pValue = (safeArray[i]) ? L'1' : L'0';
			++pValue;
		}
		bstr_t value(SysAllocStringLen((wchar_t*) bytes.GetData(), numChars), false);
		_COM_ASSERT(numChars == value.length());
			
		m_totalBytes += value.length() * BSTRSIZE;
		m_Values.insert(m_Values.end(), value);

		hr = SafeArrayUnaccessData(psa);
		if (FAILED(hr)) {
			_com_issue_error(hr);
		}
	}
}

// Change the element type of a R8 or BOOL variant safearray, otherwise no change is done
void DBMemoryBuilder::ChangeArrayType(variant_t &vArray, VARTYPE vtNewType)
{
	COleSafeArray saOrig(vArray);
	VARTYPE vtOrigType = vArray.vt & VT_TYPEMASK;
	int origSize = GetNumTypeSizeFromFormat(vtOrigType);
	unsigned int numElements = GetNumArrayElements(vArray);
	unsigned int i;

	if (vtOrigType == vtNewType) 
		return;

	COleSafeArray saNew;
	DWORD bounds[] = {numElements};
	saNew.Create(vtNewType, 1, bounds);

	if (vtOrigType == VT_R8) 
	{
		double *sourceData;
		saOrig.AccessData((void**)&sourceData);
		if (vtNewType == VT_R4) {
			float *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (float) *sourceData;
			saNew.UnaccessData();
		} else if (vtNewType == VT_UI4) {
			unsigned int *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (unsigned int) *sourceData;
			saNew.UnaccessData();
		} else if (vtNewType == VT_I4) {
			int *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (int) *sourceData;
			saNew.UnaccessData();
		} else if (vtNewType == VT_UI2) {
			unsigned short *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (unsigned short) *sourceData;
			saNew.UnaccessData();
		} else if (vtNewType == VT_I2) {
			short *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (short) *sourceData;
			saNew.UnaccessData();
		} else if (vtNewType == VT_UI1) {
			unsigned char *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (unsigned char) *sourceData;
			saNew.UnaccessData();
		} else if (vtNewType == VT_I1) {
			char *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++) 
				*destData = (char) *sourceData;
			saNew.UnaccessData();
		}
		else {
			_COM_ASSERT(!"Code does not support converting R8 array to the requested type");
			return;
		}
		saOrig.UnaccessData();
		vArray.Attach(saNew.Detach());
	} else if (vtOrigType == VT_BOOL) {
		VARIANT_BOOL *sourceData;
		saOrig.AccessData((void**)&sourceData);
		if (vtNewType == VT_R8) {
			double *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (double) (*sourceData)?1.0:0.0;
			saNew.UnaccessData();
		} else if (vtNewType == VT_R4) {
			float *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (float) (*sourceData)?(float)1.0:(float)0.0;
			saNew.UnaccessData();
		} else if (vtNewType == VT_UI4) {
			unsigned int *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (unsigned int) (*sourceData)?1:0;
			saNew.UnaccessData();
		} else if (vtNewType == VT_I4) {
			int *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (int) (*sourceData)?1:0;
			saNew.UnaccessData();
		} else if (vtNewType == VT_UI2) {
			unsigned short *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (unsigned short) (*sourceData)?1:0;
			saNew.UnaccessData();
		} else if (vtNewType == VT_I2) {
			short *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (short) (*sourceData)?1:0;
			saNew.UnaccessData();
		} else if (vtNewType == VT_UI1) {
			unsigned char *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++)
				*destData = (unsigned char) (*sourceData)?1:0;
			saNew.UnaccessData();
		} else if (vtNewType == VT_I1) {
			char *destData;
			saNew.AccessData((void **)&destData);
			for (i = 0; i < numElements; i++,destData++,sourceData++) 
				*destData = (char) (*sourceData)?1:0;
			saNew.UnaccessData();
		}
		else {
			_COM_ASSERT(!"Code does not support converting BOOL array to the requested type");
			return;
		}
		saOrig.UnaccessData();
		vArray.Attach(saNew.Detach());
	}
	else {
		_COM_ASSERT(!"Cannot convert array data, unexpected array type");
		return;
	}
}

// Write elements of list to database field
void DBMemoryBuilder::Write(ADOLoggingObject &pDBObj, _bstr_t &name, unsigned int maxByteSize)
{
	if (m_Values.size() == 0) {
		_COM_ASSERT(m_totalBytes == 0);
		if (pDBObj.IsCommand()) { // write null if command, not recordset
			_variant_t vtNull;
			vtNull.ChangeType(VT_NULL);
			pDBObj.Assign(name, vtNull); 
		}
	} else {
		if (m_type == tsDBOptionsColumnType_Binary) 
			WriteBinaryValues(pDBObj, name, maxByteSize);
		else
			WriteStringValues(pDBObj, name, maxByteSize);
	}
}

// Build byte array of all string values, delimit elements by \r\n, truncate if necessary, and write to DB
void DBMemoryBuilder::WriteStringValues(ADOLoggingObject &pDBObj, _bstr_t &name, unsigned int maxByteSize)
{
	_COM_ASSERT(m_totalBytes % BSTRSIZE == 0);

	// Build strings into byte array
	CByteArray bytes;
	unsigned int numBytes = m_totalBytes + ((m_Values.size() - 1) * STRING_DELIMITER_LEN * BSTRSIZE);
	bytes.SetSize(numBytes * BSTRSIZE);
	wchar_t* pValue = (wchar_t*) bytes.GetData();

	std::vector<_variant_t>::iterator item;
	for (item = m_Values.begin(); item != m_Values.end(); ++item)
	{
		if (item->vt != VT_BSTR)
			RAISE_ERROR_WITH_DESC(TSDBLogErr_ContainerTypeHeterogeneous, TS::TS_Err_UnexpectedType, "");

		if (item != m_Values.begin()) {
			wcsncpy(pValue, STRING_DELIMITER, STRING_DELIMITER_LEN);	// Add delimiter
			pValue += STRING_DELIMITER_LEN;
		}

		bstr_t itemStr(*item);
		wcsncpy(pValue, itemStr, itemStr.length());		// Add element
		pValue += itemStr.length();
	}
	
	// Write string to DB
	if (m_type == tsDBOptionsColumnType_Bstr || m_type == tsDBOptionsColumnType_GUID) 
	{
		unsigned int bstrSize = ((maxByteSize > 0) ? min(maxByteSize, numBytes) : numBytes) / BSTRSIZE;
		bstr_t bstrValue(SysAllocStringLen((wchar_t*) bytes.GetData(), bstrSize), false);
		_COM_ASSERT(bstrValue.length() == bstrSize);

		pDBObj.Assign(name, bstrValue);
		#ifdef _DEBUG
		DebugPrintf("SetColumn:        %s='%s'\n", (char*)name, (char*)bstrValue);
		#endif
	} else { 
		_COM_ASSERT (m_type == tsDBOptionsColumnType_VarChar);
		bstr_t bstrValue(SysAllocStringLen((wchar_t*) bytes.GetData(), numBytes / BSTRSIZE), false);
		_COM_ASSERT(bstrValue.length() == numBytes / BSTRSIZE);
		char *tempPtr = (char *)bstrValue;
		if (maxByteSize > 0 && maxByteSize < strlen(tempPtr)) 
			tempPtr[maxByteSize] = 0;
		pDBObj.Assign(name, (variant_t)tempPtr);
		#ifdef _DEBUG
		DebugPrintf("SetColumn:        %s='%s'\n", (char*)name, tempPtr);
		#endif
	}
}

// Loop on list and write binary values to DB
void DBMemoryBuilder::WriteBinaryValues(ADOLoggingObject &pDBObj, _bstr_t &name, unsigned int maxByteSize)
{
	unsigned int totalBytesWritten = 0, bytesWritten;
	std::vector<_variant_t>::iterator item;
	bool stop = false;
	bool stringElements = false;

	for (item = m_Values.begin(); item != m_Values.end() || stop; ++item)
	{
		if ((item->vt & VT_TYPEMASK) == VT_BSTR)
			stringElements = true;
		else if (stringElements)
			RAISE_ERROR_WITH_DESC(TSDBLogErr_ContainerTypeHeterogeneous, TS::TS_Err_UnexpectedType, "");

		if ((item->vt & VT_ARRAY) == VT_ARRAY)
			 WriteBinaryArray(pDBObj, name, *item, maxByteSize, bytesWritten, stop);
		else
			 WriteBinaryValue(pDBObj, name, *item, maxByteSize, bytesWritten, stop);
		
		totalBytesWritten += bytesWritten;
		if (maxByteSize > 0) {
			if (bytesWritten >= maxByteSize)
				break;
			else
				maxByteSize -= bytesWritten;
		}
	}

	// Add terminator character if binary string values were written
	if (stringElements)
		WriteBinaryZeros(pDBObj, name, BSTRSIZE);

	if (!totalBytesWritten && pDBObj.IsCommand()) {
		_variant_t vtNull;
		vtNull.ChangeType(VT_NULL);
		pDBObj.Assign(name, vtNull); 
	}
}

// Write binary value to DB as safearray. Will add NULL delimiter for string values
void DBMemoryBuilder::WriteBinaryValue(ADOLoggingObject &pDBObj, _bstr_t &name, variant_t &value, unsigned int maxByteSize, unsigned int &bytesWritten, bool &stop)
{
	unsigned int valueBytes = 0;
	unsigned int totalBytes = 0;
	void *pSource = NULL;
	bytesWritten = 0;
	_COM_ASSERT((value.vt & VT_ARRAY) == 0);

	bool isBstr = (value.vt == VT_BSTR);
	if (isBstr)	{
			valueBytes = SysStringLen(value.bstrVal) * BSTRSIZE;
			totalBytes = valueBytes + BSTRSIZE; // Add for NULL delimiter
			pSource = value.bstrVal;
	} else {
			totalBytes = valueBytes = GetNumTypeSizeFromFormat(value.vt); // throws error if not a number
			pSource = &(value.cVal);
	}

	if (maxByteSize == 0 || totalBytes <= maxByteSize) 
	{
		DWORD bounds[] = {totalBytes};
		COleSafeArray saChunk;
		saChunk.Create(VT_UI1, 1, bounds);
		BYTE *destData;
		saChunk.AccessData((void **)&destData);
		memcpy(destData, pSource, valueBytes);
		if (isBstr) {
			destData += valueBytes;
			memset(destData, 0, BSTRSIZE);	// Add NULL delimiter
		}
		saChunk.UnaccessData();

		variant_t tempVariant(saChunk.Detach(), false);
		pDBObj.AppendChunk(name, tempVariant);

		bytesWritten = totalBytes;
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// For string values, we will add a NUL delimiter after each value
void DBMemoryBuilder::WriteBinaryArray(ADOLoggingObject &pDBObj, _bstr_t &name, variant_t &value, unsigned int maxByteSize, unsigned int &bytesWritten, bool &stop)
{
    unsigned int elementByteSize = 0;
    unsigned int numElements = GetNumArrayElements(value, &elementByteSize);
	bytesWritten = 0;

	if (!numElements)
		return;

    COleSafeArray sa(value);
	if ((value.vt & VT_TYPEMASK) == VT_BSTR) {
		BSTR* sourceData;
		sa.AccessData((void**)(&sourceData));
		for (unsigned int i = 0; i < numElements; i++) 
		{
			_bstr_t b = sourceData[i];
			unsigned int size = b.length() + BSTRSIZE;
			if (maxByteSize > 0) {
				if (size > maxByteSize)
					break;
				else
					maxByteSize -= size;
			}
			bytesWritten += size;
			pDBObj.AppendChunk(name, b);
			WriteBinaryZeros(pDBObj, name, BSTRSIZE); // Add delimeter
		}
		sa.UnaccessData();
		return;
	} else { 
		// Array is numeric
		BYTE *sourceData;
		sa.AccessData((void**)(&sourceData));
		unsigned int totalBytes = elementByteSize * numElements;

		// Determine max number of elements
		if ((maxByteSize > 0) && (totalBytes > maxByteSize))
			totalBytes = (maxByteSize / elementByteSize) * elementByteSize;
		
		unsigned int chunkSize =  min(totalBytes, APPEND_CHUNK_BLOCK_SIZE);
		unsigned int chunkTotal = 0;
    
		while (chunkTotal < totalBytes)
		{
			int nextChunkSize = (totalBytes - chunkTotal < chunkSize ) ? totalBytes - chunkTotal : chunkSize;
			DWORD bounds[] = {nextChunkSize};
			COleSafeArray saChunk;
			saChunk.Create(VT_UI1, 1, bounds);

			void *destData;
			saChunk.AccessData(&destData);
			memcpy(destData, (void *)sourceData, nextChunkSize);
			chunkTotal += nextChunkSize;
			sourceData += nextChunkSize;
			saChunk.UnaccessData();

			variant_t value(saChunk.Detach(), false);
			pDBObj.AppendChunk(name, value);
		}
		sa.UnaccessData();
		bytesWritten += totalBytes;
	}
	#ifdef _DEBUG
	DebugPrintf("SetColumn:        %s=...\n", (char*)name);
	#endif
}

void DBMemoryBuilder::WriteBinaryZeros(ADOLoggingObject &pDBObj, _bstr_t &name, unsigned int bytes)
{
	DWORD bounds[] = {bytes};
	COleSafeArray saChunk;
	saChunk.Create(VT_UI1, 1, bounds);
	void *destData;
	saChunk.AccessData(&destData);
	memset(destData, 0, BSTRSIZE);
	saChunk.UnaccessData();
	variant_t tempVariant(saChunk.Detach(), false);
	pDBObj.AppendChunk(name, tempVariant);
}

void DBMemoryBuilder::EscapeCRLF(_bstr_t &stringValue)
{
	int origSize = stringValue.length();
	BSTR origValue = stringValue.GetBSTR();
	int itemsFound = 0;
	for (int i=0; i < origSize; i++)
	{
		if (origValue[i] == '\\' || origValue[i] == '\r' || origValue[i] == '\n')
			itemsFound++;
	}
	if (itemsFound)
	{
		CStringW newValue;
		newValue.Preallocate(origSize + itemsFound);
		wchar_t *pOrig = stringValue.GetBSTR();
		wchar_t *pNew = newValue.GetBuffer();
		for (int i=0; i <= origSize; i++)
		{
			if (*pOrig == '\r') {
				*pNew = '\\'; 
				pNew++;
				*pNew = 'r'; 
				pNew++;
			} else if (*pOrig == '\n') {
				*pNew = '\\'; 
				pNew++;
				*pNew = 'n'; 
				pNew++;
			} else if (*pOrig == '\\') {
				*pNew = '\\'; 
				pNew++;
				*pNew = '\\'; 
				pNew++;
			} else {
				*pNew = *pOrig; 
				pNew++; 
			}
			pOrig++;
		}
		newValue.ReleaseBuffer();
		_COM_ASSERT(newValue.GetLength() == origSize + itemsFound);
		stringValue = newValue;
	}
}

int DBMemoryBuilder::GetNumTypeSizeFromFormat(VARTYPE numType)
{
	int size;
	if (numType == VT_R8 || numType == VT_I8 || numType == VT_UI8)
		size = 8;
	else if (numType == VT_R4 || numType == VT_I4 || numType == VT_UI4)
		size = 4;
	else if (numType == VT_I2 || numType == VT_UI2 || numType == VT_BOOL)
		size = 2;
	else if (numType == VT_I1 || numType == VT_UI1)
		size = 1;
	else
		RAISE_ERROR_WITH_DESC(TSDBLogErr_InvalidFormatValue, TS::TS_Err_UnRecognizedValue, "");

	return size;
}

// Throws error if invalid numeric type is passed
VARTYPE DBMemoryBuilder::GetNumTypeFromFormat(const _bstr_t &format, VARTYPE defaultVarType)
{
	VARTYPE numVarType = defaultVarType;

	if (format.length() > 0) 
	{
		if (!_wcsicmp(format, L"R8"))
			numVarType = VT_R8;
		else if (!_wcsicmp(format, L"R4"))
			numVarType = VT_R4;
		else if (!_wcsicmp(format, L"I8"))
			numVarType = VT_I8;
		else if (!_wcsicmp(format, L"I4"))
			numVarType = VT_I4;
		else if (!_wcsicmp(format, L"I2"))
			numVarType = VT_I2;
		else if (!_wcsicmp(format, L"I1"))
			numVarType = VT_I1;
		else if (!_wcsicmp(format, L"UI8"))
			numVarType = VT_UI8;
		else if (!_wcsicmp(format, L"UI4"))
			numVarType = VT_UI4;
		else if (!_wcsicmp(format, L"UI2"))
			numVarType = VT_UI2;
		else if (!_wcsicmp(format, L"UI1"))
			numVarType = VT_UI1;
		else
			RAISE_ERROR_WITH_DESC(TSDBLogErr_InvalidFormatValue, TS::TS_Err_UnRecognizedValue, "");
	}
	return numVarType;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Initialize column's type and key details
CDbColumn::CDbColumn(TS::IEnginePtr& pEngine, TS::PropertyObjectPtr& pDBColumn, Statement* pStmt)
{
    double tempVal;
    long tempLong; // used in the GET_ENUM macro
	_variant_t avtString;
	VARIANT    avString;
	SAFEARRAY *asfString;
	TSUTIL::SafeArray<BSTR, VT_BSTR> stringSafeArray;

	// Process column settings
    try {
		m_name = pDBColumn->GetValString("Name",0);
		m_pStatement = pStmt;
		GET_ENUM(ColDataType,"Type",m_type);

		switch (m_type) {
		case tsDBOptionsColumnType_SmallInt:
			m_AdoType = adSmallInt;
			break;
		case tsDBOptionsColumnType_Integer:
			m_AdoType = adInteger;
			break;
		case tsDBOptionsColumnType_Float:
			m_AdoType = adSingle;
			break;
		case tsDBOptionsColumnType_DoublePrecision:
			m_AdoType = adDouble;
			break;

		case tsDBOptionsColumnType_Bstr:
			m_AdoType = adBSTR;
			break;

		case tsDBOptionsColumnType_VarChar:
			m_AdoType = adVarChar;
			break;

		case tsDBOptionsColumnType_Boolean:
			m_AdoType = adBoolean;
			break;
		case tsDBOptionsColumnType_Binary:
			m_AdoType = adVarBinary;
			break;

		case tsDBOptionsColumnType_DateTime:
			{
				CTSDatalink * dataLink = (m_pStatement->GetDatalink());
				if (dataLink)
				{
					// Fix for when Access swaps month and day in adDBTimeStamp column for parameterized insert statements
					// logging in "European" local that uses dd/mm/yyyy format
					m_AdoType = dataLink->GetDateDataType();
				}
				else
				{
					#if _DEBUG
					assert(0);
					#endif
					m_AdoType = adDBTimeStamp;
				}
				break;
			}
		case tsDBOptionsColumnType_GUID:
			m_AdoType = adGUID;
			break;
		}

		GET_ENUM(ParameterDirectionEnum, "Direction", m_direction);

		m_expressionExpr = pEngine->NewExpression();
		m_expressionExpr->Text = pDBColumn->GetValString("Expression",0);
   		m_expressionExpr->Tokenize(0, 0);

		tempVal = pDBColumn->GetValNumber("Size",0);
		m_size = (unsigned int)tempVal;
		m_format = pDBColumn->GetValString("Format",0);

		m_preconditionExpr = pEngine->NewExpression();
		m_preconditionExpr->Text = pDBColumn->GetValString("Precondition",0);
   		m_preconditionExpr->Tokenize(0, 0);

		avtString = pDBColumn->GetValVariant("ExpectedProperties",0);
 		avString = avtString.Detach();
		asfString = V_ARRAY(&avString);
		stringSafeArray.Set(asfString, true);
		stringSafeArray.GetVector(m_expectedProperties);

		GET_ENUM(PrimaryKeyTypes,"PrimaryKeyType",m_primaryKeyType);
		m_isPrimaryKey = (m_primaryKeyType != tsDBOptionsPrimaryKeyType_NotKey);
		m_primaryKeyStatement = pDBColumn->GetValString("PrimaryKeyStatementText",0);
		m_foreignKeyName = pDBColumn->GetValString("ForeignKeyStatementName",0);
		m_isForeignKey = (m_foreignKeyName != _bstr_t(""));
		m_isStepRecursiveKey = false;
		m_pForeignKeyStmt = NULL;
		if (m_isPrimaryKey) 
		{
			switch (m_primaryKeyType) 
			{
			case tsDBOptionsPrimaryKeyType_AutoGenerated:
			case tsDBOptionsPrimaryKeyType_StoreGuid:
			case tsDBOptionsPrimaryKeyType_UseExpression:
				// nothing to do
				break;

			case tsDBOptionsPrimaryKeyType_GetValueFromStatement:
			case tsDBOptionsPrimaryKeyType_GetValueFromStoredProcOutputValue:
			case tsDBOptionsPrimaryKeyType_GetValueFromStoredProcReturnValue:
				//set up our prepared statement
				CREATEADOINSTANCE(m_pPrimaryKeyCommand, Command, "Command"); 
				m_pPrimaryKeyCommand->ActiveConnection = (m_pStatement->GetDatalink())->m_spAdoCon;
				
				m_pPrimaryKeyCommand->CommandText      = m_primaryKeyStatement;
				m_pPrimaryKeyCommand->CommandType      = adCmdText;
				m_pPrimaryKeyCommand->put_Prepared(true);

				if ( m_primaryKeyType != tsDBOptionsPrimaryKeyType_GetValueFromStatement)
				{
					//assume that the output param is the same datatype as the key column
					_ParameterPtr pprmCol = m_pPrimaryKeyCommand->CreateParameter(GetName(), GetAdoType(), 
						(m_primaryKeyType == tsDBOptionsPrimaryKeyType_GetValueFromStoredProcReturnValue)? adParamReturnValue : adParamOutput,
						m_size, vtMissing);
					m_pPrimaryKeyCommand->Parameters->Append(pprmCol);
				}
				break;
			}
		}
	}
	catch (TSDBLogException &e) {
		_bstr_t description = GetTSDBLogResString(pEngine, "Err_ErrorOccurredForColumn");
		description += m_name;
		description += "\n";
		description += e.mDescription;
		e.mDescription = description;
		throw e;
	}		
	catch (_com_error &ce) {
		_bstr_t description = GetTSDBLogResString(pEngine, "Err_ErrorOccurredForColumn");
		description += m_name;
		description += "\n";
		description += ce.Description();
		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToInitStatement, ce.Error(), description);
	}
	catch (...) {
		RAISE_ERROR(TSDBLogErr_FailedToInitStatement);
	}		
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Cleanup column object
CDbColumn::~CDbColumn()
{
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Upates pointer to statement that the foreign key column references
void CDbColumn::InitForeignKeyInformation(void)
{
	CTSDatalink* pDatalink = m_pStatement->GetDatalink();
	Statement *pStatement = NULL;
	POSITION pos;

	pos = pDatalink->m_UUTStatements.GetHeadPosition();
	while (pos) {
		pStatement = pDatalink->m_UUTStatements.GetNext(pos);
		if ((pStatement)->HasSameName(m_foreignKeyName))
			goto Leave;
	}

	pos = pDatalink->m_stepStatements.GetHeadPosition();
	while (pos) {
		pStatement = pDatalink->m_stepStatements.GetNext(pos);
		if ((pStatement)->HasSameName(m_foreignKeyName))
			goto Leave;
	}

	pos = pDatalink->m_propStatements.GetHeadPosition();
	while (pos) {
		pStatement = pDatalink->m_propStatements.GetNext(pos);
		if ((pStatement)->HasSameName(m_foreignKeyName))
			goto Leave;
	}

	pStatement = NULL;
Leave:
	m_pForeignKeyStmt = pStatement;
	m_isStepRecursiveKey = ((pStatement == m_pStatement) && (pStatement->GetResultType() == tsDBOptionsResultType_Step));
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Return whether the column expected properties exist and preconditions expression is true
bool CDbColumn::ValidatePreConditions(LoggingInfo &loggingInfo)
{
	TS::PropertyObjectPtr pProperty = loggingInfo.GetContext()->AsPropertyObject();
	try {
		if (m_expectedProperties.size() > 0)
		{
			VARIANT_BOOL exists;
			for (unsigned int i=0; (i < m_expectedProperties.size()); i++)
			{
				exists = pProperty->Exists(m_expectedProperties[i], 0);
				if (!exists)
					return false;
			}
		}
	}
	catch( _com_error &ce) {
		//_bstr_t description = ce.Description();
		RAISE_ERROR_WITH_DESC(TSDBLogErr_ErrUnknown, ce.Error(), ce.Description());
	}

	// if precondition exists, evaluate it
	if (m_preconditionExpr->NumTokens > 0) {
		TS::PropertyObjectPtr pEvalResult;
		try {
			pEvalResult = m_preconditionExpr->Evaluate(pProperty, 0);
		}
		catch (_com_error &ce) {
			RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedPrecondition, ce.Error(), ce.Description());
		}
		if (!pEvalResult->GetValBoolean("",0))
			return false;
	}

	return true;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Evaluate the column expression if preconditions are met
void CDbColumn::UpdateOutputValue(LoggingInfo &loggingInfo, _variant_t vValue)
{
    if (!ValidatePreConditions(loggingInfo))
		return;

	m_expressionExpr->Evaluate(loggingInfo.GetContext()->AsPropertyObject(), 0)->SetValVariant("", 0, vValue);
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// For parameterized command object, set the column value based on its type.  
// Primary keys may be generated, and foreign key values must be determined
void CDbColumn::SetColumn(LoggingInfo &loggingInfo, ADOLoggingObject &pDBObj, int depth)
{
	TS::SequenceContextPtr pContext = loggingInfo.GetContext();
	double tempVal;
    _bstr_t tempStr;
	_variant_t vtNull;
	vtNull.ChangeType(VT_NULL);

    try {
        if ( ValidatePreConditions(loggingInfo) == false )
		{
			if (pDBObj.IsCommand())
				pDBObj.Assign(m_name, vtNull);
			return; 
		}

		if (m_isPrimaryKey && (m_primaryKeyType != tsDBOptionsPrimaryKeyType_UseExpression))
		{
            GUID guid;
            RPC_STATUS rStatus;
            unsigned char * string;

            switch (m_primaryKeyType) {
            case tsDBOptionsPrimaryKeyType_AutoGenerated:
                // nothing to do
                break;

            case tsDBOptionsPrimaryKeyType_StoreGuid:
                rStatus =  ::UuidCreate(&guid);
                ::UuidToString(&guid, &string);
				
				try {
					char szGuid[64];
					sprintf(szGuid, "{%s}", string);
					#ifdef _DEBUG
						DebugPrintf("SetColumn:        %s={%s}\n", (char*)m_name, string);
					#endif
					pDBObj.Assign((_bstr_t)m_name, (variant_t)szGuid);
                }
                catch( _com_error &e){
                    ::RpcStringFree(&string);
                    throw e;
                }
                ::RpcStringFree(&string);

                break;

            case tsDBOptionsPrimaryKeyType_GetValueFromStatement:
			case tsDBOptionsPrimaryKeyType_GetValueFromStoredProcOutputValue:
			case tsDBOptionsPrimaryKeyType_GetValueFromStoredProcReturnValue:
                SetPrimaryKey(loggingInfo, pDBObj);
                break;
            }
        }
        else if (m_isForeignKey) {
            SetForeignKey(loggingInfo, pDBObj, depth);
        }
        else { // just a regular column
            TS::PropertyObjectPtr pValue;
            pValue = m_expressionExpr->Evaluate(pContext->AsPropertyObject(), 0);

            switch (m_type) {
            case tsDBOptionsColumnType_SmallInt:
				tempVal = pValue->GetValNumber("", COERCE_FROM_OPTIONS);
				pDBObj.Assign(m_name, (variant_t)(short)tempVal);
				#ifdef _DEBUG
					DebugPrintf("SetColumn:        %s=%d\n", (char*)m_name, (int)tempVal);
				#endif
                break;
            case tsDBOptionsColumnType_Integer:
				tempVal = pValue->GetValNumber("", COERCE_FROM_OPTIONS);
				pDBObj.Assign(m_name, (variant_t)(long)tempVal);
				#ifdef _DEBUG
					DebugPrintf("SetColumn:        %s=%d\n", (char*)m_name, (int)tempVal);
				#endif
                break;
            case tsDBOptionsColumnType_Float:
				tempVal = pValue->GetValNumber("", COERCE_FROM_OPTIONS);
				pDBObj.Assign(m_name, (variant_t)(float)tempVal);
				#ifdef _DEBUG
					DebugPrintf("SetColumn:        %s=%fl\n", (char*)m_name, (double)tempVal);
				#endif
                break;
            case tsDBOptionsColumnType_DoublePrecision:
				tempVal = pValue->GetValNumber("", COERCE_FROM_OPTIONS);
				pDBObj.Assign(m_name, (variant_t)(double)tempVal);
				#ifdef _DEBUG
					DebugPrintf("SetColumn:        %s=%fl\n", (char*)m_name, (double)tempVal);
				#endif
                break;

            case tsDBOptionsColumnType_Boolean:
				{
				VARIANT_BOOL value = pValue->GetValBoolean("", COERCE_FROM_OPTIONS);
				pDBObj.Assign(m_name, (variant_t)value);
				#ifdef _DEBUG
					DebugPrintf("SetColumn:        %s='%s'\n", (char*)m_name, (char*)(pValue->GetValString("", COERCE_FROM_OPTIONS)));
				#endif
				}
				break;

            case tsDBOptionsColumnType_DateTime:
			{
                tempStr = pValue->GetValString("", 0);
			    _variant_t  timeVariant;

				GetDateTimeValue( tempStr, timeVariant);
				if (timeVariant.vt == VT_DATE) {
					pDBObj.Assign(m_name, timeVariant);
					#ifdef _DEBUG
						DebugPrintf("SetColumn:        %s=%'s'\n", (char*)m_name, (char*)tempStr);
					#endif
				} else if (pDBObj.IsCommand()) {
					pDBObj.Assign(m_name, vtNull);
					#ifdef _DEBUG
						DebugPrintf("SetColumn:        %s failed to properly convert date% ('s')\n", (char*)m_name, (char*)tempStr);
					#endif
				}
				break;
			}
            case tsDBOptionsColumnType_Bstr:
            case tsDBOptionsColumnType_VarChar:
			case tsDBOptionsColumnType_GUID:
				{
				DBMemoryBuilder memoryBuilder(m_type);
				memoryBuilder.Append(pValue, m_format);
				memoryBuilder.Write(pDBObj, m_name, (m_type == tsDBOptionsColumnType_Bstr) ? (m_size * BSTRSIZE) : m_size);
				}
				break;

            case tsDBOptionsColumnType_Binary:
				{
				DBMemoryBuilder memoryBuilder(m_type);
				memoryBuilder.Append(pValue, m_format);
				memoryBuilder.Write(pDBObj, m_name, (m_type == tsDBOptionsColumnType_Bstr) ? (m_size * BSTRSIZE) : m_size);
				}
				break;

			default:
				RAISE_ERROR_WITH_DESC(TSDBLogErr_InvalidColumnType, TS::TS_Err_UnRecognizedValue, "");
				break;
            }
        }
    }
	catch (TSDBLogException &e) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForColumn");
		description += m_name;
		description += "\n";
		if (loggingInfo.GetPropPath().length())	{
			description += GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForProperty");
			description += loggingInfo.GetPropPath();
			description += "\n";
		}
		description += e.mDescription;

		e.mDescription = description;
		throw e;
	}		
    catch( _com_error &ce) {
		// ce.Error() = 0x800A0CC1 if name is not found in recordset
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForColumn");
		description += m_name;
		description += "\n";
		if (loggingInfo.GetPropPath().length())	{
			description += GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForProperty");
			description += loggingInfo.GetPropPath();
			description += "\n";
		}

		try {
			if (pDBObj.IsRecordSet()) {
				long fieldStatus = pDBObj.GetStatus(m_name);
				if (fieldStatus != adFieldOK) {
					description += "\nColumn Status: ";
					description += GetColumnStatus(fieldStatus);
				}
			}
		}
		catch( _com_error &) {}

		description += ::GetTextForAdoErrors(m_pStatement->GetDatalink()->m_spAdoCon, ce.Description());

		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToSetColumn, ce.Error(), description);
	}
	catch (...) {
		_bstr_t description = GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForColumn");
		description += m_name;
		description += "\n";
		if (loggingInfo.GetPropPath().length())	{
			description += GetTSDBLogResString(pContext->Engine, "Err_ErrorOccurredForProperty");
			description += loggingInfo.GetPropPath(); 
			description += "\n";
		}

		RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedToSetColumn, TSDBLogErr_ErrUnknown, description);
	}		

}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
_bstr_t CDbColumn::GetColumnStatus(long status)
{
	_bstr_t fieldStatusText;

	// Get column has status
	switch (status) {

	case adFieldAlreadyExists:
		fieldStatusText = "adFieldAlreadyExists(26) - Indicates that the specified field already exists.";
		break;
	case adFieldBadStatus:
		fieldStatusText = "adFieldBadStatus(12) - Indicates that an invalid status value was sent from ADO to the OLE DB provider. Possible causes include an OLE DB 1.0 or 1.1 provider, or an improper combination of Value and Status.";
		break;
	case adFieldCannotComplete:
		fieldStatusText = "adFieldCannotComplete(20) - Indicates that the server of the URL specified by Source could not complete the operation.";
		break;
	case adFieldCannotDeleteSource:
		fieldStatusText = "adFieldCannotDeleteSource(23) - Indicates that during a move operation, a tree or subtree was moved to a new location, but the source could not be deleted.";
		break;
	case adFieldCantConvertValue:
		fieldStatusText = "adFieldCantConvertValue(2) - Indicates that the field cannot be retrieved or stored without loss of data.";
		break;
	case adFieldCantCreate:
		fieldStatusText = "adFieldCantCreate(7) - Indicates that the field could not be added because the provider exceeded a limitation (such as the number of fields allowed).";
		break;
	case adFieldDataOverflow:
		fieldStatusText = "adFieldDataOverflow(6) - Indicates that the data returned from the provider overflowed the data type of the field.";
		break;
	case adFieldDefault:
		fieldStatusText = "adFieldDefault(13) - Indicates that the default value for the field was used when setting data.";
		break;
	case adFieldDoesNotExist:
		fieldStatusText = "adFieldDoesNotExist(16) - Indicates that the field specified does not exist.";
		break;
	case adFieldIgnore:
		fieldStatusText = "adFieldIgnore(15) - Indicates that this field was skipped when setting data values in the source. No value was set by the provider.";
		break;
	case adFieldIntegrityViolation:
		fieldStatusText = "adFieldIntegrityViolation(10) - Indicates that the field cannot be modified because it is a calculated or derived entity.";
		break;
	case adFieldInvalidURL:
		fieldStatusText = "adFieldInvalidURL(17) - Indicates that the data source URL contains invalid characters.";
		break;
	case adFieldIsNull:
		fieldStatusText = "adFieldIsNull(3) - Indicates that the provider returned a null value.";
		break;
	case adFieldOK:
		fieldStatusText = "adFieldOK(0) - Default. Indicates that the field was successfully added or deleted.";
		break;
	case adFieldOutOfSpace:
		fieldStatusText = "adFieldOutOfSpace(22) - Indicates that the provider is unable to obtain enough storage space to complete a move or copy operation.";
		break;
	case adFieldPendingChange:
		fieldStatusText = "adFieldPendingChange(0x40000) - Indicates either that the field has been deleted and then re-added, perhaps with a different data type, or that the value of the field which previously had a status of adFieldOK has changed. The final form of the field will modify the Fields collection after the Update method is called.";
		break;
	case adFieldPendingDelete:
		fieldStatusText = "adFieldPendingDelete(0x20000) - Indicates that the Delete operation caused the status to be set. The field has been marked for deletion from the Fields collection after the Update method is called. ";
		break;
	case adFieldPendingInsert:
		fieldStatusText = "adFieldPendingInsert(0x10000) - Indicates that the Append operation caused the status to be set. The Field has been marked to be added to the Fields collection after the Update method is called.";
		break;
	case adFieldPendingUnknown:
		fieldStatusText = "adFieldPendingUnknown(0x80000) - Indicates that the provider cannot determine what operation caused field status to be set.";
		break;
	case adFieldPendingUnknownDelete:
		fieldStatusText = "adFieldPendingUnknownDelete(0x100000) - Indicates that the provider cannot determine what operation caused field status to be set, and that the field will be deleted from the Fields collection after the Update method is called.";
		break;
	case adFieldPermissionDenied:
		fieldStatusText = "adFieldPermissionDenied(9) - Indicates that the field cannot be modified because it is defined as read-only.";
		break;
	case adFieldReadOnly:
		fieldStatusText = "adFieldReadOnly(24) - Indicates that the field in the data source is defined as read-only.";
		break;
	case adFieldResourceExists:
		fieldStatusText = "adFieldResourceExists(19) - Indicates that the provider was unable to perform the operation because an object already exists at the destination URL and it is not able to overwrite the object.";
		break;
	case adFieldResourceLocked:
		fieldStatusText = "adFieldResourceLocked(18) - Indicates that the provider was unable to perform the operation because the data source is locked by one or more other application or process.";
		break;
	case adFieldResourceOutOfScope:
		fieldStatusText = "adFieldResourceOutOfScope(25) - Indicates that a source or destination URL is outside the scope of the current record.";
		break;
	case adFieldSchemaViolation:
		fieldStatusText = "adFieldSchemaViolation(11) - Indicates that the value violated the data source schema constraint for the field.";
		break;
	case adFieldSignMismatch:
		fieldStatusText = "adFieldSignMismatch(5) - Indicates that data value returned by the provider was signed but the data type of the ADO field value was unsigned.";
		break;
	case adFieldTruncated:
		fieldStatusText = "adFieldTruncated(4) - Indicates that variable-length data was truncated when reading from the data source.";
		break;
	case adFieldUnavailable:
		fieldStatusText = "adFieldUnavailable(8) - Indicates that the provider could not determine the value when reading from the data source. For example, the row was just created, the default value for the column was not available, and a new value had not yet been specified.";
		break;
	case adFieldVolumeNotFound:
		fieldStatusText = "adFieldVolumeNotFound(21) - Indicates that the provider is unable to locate the storage volume indicated by the URL.";
		break;
	}

	return fieldStatusText;
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Get the last foreign key value that the column references
void CDbColumn::GetForeignKeyValue(LoggingInfo &loggingInfo, _variant_t &KeyVal, int depth)
{
	Statement *pStatement = GetForeignKeyStmt();
	if (pStatement != NULL) 
	{
		if (depth < 0)
			KeyVal = pStatement->GetLastKeyValue(loggingInfo); 
		else
			KeyVal = pStatement->GetKeyValueForDepth(loggingInfo, depth);
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// For parameterized command objects, set the value of the foreign key
void CDbColumn::SetForeignKey(LoggingInfo &loggingInfo, ADOLoggingObject &pDBObj, int depth)
{
	_variant_t keyVal;
	
	GetForeignKeyValue(loggingInfo, keyVal, depth);    
	#ifdef _DEBUG
		DebugPrintf("SetColumn:        %s=%s\n", (char*)m_name, (char*)((keyVal.vt==1)?"NULL":(_bstr_t)keyVal));
	#endif
	pDBObj.Assign(m_name, keyVal);
}


//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Gets the primary key returned when executing a stored procedure or SQL query
void CDbColumn::GetPrimaryKeyValue(LoggingInfo &loggingInfo, _variant_t &KeyVal)
{
    _RecordsetPtr   pTempRS;

	if (m_primaryKeyType == tsDBOptionsPrimaryKeyType_GetValueFromStatement  )
	{
		CREATEADOINSTANCE(pTempRS, Recordset, "Recordset"); 
		pTempRS->PutRefActiveConnection( (m_pStatement->GetDatalink())->m_spAdoCon );
		pTempRS->PutRefSource(m_pPrimaryKeyCommand);
		pTempRS->CursorLocation = adUseClient;
		
		HRESULT hr = pTempRS->Open(vtMissing,vtMissing,adOpenStatic,adLockOptimistic,-1);      
		if (hr != S_OK) 
			RAISE_ERROR_WITH_DESC(TSDBLogErr_FailedGetPrimaryKey, hr, "");

		pTempRS->MoveFirst();
		long numFields = pTempRS->Fields->Count;
		if (numFields > 0) 
			KeyVal = RsITEM(pTempRS, 0L);
		else
			KeyVal = vtMissing;
    
		pTempRS->Close();
	}
	else  //output param or return value param of stored procedure
	{
		m_pPrimaryKeyCommand->Execute(NULL, NULL, adCmdStoredProc);
		KeyVal = CmdITEM(m_pPrimaryKeyCommand, GetName());
	}
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// For parameterized command object, set the primary key column value
void CDbColumn::SetPrimaryKey(LoggingInfo &loggingInfo, ADOLoggingObject &pDBObj)
{
	_variant_t keyVal;

	GetPrimaryKeyValue(loggingInfo, keyVal);
	#ifdef _DEBUG
		DebugPrintf("SetColumn:        %s=%s\n", (char*)m_name, (char*)(_bstr_t)keyVal);
	#endif
	pDBObj.Assign(m_name, keyVal);
}

//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
// Convert a string value to a variant date/time value
void CDbColumn::GetDateTimeValue( _bstr_t str, _variant_t &timeVariant)
{
    LCID        lcid = GetUserDefaultLCID();
    int status = FormatToVariant(lcid, (char*)str, (char*)m_format, timeVariant);
}

