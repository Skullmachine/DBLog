#include "tsutilcpp.h"
#include "stdafx.h"
#include "format.h"

// lifted verbatim from cvi_db.c, comments and all, with the following changes:
// 1. replaced DBStrICmp with MS MBstricmp function
// 2. needed to cast malloc calls with (char*)
// 3. changed CA_Variant calls to _variant_t assignments
// 4. changed parameter type from VARIANT* to _variant_t

#define INIT_TIME_DATA {																\
{"AM", "PM"},                                           \
{"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"},      \
{"Sunday", "Monday", "Tuesday", "Wednesday",            \
 "Thursday", "Friday", "Saturday"},                     \
{"Jan", "Feb", "Mar", "Apr", "May", "Jun",              \
 "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"},             \
{"January", "February", "March", "April",               \
 "May", "June", "July", "August",                       \
 "September", "October", "November", "December"},       \
}

struct time_env {
  char *AmPm[2];
  char *AbrevDays[7];
  char *FullDays[7];
  char *AbrevMons[12];
  char *FullMons[12];
};

#define MBstricmp(s1, s2) _mbsicmp((unsigned char *) s1, (unsigned char *) s2)
#define MBstrstr(s1, s2) _mbsstr((unsigned char *) s1, (unsigned char *) s2)
#define MBstrlen(s1) _mbslen((unsigned char *) s1)
#define MBstrcpy(s1, s2) _mbscpy((unsigned char *) s1, (unsigned char *) s2)

/* to support crappy InterSolv format shit */
static const short leapMonthLenTab[] = {31,29,31,30,31,30,31,31,30,31,30,31};
static const short regMonthLenTab[]  = {31,28,31,30,31,30,31,31,30,31,30,31};
static struct time_env sTimeEnv = INIT_TIME_DATA;

#if _DEBUG
int __cdecl DebugPrintf(const char *buf, ...);
#endif

static int GetMonthNumFromAbbrevName(
	LCID lcid,
	int forceUS,
	char* buffer)
{
	char tempBuf[30];
	int i;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */
	
	if (forceUS) {
		for (i = 0; i < 12; i++) {
			if (MBstricmp(buffer, sTimeEnv.AbrevMons[i]) == 0)
				return i;
		}
		return 0;
	} else {
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME1, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 0;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME2, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 1;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME3, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 2;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME4, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 3;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME5, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 4;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME6, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 5;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME7, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 6;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME8, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 7;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME9, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 8;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME10, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 9;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME11, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 10;
		GetLocaleInfo(lcid, LOCALE_SABBREVMONTHNAME12, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 11;
		return 0;
	}
}

static int GetMonthNumFromFullName(
	LCID lcid,
	int forceUS,
	char* buffer)
{
	char tempBuf[30];
	int i;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */
	
	if (forceUS) {
		for (i = 0; i < 12; i++) {
			if (MBstricmp(buffer, sTimeEnv.FullMons[i]) == 0)
				return i;
		}
		return 0;
	} else {
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME1, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 0;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME2, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 1;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME3, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 2;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME4, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 3;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME5, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 4;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME6, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 5;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME7, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 6;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME8, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 7;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME9, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 8;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME10, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 9;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME11, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 10;
		GetLocaleInfo(lcid, LOCALE_SMONTHNAME12, tempBuf, 30);
		if (MBstricmp(buffer, tempBuf) == 0) return 11;
		return 0;
	}
}

// Returns TRUE if the formated date contains PM or the localized symbol for PM
static int ExtractAMPMFromFluff(
	LCID lcid,
	int forceUS,
	char* fmt,
	unsigned int fmtLen,
	unsigned int* fi,
	char *fluff,
	unsigned int fluffLen,
	unsigned int* fdi)
{
	int retVal = FALSE;
	int count = 0;
	char amStrLocal[30];
	char pmStrLocal[30];
	char *amStrUS = sTimeEnv.AmPm[0];
	char *pmStrUS = sTimeEnv.AmPm[1];
	bool matchWasLocalFormat = false;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */

	/* skip over the am/pm stuff in the format */
	while ( ((*fi) < fmtLen) && (fmt[(*fi)] != 'a') && (fmt[(*fi)] != 'A') )
		(*fi)++;
	while ( ((*fi) < fmtLen) && ( (fmt[(*fi)] == 'a') || (fmt[(*fi)] == 'A') ||
	                            (fmt[(*fi)] == 'p') || (fmt[(*fi)] == 'P') ||
	                            (fmt[(*fi)] == 'm') || (fmt[(*fi)] == 'M') ||
	                            (fmt[(*fi)] == '/') )) {
		(*fi)++;
		count++;
	}

	if (forceUS) {
		MBstrcpy(amStrLocal, amStrUS);
		MBstrcpy(pmStrLocal, pmStrUS);
	} else {
		GetLocaleInfo(lcid, LOCALE_S1159, amStrLocal, 30);
		GetLocaleInfo(lcid, LOCALE_S2359, pmStrLocal, 30);
	}
	/* allow some slop (and only compare the first character when found) */
	while ( ((*fdi) < fluffLen) && 
		    (fluff[(*fdi)] != tolower(amStrLocal[0])) && (fluff[(*fdi)] != toupper(amStrLocal[0])) &&
	        (fluff[(*fdi)] != tolower(pmStrLocal[0])) && (fluff[(*fdi)] != toupper(pmStrLocal[0])) &&
		    (fluff[(*fdi)] != tolower(amStrUS[0])) && (fluff[(*fdi)] != toupper(amStrUS[0])) &&
	        (fluff[(*fdi)] != tolower(pmStrUS[0])) && (fluff[(*fdi)] != toupper(pmStrUS[0])) )
		(*fdi)++;
	if ( (fluff[(*fdi)] == tolower(pmStrLocal[0])) || (fluff[(*fdi)] == toupper(pmStrLocal[0])) )
	{
		retVal = TRUE;
		matchWasLocalFormat = true;
	}
	else if ( (fluff[(*fdi)] == tolower(pmStrUS[0])) || (fluff[(*fdi)] == toupper(pmStrUS[0])) )
	{
		retVal = TRUE;
		matchWasLocalFormat = false;
	}
	
	// If the format was a/p or A/P, then assume that formated date contains only a single character from the symbol
	if (count == 3) { 
		(*fdi)++;			
	} else {
		// Assume that formated date contained entire US or local symbol, so move index accordingly
		if (retVal) {
			if (matchWasLocalFormat)
				(*fdi) += MBstrlen(pmStrLocal);
			else
				(*fdi) += MBstrlen(pmStrUS);
		} else { 
			if (matchWasLocalFormat)
				(*fdi) += MBstrlen(amStrLocal);
			else
				(*fdi) += MBstrlen(amStrUS);
		}
	}
	return retVal;
}

static int ExtractNumberFromFluff(
	char *fluff,
	unsigned int fluffLen,
	unsigned int* fdi)
{
	int tempInt = 0;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */
	
	/* find the next numeric char */
	while( ((*fdi) < fluffLen) && (!isdigit(fluff[*fdi])) )(*fdi)++;
	/* get the number */
	while( ((*fdi) < fluffLen) && (isdigit(fluff[*fdi])) ) {
		tempInt = (tempInt * 10) + (int)(fluff[*fdi] - '0');
		(*fdi)++;
	}
	return tempInt;
}

static int ExtractMonthFromFluff(
	LCID lcid,
	int forceUS,
	char* fmt,
	unsigned int fmtLen,
	unsigned int* fi,
	char *fluff,
	unsigned int fluffLen,
	unsigned int* fdi)
{
	int i;
	int count;
	char buffer[100];
	int tempInt;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */
	count = 0;
	while (((*fi) < fmtLen) && ((fmt[(*fi)] == 'm') || (fmt[(*fi)] == 'M')) ) {
		count++;
		(*fi)++;
	}
	switch (count) {
		case 1:
			tempInt = ExtractNumberFromFluff(fluff, fluffLen, fdi);
			return tempInt - 1;
			break;
		case 2:
			tempInt = ExtractNumberFromFluff(fluff, fluffLen, fdi);
			return tempInt - 1;
			break;
		case 3:
			/* Allow some slack before the text starts */
			while( ((*fdi) < fluffLen) && (!isalpha(fluff[(*fdi)])) ) (*fdi)++;
			/* get the month name */
			i = 0;
			while( ((*fdi) < fluffLen) && (isalpha(fluff[(*fdi)])) ) {
				buffer[i++] = fluff[(*fdi)++];
			}
			buffer[i] = '\0';
			return GetMonthNumFromAbbrevName(lcid, forceUS, buffer);
			break;
		case 4:
		default: /* 
			/* Allow some slack before the text starts */
			while( ((*fdi) < fluffLen) && (!isalpha(fluff[(*fdi)])) ) (*fdi)++;
			/* get the month name */
			i = 0;
			while( ((*fdi) < fluffLen) && (isalpha(fluff[(*fdi)])) ) {
				buffer[i++] = fluff[(*fdi)++];
			}
			buffer[i] = '\0';
			return GetMonthNumFromFullName(lcid, forceUS, buffer);
			break;
			
	}
	return 0;
}

static int ExtractDayFromFluff(
	char* fmt,
	unsigned int fmtLen,
	unsigned int* fi,
	char *fluff,
	unsigned int fluffLen,
	unsigned int* fdi)
{
	int count;
	int tempInt;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */
	count = 0;
	while (((*fi) < fmtLen) && ((fmt[(*fi)] == 'd') || (fmt[(*fi)] == 'D')) ) {
		count++;
		(*fi)++;
	}
	switch (count) {
		case 1:
		case 2:
			tempInt = ExtractNumberFromFluff(fluff, fluffLen, fdi);
			return tempInt;
			break;
		case 3: /* useless fucking day of the week just skip over it */
		case 4:
		default: /* 
			/* Allow some slack before the text starts */
			while( ((*fdi) < fluffLen) && (!isalpha(fluff[(*fdi)])) ) (*fdi)++;
			/* skip the day of the week */
			while( ((*fdi) < fluffLen) && (isalpha(fluff[(*fdi)])) ) (*fdi)++;
			break;
			
	}
	return 0;
}

static double ExtractFractionFromFluff(
	char *fluff,
	unsigned int fluffLen,
	unsigned int* fdi)
{
	int i;
	double temp = 0.0;
	int len = 0;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */
	
	/* find the next numeric char */
	while( ((*fdi) < fluffLen) && (!isdigit(fluff[*fdi])) )(*fdi)++;
	/* get the number */
	while( ((*fdi) < fluffLen) && (isdigit(fluff[*fdi])) ) {
		temp = (temp * 10) + (int)(fluff[*fdi] - '0');
		(*fdi)++;
		len++;
	}
	for (i = 0; i < len; i++)
		temp = temp / 10.0;
	return temp;
}

static int AppendToFormat(
	char* format,
	unsigned int* fi,
	char* str)
{
	/* assumes caller has hdbc list write CS (callers OK: 1998-04-07) */

	unsigned int i;
	for (i = 0; i < MBstrlen(str); i++)
		format[(*fi)++] = str[i];
	return DB_SUCCESS;
}

#define LOTS_OF_DIGITS 18
static int GenerateGFFormat(
	LCID lcid,
	char *format,
	unsigned int* fi)
{	
	int i;
	char* localeBuf;
	int localeBufLen;
	int fracDigits = 0;
	int sepCount;

	/* assumes caller has hdbc list write CS (callers OK: 1998-04-07) */

	localeBufLen = GetLocaleInfo(lcid, LOCALE_IDIGITS, NULL, 0);
	localeBuf = (char*)malloc(localeBufLen + 1);
	localeBufLen = GetLocaleInfo(lcid, LOCALE_IDIGITS, localeBuf, localeBufLen);
	sscanf(localeBuf, "%d", &fracDigits);
	free(localeBuf);

	localeBufLen = GetLocaleInfo(lcid, LOCALE_SGROUPING, NULL, 0);
	localeBuf = (char*)malloc(localeBufLen + 1);
	localeBufLen = GetLocaleInfo(lcid, LOCALE_SGROUPING, localeBuf, localeBufLen);
	sscanf(localeBuf, "%d", &sepCount);
	free(localeBuf);

	for (i = LOTS_OF_DIGITS; i > 0; i--) {
		if ((i % sepCount == 0) && (i != LOTS_OF_DIGITS))
			format[(*fi)++] = ',';
		if (i == 1)
			format[(*fi)++] = '0';
		else
			format[(*fi)++] = '#';
	}
	if (fracDigits > 0) {
		format[(*fi)++] = '.';
		for (i = 0; i < fracDigits; i++)
			format[(*fi)++]  = '0';
	}
	return DB_SUCCESS;
}

static int GenerateGCFormat(
	LCID lcid,
	char *posFmt,
	char *negFmt,
	unsigned int* fi)
{	
	int i;
	char* localeBuf;
	int localeBufLen;
	int fracDigits = 0;
	int sepCount;
	int currMarkPos;

	/* assumes caller has hdbc list write CS (callers OK: 1998-04-07) */

	localeBufLen = GetLocaleInfo(lcid, LOCALE_ICURRDIGITS, NULL, 0);
	localeBuf = (char*)malloc(localeBufLen + 1);
	localeBufLen = GetLocaleInfo(lcid, LOCALE_ICURRDIGITS, localeBuf, localeBufLen);
	sscanf(localeBuf, "%d", &fracDigits);
	free(localeBuf);

	localeBufLen = GetLocaleInfo(lcid, LOCALE_SMONGROUPING, NULL, 0);
	localeBuf = (char*)malloc(localeBufLen + 1);
	localeBufLen = GetLocaleInfo(lcid, LOCALE_SMONGROUPING, localeBuf, localeBufLen);
	sscanf(localeBuf, "%d", &sepCount);
	free(localeBuf);

	localeBufLen = GetLocaleInfo(lcid, LOCALE_ICURRENCY, NULL, 0);
	localeBuf = (char*)malloc(localeBufLen + 1);
	localeBufLen = GetLocaleInfo(lcid, LOCALE_ICURRENCY, localeBuf, localeBufLen);
	sscanf(localeBuf, "%d", &currMarkPos);
	free(localeBuf);
	
	if (currMarkPos == 0) {
		posFmt[(*fi)++] = '$';
	} else if (currMarkPos == 2) {
		posFmt[(*fi)++] = '$';
		posFmt[(*fi)++] = ' ';
	}

	for (i = LOTS_OF_DIGITS; i > 0; i--) {
		if ( (i % sepCount == 0) && (i != LOTS_OF_DIGITS))
			posFmt[(*fi)++] = ',';
		if (i == 1)
			posFmt[(*fi)++] = '0';
		else
			posFmt[(*fi)++] = '#';
	}
	if (fracDigits > 0) {
		posFmt[(*fi)++] = '.';
		for (i = 0; i < fracDigits; i++)
			posFmt[(*fi)++] = '0';
	}
	if (currMarkPos == 1) {
		posFmt[(*fi)++] = '$';
	} else if (currMarkPos == 3) {
		posFmt[(*fi)++] = ' ';
		posFmt[(*fi)++] = '$';
	}
	localeBufLen = GetLocaleInfo(lcid, LOCALE_ICURRENCY, NULL, 0);
	localeBuf = (char*)malloc(localeBufLen + 1);
	localeBufLen = GetLocaleInfo(lcid, LOCALE_ICURRENCY, localeBuf, localeBufLen);
	sscanf(localeBuf, "%d", &currMarkPos);
	free(localeBuf);
	posFmt[(*fi)] = '\0';
	if (negFmt != NULL) {
		*fi = 0;
		switch (currMarkPos) {
			case 0:
				negFmt[(*fi)++] = '(';
				negFmt[(*fi)++] = '$';
				break;
			case 1:
				negFmt[(*fi)++] = '-';
				negFmt[(*fi)++] = '$';
				break;
			case 2:
				negFmt[(*fi)++] = '$';
				negFmt[(*fi)++] = '-';
				break;
			case 3:
				negFmt[(*fi)++] = '$';
				break;
			case 4:
				negFmt[(*fi)++] = '$';
				negFmt[(*fi)++] = '(';
				break;
			case 5:
			case 8:
				negFmt[(*fi)++] = '-';
				break;
			case 9:
				negFmt[(*fi)++] = ' ';
				negFmt[(*fi)++] = '-';
				negFmt[(*fi)++] = '$';
		}
		for (i = LOTS_OF_DIGITS; i > 0; i--) {
			if ((i % sepCount == 0) && (i != LOTS_OF_DIGITS))
				negFmt[(*fi)++] = ',';
			if (i == 1)
				negFmt[(*fi)++] = '0';
			else
				negFmt[(*fi)++] = '#';
		}
		if (fracDigits > 0) {
			negFmt[(*fi)++] = '.';
			for (i = 0; i < fracDigits; i++)
				negFmt[(*fi)++] = '0';
		}
		if (currMarkPos == 1) {
			negFmt[(*fi)++] = '$';
		} else if (currMarkPos == 3) {
			negFmt[(*fi)++] = ' ';
			negFmt[(*fi)++] = '$';
		}
		switch (currMarkPos) {
			case 0:
				negFmt[(*fi)++] = ')';
				break;
			case 3:
				negFmt[(*fi)++] = '-';
				break;
			case 4:
				negFmt[(*fi)++] = '$';
				negFmt[(*fi)++] = ')';
				break;
			case 5:
				negFmt[(*fi)++] = '$';
				break;
			case 6:
				negFmt[(*fi)++] = '-';
				negFmt[(*fi)++] = '$';
				break;
			case 7:
				negFmt[(*fi)++] = '$';
				negFmt[(*fi)++] = '-';
				break;
			case 8:
				negFmt[(*fi)++] = ' ';
				negFmt[(*fi)++] = '$';
				break;
			case 10:
				negFmt[(*fi)++] = ' ';
				negFmt[(*fi)++] = '$';
				negFmt[(*fi)++] = '-';
				break;
		}
	}
	return DB_SUCCESS;
}

#define GT_US_FORMAT "h:mm:ss AM/PM"
static int GenerateGTFormat(
	LCID lcid,
	int forceUS,
	char *format,
	unsigned int* fi)
{
	unsigned int i;
	char* localeBuf;
	int localeBufLen;
	char* timeSep;
	int use24Hour;
	int meridianIsPrefixed;
	int timeLeadZero;
	
	/* assumes caller has hdbc list write CS (callers OK: 1998-04-07) */

	if (forceUS)
		for (i = 0; i < MBstrlen(GT_US_FORMAT); i++)
			format[(*fi)++] = GT_US_FORMAT[i];
	else {
		localeBufLen = GetLocaleInfo(lcid, LOCALE_STIME, NULL, 0);
		timeSep = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_STIME, timeSep, localeBufLen);

		localeBufLen = GetLocaleInfo(lcid, LOCALE_ITIME, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_ITIME, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &use24Hour);
		free(localeBuf);

		localeBufLen = GetLocaleInfo(lcid, LOCALE_ITIMEMARKPOSN, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_ITIMEMARKPOSN, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &meridianIsPrefixed);
		free(localeBuf);
		
		if (!use24Hour && meridianIsPrefixed) {
			AppendToFormat(format, fi, "l/l ");
		}

		localeBufLen = GetLocaleInfo(lcid, LOCALE_ITLZERO, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_ITLZERO, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &timeLeadZero);
		free(localeBuf);
		
		if (timeLeadZero)
			AppendToFormat(format, fi, "hh");
		else
			AppendToFormat(format, fi, "h");

		AppendToFormat(format, fi, timeSep);
		
		AppendToFormat(format, fi, "ii");

		AppendToFormat(format, fi, timeSep);
		
		AppendToFormat(format, fi, "ss");

		if (!use24Hour && !meridianIsPrefixed) {
			AppendToFormat(format, fi, " l/l ");
		}
		free(timeSep);
	}
	
	return DB_SUCCESS;	
}

#define GL_US_FORMAT "Dddd, Mmmm d, yyyy"
static int GenerateGLFormat(
	LCID lcid,
	int forceUS,
	char *format,
	unsigned int* fi)
{
	unsigned int i;
	char* localeBuf;
	int localeBufLen;
	char* dateSep;
	int dateOrder;
	int monLeadZero;
	int dayLeadZero;
	int fullCentury;
	
	/* assumes caller has hdbc list write CS (callers OK: 1998-04-07) */

	if (forceUS)
		for (i = 0; i < MBstrlen(GL_US_FORMAT); i++)
			format[(*fi)++] = GL_US_FORMAT[i];
	else {
		localeBufLen = GetLocaleInfo(lcid, LOCALE_SDATE, NULL, 0);
		dateSep = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_SDATE, dateSep, localeBufLen);
		
		localeBufLen = GetLocaleInfo(lcid, LOCALE_ILDATE, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_ILDATE, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &dateOrder);
		free(localeBuf);

		localeBufLen = GetLocaleInfo(lcid, LOCALE_ICENTURY, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_ICENTURY, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &fullCentury);
		free(localeBuf);

		localeBufLen = GetLocaleInfo(lcid, LOCALE_IMONLZERO, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_IMONLZERO, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &monLeadZero);
		free(localeBuf);

		localeBufLen = GetLocaleInfo(lcid, LOCALE_IDAYLZERO, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_IDAYLZERO, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &dayLeadZero);
		free(localeBuf);
		
		AppendToFormat(format, fi, "Dddd, ");

		switch (dateOrder) {
			case 0:
				AppendToFormat(format, fi, "Mmmm");
				AppendToFormat(format, fi, " ");
				AppendToFormat(format, fi, "d");
				AppendToFormat(format, fi, ", ");
				AppendToFormat(format, fi, "yyyy");
				break;
			case 1:
				AppendToFormat(format, fi, "d");
				AppendToFormat(format, fi, " ");
				AppendToFormat(format, fi, "Mmmm");
				AppendToFormat(format, fi, " ");
				AppendToFormat(format, fi, "yyyy");
				break;
			case 2:
				AppendToFormat(format, fi, "yyyy");
				AppendToFormat(format, fi, " ");
				AppendToFormat(format, fi, "Mmmm");
				AppendToFormat(format, fi, " ");
				AppendToFormat(format, fi, "d");
				break;
		}
		free(dateSep);
	}
	
	return DB_SUCCESS;	
}

#define GD_US_FORMAT "m/d/yy"
static int GenerateGDFormat(
	LCID lcid,
	int forceUS,
	char *format,
	unsigned int* fi)
{
	unsigned int i;
	char* localeBuf;
	int localeBufLen;
	char* dateSep;
	int dateOrder;
	int monLeadZero;
	int dayLeadZero;
	int fullCentury;
	int useNumericMonth = TRUE;
	
	/* assumes caller has hdbc list write CS (callers OK: 1998-04-07) */

	if (forceUS)
		for (i = 0; i < MBstrlen(GD_US_FORMAT); i++)
			format[(*fi)++] = GD_US_FORMAT[i];
	else {
		localeBufLen = GetLocaleInfo(lcid, LOCALE_SDATE, NULL, 0);
		dateSep = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_SDATE, dateSep, localeBufLen);
		
		localeBufLen = GetLocaleInfo(lcid, LOCALE_IDATE, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_IDATE, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &dateOrder);
		free(localeBuf);

		localeBufLen = GetLocaleInfo(lcid, LOCALE_ICENTURY, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_ICENTURY, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &fullCentury);
		free(localeBuf);

		localeBufLen = GetLocaleInfo(lcid, LOCALE_IMONLZERO, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_IMONLZERO, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &monLeadZero);
		free(localeBuf);
		
		localeBufLen = GetLocaleInfo(lcid, LOCALE_SSHORTDATE, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_SSHORTDATE, localeBuf, localeBufLen);
		if (MBstrstr (localeBuf, "MMM") != NULL)
			useNumericMonth = FALSE;
		free(localeBuf);


		localeBufLen = GetLocaleInfo(lcid, LOCALE_IDAYLZERO, NULL, 0);
		localeBuf = (char*)malloc(localeBufLen + 1);
		localeBufLen = GetLocaleInfo(lcid, LOCALE_IDAYLZERO, localeBuf, localeBufLen);
		sscanf(localeBuf, "%d", &dayLeadZero);
		free(localeBuf);

		switch (dateOrder) {
			case 0:
				if (!useNumericMonth) {
					AppendToFormat(format, fi, "Mmm");
				} else {
					if (monLeadZero) 
						AppendToFormat(format, fi, "mm");
					else
						AppendToFormat(format, fi, "m");
				}
				AppendToFormat(format, fi, dateSep);
				if (dayLeadZero) 
					AppendToFormat(format, fi, "dd");
				else
					AppendToFormat(format, fi, "d");
				AppendToFormat(format, fi, dateSep);
				if (fullCentury)
					AppendToFormat(format, fi, "yyyy");
				else
					AppendToFormat(format, fi, "yy");
				break;
			case 1:
				if (dayLeadZero) 
					AppendToFormat(format, fi, "dd");
				else
					AppendToFormat(format, fi, "d");
				AppendToFormat(format, fi, dateSep);
				if (!useNumericMonth) {
					AppendToFormat(format, fi, "Mmm");
				} else {
					if (monLeadZero) 
						AppendToFormat(format, fi, "mm");
					else
						AppendToFormat(format, fi, "m");
				}
				AppendToFormat(format, fi, dateSep);
				if (fullCentury)
					AppendToFormat(format, fi, "yyyy");
				else
					AppendToFormat(format, fi, "yy");
				break;
			case 2:
				if (fullCentury)
					AppendToFormat(format, fi, "yyyy");
				else
					AppendToFormat(format, fi, "yy");
				AppendToFormat(format, fi, dateSep);
				if (!useNumericMonth) {
					AppendToFormat(format, fi, "Mmm");
				} else {
					if (monLeadZero) 
						AppendToFormat(format, fi, "mm");
					else
						AppendToFormat(format, fi, "m");
				}
				AppendToFormat(format, fi, dateSep);
				if (dayLeadZero) 
					AppendToFormat(format, fi, "dd");
				else
					AppendToFormat(format, fi, "d");
				break;
		}
		free(dateSep);
	}	
	return DB_SUCCESS;	
}

#define ISLEAP(year) \
  (((year) & 0x3) == 0 && ((1900+year) % 400 == 0 || (year) % 100 != 0))

static long LeapDaysFrom1900(long year)
{
	long i, result;
//	if (year > 1801 && year < 2099) {
	if (year > 0 && year < 99) {
		/* This method valid only from 1900 to 2099 but is much simpler */
  		return (year > 0 ? (year - 1) / 4 : year <= -4 ? 1 + (4 - year) / 4 : 0);
  	} else 	{
  		/* This method is valid for all years, but is inefficient for large years 	*/
  		if (year < 0) {
  			/* damn this shit is tricky.  for dates before 1900 we have to include	*/
  			/* the year in question in the leap year check because we are counting	*/
  			/* negative days from 1900.  Later, when we do the month part we'll be	*/
  			/* going back towards 1900 again and may or may not add the leap day  	*/
  			/* back in.  My head hurts!												*/
		  	for (i = 0,result = 0; i >= year; i--) {
				if (ISLEAP(i)) {
				  	result++;
				  	i-=3;
				}
			}
  		} else {
  			/* for dates after 1900, we don't check the year in question.  We'll	*/
  			/* pick up any leap days when we do the month calculation.				*/
		  	for (i = 0,result = 0; i < year; i++) {
				if (ISLEAP(i)) {
				  	result++;
				  	i+=3;
				}
			}
		}
	 	return result;
	}
}

static int TimeStructToDate(
	struct tm* dt,
	DATE *dateTime,
	double fracSec)
{
	double fracDay;
	int numDays;
	int year;
	int i;
  	const short *monLenTab;
  	long leapDays;
	
	fracDay =  (dt->tm_hour + (dt->tm_min + ((dt->tm_sec + fracSec) / 60.0) ) / 60.0) / 24.0;

	if (dt->tm_year == 0 && dt->tm_mon == 0 && dt->tm_mday == 0) {
		*dateTime = fracDay;
	} else {
		if (dt->tm_year > 100) 
			year = dt->tm_year - 1900;
		else {
			year = dt->tm_year;
			if (year >= 0 && year < 70)
				year += 100;
		}
	
		/* number of days in the years since 1900 */		
		leapDays = LeapDaysFrom1900(year);
		if (year >= 0)
			numDays = leapDays + 365L * year;
		else
			numDays = (365L * year) - leapDays;

		/* days in each month since beginning of year */
		monLenTab = ISLEAP(year) ? leapMonthLenTab : regMonthLenTab;
		if (dt->tm_mon >= 12) goto FormatError;
		for (i = 0; i < dt->tm_mon; i++)
			numDays += monLenTab[i];
		/* days into the month */
		if (dt->tm_mday > 31) goto FormatError;
		numDays += dt->tm_mday;

		/* correct for difference in starting date between old routines and MS DATE type */
		numDays += 1;

		if (numDays < 0)
			*dateTime = (double)numDays - fracDay;
		else
			*dateTime = (double)numDays + fracDay;
	}
	return DB_SUCCESS;
FormatError:
	return DB_FORMAT_ERROR;
}

static int PreScanFormat(
	LCID lcid,
	int forceUS,
	char* fmt,
	int* useScaleFactor,
	double* scaleFactor)
{
	unsigned int fi, ti, si;
	int status;
	int isDivisionFactor;
	char scaleStr[30];
	char* tempFmt = NULL;
	unsigned int fmtLen;

	/* assumes caller has hdbc CS (callers OK, 1998-04-07) */
	fmtLen = MBstrlen(fmt);
	tempFmt = (char*)malloc(MBstrlen(fmt) + 1);
	if (tempFmt == NULL) {
		status = DB_OUT_OF_MEMORY;
		goto FormatError;
	}
	MBstrcpy(tempFmt, fmt);
	
	ti = 0;
	fi = 0;
	while (ti < fmtLen) {
		switch (tempFmt[ti]) {
			case '"':
				fmt[fi++] = tempFmt[ti++];
				while (ti < fmtLen && tempFmt[ti] != '"')
					fmt[fi++] = tempFmt[ti++];
				if (ti < fmtLen && tempFmt[ti] == '"')
					fmt[fi++] = tempFmt[ti++];
				break;
			case '\'':
				fmt[fi++] = tempFmt[ti++];
				while (ti < fmtLen && tempFmt[ti] != '\'')
					fmt[fi++] = tempFmt[ti++];
				if (ti < fmtLen && tempFmt[ti] == '\'')
					fmt[fi++] = tempFmt[ti++];
				break;
			case '[':
				if (ti + 1 < fmtLen && tempFmt[ti + 1] == 'S') {
					fmt[fi++] = tempFmt[ti++];
					fmt[fi++] = tempFmt[ti++];
					if (ti >= fmtLen) {status = DB_FORMAT_ERROR; goto FormatError;}
					if (tempFmt[ti] == '/')
						isDivisionFactor = TRUE;
					else
						isDivisionFactor = FALSE;
					fmt[fi++] = tempFmt[ti++];
					si = 0;
					while (ti < fmtLen && tempFmt[ti] != ']') {
						if (!isdigit(tempFmt[ti]) && tempFmt[ti] != '.') 
							{status = DB_FORMAT_ERROR; goto FormatError;}
						scaleStr[si++] = fmt[fi++] = tempFmt[ti++];
					}
					if (ti >= fmtLen) {status = DB_FORMAT_ERROR; goto FormatError;}
					fmt[fi++] = tempFmt[ti++];
					sscanf(scaleStr, "%le", scaleFactor);
					if (isDivisionFactor)
						*scaleFactor = 1.0 / (*scaleFactor);
					*useScaleFactor = TRUE;
				} else {
					while (ti < fmtLen && tempFmt[ti] != ']')
						fmt[fi++] = tempFmt[ti++];
				}
				break;
			case '%':
				if (fi > 0 && (fmt[fi - 1] == '#' || fmt[fi - 1] == '?' || 
				               fmt[fi - 1] == '0' || fmt[fi - 1] == '.')) {
					*useScaleFactor = TRUE;
					*scaleFactor = 1.0 / 100.0;
				}
				fmt[fi++] = tempFmt[ti++];
				break;
			case 'G':
				ti++;
				if (ti < fmtLen) {
					switch (tempFmt[ti]) {
						case 'N':
						case 'F':
							status = GenerateGFFormat(lcid, fmt, &fi);
							if (status != DB_SUCCESS) goto FormatError;
							ti++;
							break;
						case 'C':
							status = GenerateGCFormat(lcid, fmt, NULL, &fi);
							if (status != DB_SUCCESS) goto FormatError;
							ti++;
							break;
						case 'D':
							status = GenerateGDFormat(lcid, forceUS, fmt, &fi);
							if (status != DB_SUCCESS) goto FormatError;
							ti++;
							if (ti < fmtLen && tempFmt[ti] == 'T') {
								fmt[fi++] = ' ';
								status = GenerateGTFormat(lcid, forceUS, fmt, &fi);
								if (status != DB_SUCCESS) goto FormatError;
								ti++;
							}
							break;
						case 'T':
							status = GenerateGTFormat(lcid, forceUS, fmt, &fi);
							if (status != DB_SUCCESS) goto FormatError;
							ti++;
							break;
						case 'L':
							status = GenerateGLFormat(lcid, forceUS, fmt, &fi);
							if (status != DB_SUCCESS) goto FormatError;
							ti++;
							if (ti < fmtLen && tempFmt[ti] == 'T') {
								fmt[fi++] = ' ';
								status = GenerateGTFormat(lcid, forceUS, fmt, &fi);
								if (status != DB_SUCCESS) goto FormatError;
								ti++;
							}
							break;
						default:
							fmt[fi++] = 'G';
							fmt[fi++] = tempFmt[ti++];
							break;
					}
				}
				break;
			default:
				fmt[fi++] = tempFmt[ti++];
				break;
		}
	}
	fmt[fi] = '\0';
	free(tempFmt);
	return DB_SUCCESS;
FormatError:
	free(tempFmt);
	return status;
}

 int FormatToVariant(
	LCID lcid,
	char* value,
	char* inputFmt,
	_variant_t &valueVariant)
{
	unsigned int i;
	char* fluffyData;
	unsigned int fdi;
	unsigned int fdLen;
	char* fmt = NULL;
	unsigned int fi;
	unsigned int fmtLen;
	char* data = NULL;
	unsigned int di;
	char* ts = NULL;
	unsigned int tsLen;
	char* dp = NULL;
	unsigned int dpLen;
	int forceUS;
	int done = FALSE;
	int status = DB_SUCCESS;
	char fmtChar;
	struct tm dt;
	int tempInt;
	int prevFieldWasHourField = FALSE;
	DATE dateTime;
	int isDateTime = FALSE;
	int twelveHour = FALSE;
	double theNum;
	int useScaleFactor = FALSE;
	double scaleFactor;
	double fracSec = 0.0;

	fmt = (char*)malloc(MBstrlen(inputFmt) + 100);
	if (fmt == NULL) {
		status = DB_OUT_OF_MEMORY;
		goto FormatError;
	}
	MBstrcpy(fmt, inputFmt);
	fmtLen = MBstrlen(fmt);
	fluffyData = (char *)value;
	fdLen = MBstrlen(fluffyData);
	data = (char*)malloc(fdLen + 100);

	forceUS = (MBstrstr (fmt, "[US]") != NULL);
	tsLen = GetLocaleInfo(lcid, LOCALE_STHOUSAND, NULL, 0);
	ts = (char*)malloc(tsLen + 1);
	GetLocaleInfo(lcid, LOCALE_STHOUSAND, ts, tsLen);
	dpLen = GetLocaleInfo(lcid, LOCALE_SDECIMAL, NULL, 0);
	dp = (char*)malloc(dpLen + 1);
	GetLocaleInfo(lcid, LOCALE_SDECIMAL, dp, dpLen);
	PreScanFormat(lcid, forceUS, fmt, &useScaleFactor, &scaleFactor);
	fmtLen = MBstrlen(fmt);
	twelveHour = ((MBstrstr (fmt, "AM/PM") != NULL) || (MBstrstr (fmt, "am/pm") != NULL) ||
			      (MBstrstr (fmt, "A/P") != NULL) || (MBstrstr (fmt, "a/p") != NULL));
				
	/* zero used fields in tm struct */
	dt.tm_sec = 0; dt.tm_min = 0; dt.tm_hour = 0; dt.tm_mday = 0; dt.tm_mon = 0; dt.tm_year = 0;

	fdi = 0; fi = 0; di = 0;
	while ((fi < fmtLen) && (fmt[fi] != '\0')) {
		switch (fmt[fi]) {
			case '"':
			case '\'':
				fmtChar = fmt[fi];
				fi++;
				while((fi < fmtLen) && (fmt[fi] != fmtChar)) {
					if (fmt[fi] == fluffyData[fdi]) {
						fi++; fdi++;
					} else {
						status = DB_FORMAT_ERROR;
						goto FormatError;
					}
				}
				fi++;
				break;
			case '\\':
				if ((fi + 1 < fmtLen))
					if (fmt[fi + 1] == fluffyData[fdi]) {
						fi += 2; fdi++;
					} else {
						status = DB_FORMAT_ERROR;
						goto FormatError;
					}
				break;
			case '-':
				if (fi + 1 >= fmtLen || (fmt[fi + 1] != '#' && fmt[fi + 1] != '?' && fmt[fi + 1] != '0')) {
					/* it's not part of a number format, so it must be some fluffy junk */
					fmtChar = fmt[fi];
					if (fluffyData[fdi] == fmtChar)
						fdi++;
					fi++;
					break;
				}
				/* fall thru into the number format crap */
			case '#':
			case '?':
			case '0':
			case '+':
			case '(':
				done = FALSE;
				while ((fdi < fdLen)) {
					if (fluffyData[fdi] == '+') {
						fdi++;
					} else if (fluffyData[fdi] == '-' || fluffyData[fdi] == '(' ) {
						data[di++] = '-';
						fdi++;
					} else if (fluffyData[fdi] == ')') { 
						fdi++;
					} else if (isdigit(fluffyData[fdi])) {
						data[di++] = fluffyData[fdi++];
					} else if (forceUS && fluffyData[fdi] == ',') {
						fdi++;
					} else if (forceUS && fluffyData[fdi] == '.') {
						for (i = 0; i < MBstrlen(dp); i++) {
							data[di++] = dp[i];
						}
						fdi++;
					} else if (!forceUS && fluffyData[fdi] == ts[0]) {
						for (i = 0; i < MBstrlen(ts); i++) {
							fdi++;
						}
					} else if (!forceUS && fluffyData[fdi] == dp[0]) {
						for (i = 0; i < MBstrlen(dp); i++) {
							data[di++] = dp[i];
							fdi++;
						}
					} else if (fluffyData[fdi] == 'e' || fluffyData[fdi] == 'E') {
						data[di] = 'e';
						fdi++;
					} else {
						goto FormatError;
					}
				}
				data[di] = '\0';
				sscanf(data, "%le", &theNum);
				if (useScaleFactor)
					theNum = theNum * scaleFactor;
				valueVariant = theNum;
				/* We've got a number.  There can't be anything useful left */
				/* in the format or the input data, so screw it and boogie	*/
				goto GitWhileGittinIsGood;
				break;
			case 'y':
				isDateTime = TRUE;
				dt.tm_year = ExtractNumberFromFluff(fluffyData, fdLen, &fdi);
				/* eat the rest of the year format */
				while (fmt[fi] == 'y') fi++;
				prevFieldWasHourField = FALSE;
				break;
			case 'm':
			case 'M':
				isDateTime = TRUE;
				if (prevFieldWasHourField) {
					dt.tm_min = ExtractNumberFromFluff(fluffyData, fdLen, &fdi);
					while (fi < fmtLen && fmt[fi] == 'm') fi++;
				} else {
					dt.tm_mon = ExtractMonthFromFluff(lcid, forceUS, fmt, fmtLen, &fi, fluffyData, 
													  fdLen, &fdi);
				}
				prevFieldWasHourField = FALSE;
				break;
			case 'd':
			case 'D':
				isDateTime = TRUE;
				tempInt = ExtractDayFromFluff(fmt, fmtLen, &fi, fluffyData, fdLen, &fdi);
				if (tempInt >= 0) 
					dt.tm_mday = tempInt;
				prevFieldWasHourField = FALSE;
				break;
			case 'h':
				isDateTime = TRUE;
				prevFieldWasHourField = TRUE;
				dt.tm_hour += ExtractNumberFromFluff(fluffyData, fdLen, &fdi);
				while (fi < fmtLen && fmt[fi] == 'h') fi++;
				break;
			case 'i':
				isDateTime = TRUE;
				dt.tm_min = ExtractNumberFromFluff(fluffyData, fdLen, &fdi);
				while (fi < fmtLen && fmt[fi] == 'i') fi++;
				prevFieldWasHourField = FALSE;
				break;
			case 's':
				isDateTime = TRUE;
				dt.tm_sec = ExtractNumberFromFluff(fluffyData, fdLen, &fdi);
				while (fi < fmtLen && fmt[fi] == 's') fi++;

				if ( (fdi < fdLen) && (fluffyData[fdi] == '.') ) {
					fracSec = ExtractFractionFromFluff(fluffyData, fdLen, &fdi);
//					fdi++;
//					while( (fdi < fdLen) && (isdigit(fluffyData[fdi])) )
//						fdi++;
				}
				if ( (fi + 1 < fmtLen) && (fmt[fi] == '.') && (fmt[fi + 1] == 's') ) {
					fi++;
					while (fi < fmtLen && fmt[fi] == 's') fi++;
				}
				prevFieldWasHourField = FALSE;
				break;
			case 'a':
			case 'A':
			case 'p':
			case 'P':
				if (twelveHour) {
					if (ExtractAMPMFromFluff(lcid, forceUS, fmt, fmtLen, &fi,
											 fluffyData, fdLen, &fdi))
						if (dt.tm_hour < 12)
							dt.tm_hour += 12;
				}
				break;
			case '%':
			case ' ':
			case '$':
			case '/':
			case '.':
			case ',':
				/* more fluff characters. */
				fmtChar = fmt[fi];
				if (fluffyData[fdi] == fmtChar)
					fdi++;
				fi++;
				break;
			case '[':
				fi++;
				while (fi < fmtLen && fmt[fi] != ']')
					fi++;
				fi++;
				break;
			default:
				fi++;
				break;
		}
	}
	if (isDateTime) {
		if ((status = TimeStructToDate(&dt, &dateTime, fracSec)) != DB_SUCCESS) goto FormatError;
		_variant_t date(dateTime, VT_DATE);
		valueVariant = date;
	}

GitWhileGittinIsGood:	
FormatError:

#if 0
  BSTR bstr;
  VarBstrFromDate(dateTime, lcid, 0, &bstr); 
  _bstr_t bstrt = bstr;
  DebugPrintf("%s\n", (char*)bstrt);
#endif
	free(fmt);
	free(ts);
	free(dp);
	free(data);
	return status;
}

