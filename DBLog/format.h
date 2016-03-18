#pragma once
#ifndef __FORMAT_H_
#define __FORMAT_H_

#define DB_SUCCESS               0       /* Success. */
#define DB_OUT_OF_MEMORY         -3      /* Insufficient memory for operation. */
#define DB_FORMAT_ERROR -71

int FormatToVariant(LCID lcid, char* value, char* inputFmt, _variant_t &valueVariant);

#endif