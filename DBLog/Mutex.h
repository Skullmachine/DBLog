// Mutex.h : Declaration of the CMutex

#ifndef __MUTEX_H_
#define __MUTEX_H_

/////////////////////////////////////////////////////////////////////////////
// CMutex
class CMutex
{
public:
	CMutex();
	~CMutex();

// IMutex
public:
	void Initialize (_bstr_t Name, bool MakeProcessUnique, IDispatch &sequenceContextDisp, _bstr_t &errMsg, long *errorCode = NULL);

private:
	HANDLE mMutexHandle;
};

#endif //__MUTEX_H_
