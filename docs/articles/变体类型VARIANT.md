# 变体类型VARIANT

VB中的绝大部分底层函数都会涉及到VARIANT，可以说如果不了解这个变量类型，是无法分析VB程序的。

这里从网上摘一段对VARIANT的介绍:

在VB的基本数据类型中，变体类型(VARIENT)可以表示任何类型的变量 ，而且VARIANT 数据类型具有跨语言的特性。



VARIANT的结构可以参考头文件VC98\Include\OAIDL.H中关于结构体tagVARIANT的定义:

```c
struct tagVARIANT
    {
    union 
        {
        struct __tagVARIANT
            {
            VARTYPE vt;
            WORD wReserved1;
            WORD wReserved2;
            WORD wReserved3;
            union 
                {
                LONGLONG llVal;
                LONG lVal;
                BYTE bVal;
                SHORT iVal;
                FLOAT fltVal;
                DOUBLE dblVal;
                VARIANT_BOOL boolVal;
                _VARIANT_BOOL bool;
                SCODE scode;
                CY cyVal;
                DATE date;
                BSTR bstrVal;
                IUnknown *punkVal;
                IDispatch *pdispVal;
                SAFEARRAY *parray;
                BYTE *pbVal;
                SHORT *piVal;
                LONG *plVal;
                LONGLONG *pllVal;
                FLOAT *pfltVal;
                DOUBLE *pdblVal;
                VARIANT_BOOL *pboolVal;
                _VARIANT_BOOL *pbool;
                SCODE *pscode;
                CY *pcyVal;
                DATE *pdate;
                BSTR *pbstrVal;
                IUnknown **ppunkVal;
                IDispatch **ppdispVal;
                SAFEARRAY **pparray;
                VARIANT *pvarVal;
                PVOID byref;
                CHAR cVal;
                USHORT uiVal;
                ULONG ulVal;
                ULONGLONG ullVal;
                INT intVal;
                UINT uintVal;
                DECIMAL *pdecVal;
                CHAR *pcVal;
                USHORT *puiVal;
                ULONG *pulVal;
                ULONGLONG *pullVal;
                INT *pintVal;
                UINT *puintVal;
                struct __tagBRECORD
                    {
                    PVOID pvRecord;
                    IRecordInfo *pRecInfo;
                    } 	__VARIANT_NAME_4;
                } 	__VARIANT_NAME_3;
            } 	__VARIANT_NAME_2;
        DECIMAL decVal;
        } 	__VARIANT_NAME_1;
    } ;
typedef VARIANT *LPVARIANT;
```

该结构体看起来非常复杂，但关键的值无非是两个，变量类型vt，还有联合结构体中与vt对应的数据类型值。



在头文件wtypes.h中能找到如下与VARTYPE有关的部分信息

```c
enum VARENUM
    {
        VT_EMPTY	= 0,
        VT_NULL	= 1,
        VT_I2	= 2,
        VT_I4	= 3,
        VT_R4	= 4,
        VT_R8	= 5,
        VT_CY	= 6,
        VT_DATE	= 7,
        VT_BSTR	= 8,
        VT_DISPATCH	= 9,
        VT_ERROR	= 10,
        VT_BOOL	= 11,
        VT_VARIANT	= 12,
        VT_UNKNOWN	= 13,
        VT_DECIMAL	= 14,
        VT_I1	= 16,
        VT_UI1	= 17,
        VT_UI2	= 18,
        VT_UI4	= 19,
        VT_I8	= 20,
        VT_UI8	= 21,
        VT_INT	= 22,
        VT_UINT	= 23,
        VT_VOID	= 24,
        VT_HRESULT	= 25,
        VT_PTR	= 26,
        VT_SAFEARRAY	= 27,
        VT_CARRAY	= 28,
        VT_USERDEFINED	= 29,
        VT_LPSTR	= 30,
        VT_LPWSTR	= 31,
        VT_RECORD	= 36,
        VT_INT_PTR	= 37,
        VT_UINT_PTR	= 38,
        VT_FILETIME	= 64,
        VT_BLOB	= 65,
        VT_STREAM	= 66,
        VT_STORAGE	= 67,
        VT_STREAMED_OBJECT	= 68,
        VT_STORED_OBJECT	= 69,
        VT_BLOB_OBJECT	= 70,
        VT_CF	= 71,
        VT_CLSID	= 72,
        VT_VERSIONED_STREAM	= 73,
        VT_BSTR_BLOB	= 0xfff,
        VT_VECTOR	= 0x1000,
        VT_ARRAY	= 0x2000,
        VT_BYREF	= 0x4000,
        VT_RESERVED	= 0x8000,
        VT_ILLEGAL	= 0xffff,
        VT_ILLEGALMASKED	= 0xfff,
        VT_TYPEMASK	= 0xfff
    } ;
typedef ULONG PROPID;
```

枚举类型VARENUM包含非常多的变量类型，其中VARTYPE属于被VARENUM包括在内的。



了解了这些信息，我们就知道如何去确定变量的值了。

//。。。。。。