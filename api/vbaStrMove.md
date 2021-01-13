# vbaStrMove

字符串指针的赋值。



## 函数原型:

```c
BSTR __fastcall _vbaStrMove(BSTR *dest, BSTR src)
{
    if ( *dest )
    {
        SysFreeString(*dest);
    }
    *dest = src;
    return src;
}
```

将src字符串指针赋值给dest，如果dest中已有字符串指针，则将其释放后再进行赋值。



## 返回值:

src字符串指针

