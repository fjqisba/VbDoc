# vbaStrCopy

字符串拷贝函数。



## 函数原型:

```c
BSTR __fastcall _vbaStrCopy(BSTR *dest, BSTR src);
```



函数主要功能为将src中的字符串进行拷贝，生成得到一份新的字符串dest，如果原先的dest已有内容，则将其内容释放再进行拷贝。



## 返回值:

dest的值