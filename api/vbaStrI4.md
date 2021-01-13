# vbaStrI4

将int类型的变量转换成BSTR。



## 函数原型:

```c
BSTR __stdcall _vbaStrI4(int value);
```

如果value的值为负数，那么得到的字符串也包含负号。



## 返回值:

value转换得到的字符串。