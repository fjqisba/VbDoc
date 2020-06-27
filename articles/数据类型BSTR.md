# 数据类型BSTR

在VB当中，字符串的类型通常是使用BSTR类型来存储和访问的。



例如在VB中进行如下定义

```vb
Dim TestStr As String
TestStr = "ABCDE"
```
程序得到的TestStr变量属于BSTR类型，内存的完整布局如下(总共占用16个字节):

004015B8  0A 00 00 00 41 00 42 00 43 00 44 00 45 00 00 00

前面四个字节代表的是字符串的有效字节数，在这个例子中，这个值是10。接下来的10个字节的内存空间为字符串数据缓冲区，最后再以2个零字节收尾。



实际上这和C语言当中SysAllocString(L"ABCDE")这句代码的意思是一样的。

BSTR和其它的字符串指针稍有不同，比如:

1. BSTR的长度是固定的，如果长度为0，那么代表""。
2. BSTR总是指向缓冲区中的第一个有效字符，例如上述例子中BSTR指针指向的地址为004015BC。



在C语言当中，操作BSTR几乎都是使用SysAllocString()，SysStringLen()，SysFreeString()这些API函数来进行的。

在VB当中同样，也是使用vbaStrCopy()，vbaStrCat()这些VB系统库函数来操作BSTR变量的。