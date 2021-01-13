# 安全数组SAFEARRAY

SAFEARRAY也就是安全数组，是COM组件通用数据交互中十分重要的数据结构。

而VB正是在COM组件基础上设计的语言，因此SAFEARRAY在VB逆向分析中是必须要了解的。

SAFEARRAY在头文件中有关的结构定义如下:

```c
typedef struct tagSAFEARRAY
{
    USHORT cDims;
    USHORT fFeatures;
    ULONG cbElements;
    ULONG cLocks;
    PVOID pvData;
    SAFEARRAYBOUND rgsabound[1];
} SAFEARRAY;
```

其中各个成员的描述如下:

cDims值等于数组的维数。

fFeatures是用来描述数组如何分配和如何被释放的标志，一般由如下值组合而成:

| fFeatures安全数组的标记 | 说明                |
| ----------------------- | ------------------- |
| FADF_AUTO 0x0001        | 在栈上创建数组      |
| FADF_STATIC 0x0002      | 在堆上创建数组      |
| FADF_EMBEDDED 0x0004    | 在结构中创建        |
| FADF_FIXEDSIZE 0x0010   | 不能改变数组大小    |
| FADF_RECORD 0x0020      | 记录容器            |
| FADF_HAVEIID 0x0040     | 有IID 身份标记 数组 |
| FADF_HAVEVARTYPE 0x0080 | VT 类型数组         |
| FADF_BSTR 0x0100        | BSTR数组            |
| FADF_UNKNOWN 0x0200     | IUnknown指针 数组   |
| FADF_DISPATCH 0x0400    | IDispatch指针 数组  |
| FADF_VARIANT 0x0800     | VARIANT 数组        |
| FADF_RESERVED 0xF0E8    | 保留值，将来使用    |

cbElements值等于数组中元素占用字节大小。

cLocks为一个计数器，用来跟踪该数组被锁定的次数。

pvData为指向数据缓冲区的指针。

rgsabound用来描述数组每维的数组结构信息，结构信息在头文件结构定义如下:

```c
typedef struct tagSAFEARRAYBOUND
{
    ULONG cElements;		//第N维数组的元素个数
    LONG lLbound;		//第N维数组的下标(下标从0开始)
}SAFEARRAYBOUND;
```

虽然rgsabound在SAFEARRAY中被定义为固定大小数组，但实际上该数组的大小是可变的。

------

例如在VB中定义如下代码，然后生成一个EXE

```vb
Dim test(1) As Integer
test(0) = 1
test(1) = 2
```

数组test将在内存中有如下表现:

0018FAB0  01 00 92 00 02 00 00 00 00 00 00 00 78 80 52 00

0018FAC0  02 00 00 00 00 00 00 00

数组test成员pvData内存表现如下:

00528078  01 00 02 00



通过分析内存可以得到如下对应结果:

该数组为一维数组

该数组是在堆上创建的、不能改变大小的、存储VT类型元素的数组

该数组每个成员的大小为两个字节

该数组未被锁定

该数组包含两个成员，分别为0x1和0x2。



再使用IDA分析生成的VB程序，会得到如下的程序代码

```c
 SAFEARRAY test;
 _vbaAryConstruct2(&test, ArrayStructdes, 2);
 test.pvData[0] = 1;
 test.pvData[1] = 2;
```

