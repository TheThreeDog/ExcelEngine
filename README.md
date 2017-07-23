
## 名称:
- ExcelEngine

## 软件简介
-    一个基于Qt的Excel操作引擎，封装了操作Excel文件的部分接口，相比直接使用QAxObject更加方便，代码可读性也更好。

## 说明:
- 开发平台: Windows 7 X64.
- 开发环境: Qt 5.7
- 第一次在项目中需要时需要读写Excel文件，第一版直接使用QAxObject，类的封装结构不够好，代码也很乱，所以封装了一个单例的ExcelEngine类，可以通过ExcelEngine的外部接口读写操作Excel的内容。QAxObject中还提供了更多的更丰富的功能，我这里并没能全部封装，只有最基本的单元格的读写操作接口。

## 用法:
- static ExcelEngine * getInstance(QWidget* parent = 0);    //获取单例模式下的对象
- int getRow_start() const;                                                     //获取起始行
- void setRow_start(int value);                                               //设置起始行                       

- int getColumn_start() const;                                               //获取起始列
- void setColumn_start(int value);                                         //设置起始列

- int getRow_count() const;                                                   //获取行数
- void setRow_count(int value);                                             //设置行数

- int getColumn_count() const;                                             //获取列数
- void setColumn_count(int value);                                       //设置列数

- bool createExcelFile(const QString &fileName);      //创建一个Excel文件
- void setExcelVisible(const bool visible);           //设置Excel是否可见
- void setSheetName(const QString &name);             //设置工作表名称
- bool writeText(int x, int y ,const QString &text);  //向单元格写入文字
- bool readText();                                    //读取单元格中的数据
- void setRowHeight(const int i);                     //设置行高
- void setColumnWidth(const int i);                   //设置列宽
- void setWarpText(const bool b);                     //设置自动换行
- void setBorderColor(const QColor &color);           //设置边框颜色
- void setFontFamily(const QString &family);          //设置字体
- void setFontSize(const int s);                      //设置字体大小
- void setFontItalic(const bool b);                   //设置是否倾斜
- void setFontBold(const bool b);                     //设置字体加粗
- void setBackgroundColor(const QColor & color);      //设置单元格背景颜色
- void clearCell();                                   //清空单元格内容
- void setActiveExcel(const QString &fileName);       //设置当前操作的Excel文件
- void closeActiveExcel();                            //关闭当前正在操作的Excel文件
- void saveExcel();                                   //保存
- void saveAsExcel(const QString &fileName);          //另存为
- QString getExcelTitle();                            //获取文件的标题
- QString getExcelValue(int x,int y);                 //获取单元格中的数据


## 开源协议:
- 雪碧软件协议
