/**************************************************************
 * File Name   : excelnode.h
 * Author      : ThreeDog
 * Date        : Fri Jul 21 15:48:56 2017
 * Description : Excel的节点类，每一个对象存放Excel表格的坐标和表格中的数据
 *
 **************************************************************/
#ifndef _EXCELNODE_H_ 
#define _EXCELNODE_H_ 
#include <QString>

class ExcelNode
{
public:
    ExcelNode();
    ExcelNode(int _x, int _y, QString _value);
    void setX(const int _x);
    void setY(const int _y);
    void setValue(const QString &_value);
    int getX();
    int getY();
    QString getValue();
    ~ExcelNode();
private:
    int x;
    int y;
    QString value;

};

#endif  //EXCELNODE
