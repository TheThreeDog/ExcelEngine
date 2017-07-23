/**************************************************************
 * File Name   : excelnode.cpp
 * Author      : ThreeDog 
 * Date        : Fri Jul 21 15:48:56 2017
 * Description : Excel的节点类，每一个对象存放Excel表格的坐标和表格中的数据
 *
 **************************************************************/

#include "excelnode.h"

ExcelNode::ExcelNode()
{

}

ExcelNode::ExcelNode(int _x, int _y, QString _value)
{
    this->x = _x;
    this->y = _y;
    this->value = _value;
}

void ExcelNode::setX(const int _x)
{
    this->x = _x;
}

void ExcelNode::setY(const int _y)
{
    this->y = _y;
}

void ExcelNode::setValue(const QString &_value)
{
    this->value = _value;
}

int ExcelNode::getX()
{
    return this->x;
}

int ExcelNode::getY()
{
    return this->y;
}

QString ExcelNode::getValue()
{
    return this->value;
}

ExcelNode::~ExcelNode()
{

}
