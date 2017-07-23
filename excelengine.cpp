/**************************************************************
 * File Name   : excelengine.cpp
 * Author      : ThreeDog
 * Date        : Tue Jul 18 13:36:13 2017
 * Description : 操作Excel的类的封装。
 *
 **************************************************************/

#include "excelengine.h"
#include <QDebug>

ExcelEngine* ExcelEngine::s_pExcelEngine = NULL;

ExcelEngine::ExcelEngine(QWidget *parent)
    :QWidget(parent)
{
    m_pExcel = new QAxObject("Excel.Application");
    m_bIsVisible = false;
    m_pExcel->setProperty("Visible",m_bIsVisible);
    m_pWorkBooks = m_pExcel->querySubObject("WorkBooks");
    m_pExcel->setProperty("Caption", "Qt Excel");
    m_iRowHeight = 30;
    m_iColumnWidth = 10;
    m_bWarpText = true;
    m_cBorderColor = QColor(0,0,0);
    m_sFontFamily = "微软雅黑";
    m_iFontSize = 10;
    m_bFontItalic = false;
    m_bFontBold = false;
    m_cBackgroundColor = QColor(255,255,255);
}

ExcelEngine *ExcelEngine::getInstance(QWidget *parent)
{
    if(s_pExcelEngine == NULL){
        s_pExcelEngine = new ExcelEngine(parent);
        return s_pExcelEngine;
    }
    else
        return s_pExcelEngine;
}

bool ExcelEngine::createExcelFile(const QString &fileName)
{
    //qDebug()<<"in createExcelFile filename:"<<fileName;
    if(fileName.isNull())
        return false;
    if(fileName == "")
        return false;
    QAxObject *p = new QAxObject();
    p->setControl("Excel.Application");
    p->dynamicCall("SetVisible(bool)",false);
    p->setProperty("DisplayAlerts",false);
    QAxObject *pp = p->querySubObject("Workbooks");
    //qDebug()<<"Create Excel File Success";
    QFile file(fileName);
    QAxObject * ppp;
    if(!file.exists()){
        pp->dynamicCall("Add");
        ppp = p->querySubObject("ActiveWorkBook");
        ppp->dynamicCall("SaveAs(const QString &)",fileName);

        ppp->dynamicCall("Close(Boolean)",false);
        p->dynamicCall("Quit(void)");
    }
    //qDebug()<<"Create Excel File Success";
    return true;
}

void ExcelEngine::setExcelVisible(const bool visible)
{
    m_bIsVisible = visible;
    //m_pExcel->setProperty("Visible",visible);
}

void ExcelEngine::setSheetName(const QString &name)
{
    m_pWorkSheet->setProperty("Name",name);
}

//第x行，第y列
bool ExcelEngine::writeText(int x, int y, const QString &text)
{
    m_pCell = m_pWorkSheet->querySubObject("Cells(int,int)", x, y);
    this->applyCellSettings();
    m_pCell->setProperty("Value", text);  //设置单元格文字
    return true;
}

bool ExcelEngine::readText()
{
    int sheet_count = m_pWorkSheets->property("Count").toInt();
    if(sheet_count > 0)
    {
        QAxObject *work_sheet = m_pWorkBook->querySubObject("Sheets(int)", 1);
        QAxObject *used_range = work_sheet->querySubObject("UsedRange");
        QAxObject *rows = used_range->querySubObject("Rows");
        QAxObject *columns = used_range->querySubObject("Columns");
        row_start = used_range->property("Row").toInt();  //获取起始行
        column_start = used_range->property("Column").toInt();  //获取起始列
        row_count = rows->property("Count").toInt();  //获取行数
        column_count = columns->property("Count").toInt();  //获取列数
        for(int i=row_start; i <= row_count; i++)
        {
            for(int j=column_start; j <= column_count;j++)
            {
                QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", i, j);
                QString value = cell->property("Value").toString();
                ExcelNode * new_node = new ExcelNode(i,j,value);
                m_vExcleNodeList.append(new_node);

            }
        }
        return true;
    }else
        return false;
}

void ExcelEngine::setRowHeight(const int i)
{
    m_pCell->setProperty("RowHeight", i);  //设置单元格行高
    this->m_iRowHeight = i;
}

void ExcelEngine::setColumnWidth(const int i)
{
    m_pCell->setProperty("ColumnWidth", i);  //设置单元格列宽
    this->m_iColumnWidth = i;
}

void ExcelEngine::setWarpText(const bool b)
{
    m_pCell->setProperty("WrapText", b);
    this->m_bWarpText = b;
}

void ExcelEngine::setBorderColor(const QColor &color)
{
    this->m_cBorderColor = color;
}

void ExcelEngine::setFontFamily(const QString &family)
{
    this->m_sFontFamily = family;
}

void ExcelEngine::setFontSize(const int s)
{
    this->m_iFontSize = s;
}

void ExcelEngine::setFontItalic(const bool b)
{

    this->m_bFontItalic = b;
}

void ExcelEngine::setFontBold(const bool b)
{
    this->m_bFontBold = b;
}

void ExcelEngine::setBackgroundColor(const QColor &color)
{
    this->m_cBackgroundColor = color;
}

void ExcelEngine::clearCell()
{
    m_pCell->dynamicCall("ClearContents()");
}

void ExcelEngine::setActiveExcel(const QString &fileName)
{
    m_pWorkBooks->dynamicCall("Open(const QString&)", fileName);
    m_pWorkBook = m_pExcel->querySubObject("ActiveWorkBook");

    m_pWorkSheets = m_pWorkBook->querySubObject("Sheets");
    m_pLastSheet = m_pWorkSheets->querySubObject("Item(int)", 1);
    m_pWorkSheet = m_pWorkBook->querySubObject("Sheets(int)", 1);
    m_pLastSheet->dynamicCall("Move(QVariant)", m_pWorkSheet->asVariant());
    m_pWorkSheet->setProperty("Name","sheet");

}

void ExcelEngine::closeActiveExcel()
{
    m_pWorkBook->dynamicCall("Close(Boolean)", false);  //关闭文件
    m_pExcel->dynamicCall("Quit(void)");  //退出
}

void ExcelEngine::saveExcel()
{
    m_pWorkBook->dynamicCall("Save()");
}

void ExcelEngine::saveAsExcel(const QString &fileName)
{
    m_pWorkBook->dynamicCall("SaveAs(const QString&)", fileName);  //另存为另一个文件
}

QString ExcelEngine::getExcelTitle()
{
    QVariant title = m_pExcel->property("Caption");
    return title.toString();
}

QString ExcelEngine::getExcelValue(int x, int y)
{
    if(m_vExcleNodeList.count()!=0){
        for(int i = 0; i < m_vExcleNodeList.count();i++){
            ExcelNode *en = m_vExcleNodeList.at(i);
            if(en->getX() == x && en->getY() == y)
                return en->getValue();
        }
    }
    return NULL;
}

void ExcelEngine::applyCellSettings()
{
    //设置单元格一些属性
    this->setRowHeight(m_iRowHeight);
    //m_pCell->setProperty("")
    this->setColumnWidth(m_iColumnWidth);
    this->setWarpText(m_bWarpText);
    m_pCell->setProperty("HorizontalAlignment", -4108);
    //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152
    m_pCell->setProperty("VerticalAlignment", -4108);
    //上对齐（xlTop）-4160 居中（xlCenter）：-4108  下对齐（xlBottom）：-4107
    m_pInterior = m_pCell->querySubObject("Interior");
    m_pInterior->setProperty("Color", m_cBackgroundColor);
    m_pBorder = m_pCell->querySubObject("Borders");
    m_pBorder->setProperty("Color", m_cBorderColor);
    m_pFont   = m_pCell->querySubObject("Font");
    m_pFont->setProperty("Name",m_sFontFamily);
    m_pFont->setProperty("Italic", m_bFontItalic);
    m_pFont->setProperty("Size", m_iFontSize);
    m_pFont->setProperty("Bold", m_bFontBold);

}

int ExcelEngine::getColumn_count() const
{
    return column_count;
}

void ExcelEngine::setColumn_count(int value)
{
    column_count = value;
}

int ExcelEngine::getRow_count() const
{
    return row_count;
}

void ExcelEngine::setRow_count(int value)
{
    row_count = value;
}

int ExcelEngine::getColumn_start() const
{
    return column_start;
}

void ExcelEngine::setColumn_start(int value)
{
    column_start = value;
}

int ExcelEngine::getRow_start() const
{
    return row_start;
}

void ExcelEngine::setRow_start(int value)
{
    row_start = value;
}

ExcelEngine::~ExcelEngine()
{

}
