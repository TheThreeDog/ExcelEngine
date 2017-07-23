/**************************************************************
 * File Name   : excelengine.h
 * Author      : ThreeDog
 * Date        : Tue Jul 18 13:36:13 2017
 * Description : 操作Excel的类的封装。
 *
 **************************************************************/
#ifndef _EXCELENGINE_H_ 
#define _EXCELENGINE_H_ 
#include <QWidget>
#include <QAxObject>
#include <QFile>
#include <QDebug>
#include <QMessageBox>
#include "excelnode.h"
class ExcelEngine:public QWidget
{
    Q_OBJECT
public:
    static ExcelEngine * getInstance(QWidget* parent = 0);
    static ExcelEngine * s_pExcelEngine;
    ~ExcelEngine();
    int getRow_start() const;
    void setRow_start(int value);

    int getColumn_start() const;
    void setColumn_start(int value);

    int getRow_count() const;
    void setRow_count(int value);

    int getColumn_count() const;
    void setColumn_count(int value);

public slots:
    bool createExcelFile(const QString &fileName);      //创建一个Excel文件
    void setExcelVisible(const bool visible);           //设置Excel是否可见
    void setSheetName(const QString &name);             //设置工作表名称
    bool writeText(int x, int y ,const QString &text);  //向单元格写入文字
    bool readText();                                    //读取单元格中的数据
    void setRowHeight(const int i);                     //设置行高
    void setColumnWidth(const int i);                   //设置列宽
    void setWarpText(const bool b);                     //设置自动换行
    void setBorderColor(const QColor &color);           //设置边框颜色
    void setFontFamily(const QString &family);          //设置字体
    void setFontSize(const int s);                      //设置字体大小
    void setFontItalic(const bool b);                   //设置是否倾斜
    void setFontBold(const bool b);                     //设置字体加粗
    void setBackgroundColor(const QColor & color);      //设置单元格背景颜色
    void clearCell();                                   //清空单元格内容
    void setActiveExcel(const QString &fileName);       //设置当前操作的Excel文件
    void closeActiveExcel();                            //关闭当前正在操作的Excel文件
    void saveExcel();                                   //保存
    void saveAsExcel(const QString &fileName);          //另存为
    QString getExcelTitle();                            //获取文件的标题
    QString getExcelValue(int x,int y);                 //获取单元格中的数据
private:
    QAxObject *m_pExcel;
    QAxObject *m_pWorkBooks;
    QAxObject *m_pWorkBook;
    QAxObject *m_pWorkSheets;
    QAxObject *m_pWorkSheet;
    QAxObject *m_pLastSheet;
    QAxObject *m_pCell;         //单元格
    QAxObject *m_pBorder;       //单元格边框
    QAxObject *m_pFont;         //单元格文字
    QAxObject *m_pInterior;     //单元格背景色
    bool m_bIsVisible;

    int m_iRowHeight;           //行高
    int m_iColumnWidth;         //列宽
    bool m_bWarpText;           //自动换行
    QColor m_cBorderColor;      //边框颜色
    QString m_sFontFamily;      //字体
    int m_iFontSize;            //字体大小
    bool m_bFontItalic;         //字体倾斜
    bool m_bFontBold;           //字体加粗
    QColor m_cBackgroundColor;  //背景颜色
    QVector<ExcelNode *> m_vExcleNodeList;              //存放所有的Excel节点

    explicit ExcelEngine(QWidget *parent = 0);
    void applyCellSettings();                           //应用单元格设置

    int row_start ;             //起始行
    int column_start;           //起始列
    int row_count ;             //行总数
    int column_count;           //列总数
};

#endif  //EXCELENGINE
