#include "excel.h"
#include <QFileDialog>
#include<ActiveQt/QAxObject>

excel::excel()
{

}
bool excel::create_excel()
{


     return true;
}
QString excel::saveas()
{
    QString file;
    QString filter;
    //如果版本低于QT5,则需要将:
    //  QStandardPaths::writableLocation(QStandardPaths::DesktopLocation),
    //改为:QDesktopServices::storageLocation(QDesktopServices::DesktopLocation),
    file = QFileDialog::getSaveFileName (
     NULL,                               //父组件
    "选择生成excel文件存放目录",                              //标题
     QStandardPaths::writableLocation(QStandardPaths::DesktopLocation),                 //设置路径, .表示当前路径,./表示更目录
     "Excel(*.xlsx)",     //过滤器
     &filter  );

    return file;
}
void  excel::Excel_SetCell(QAxObject *worksheet,EXcel_ColumnType column,int row,QColor color,QString text)
{

  QAxObject *cell = worksheet->querySubObject("Cells(int,int)", row, column);
  cell->setProperty("Value", text);
  QAxObject *font = cell->querySubObject("Font");
  font->setProperty("Color", color);
}

//把QVariant转为QList<QList<QVariant> >,用于快速读出的
void excel::castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res)
{
    QVariantList varRows = var.toList();
    if(varRows.isEmpty())
    {
        return;
    }

    const int rowCount = varRows.size();
    QVariantList rowData;

    for(int i=0;i<rowCount;++i)
    {
        rowData = varRows[i].toList();
        res.push_back(rowData);
    }
}

//把QList<QList<QVariant> > 转为QVariant,用于快速写入的
void excel::castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res)
{
    QVariantList vars;
    const int rows = cells.size();
    for(int i=0;i<rows;++i)
    {
        vars.append(QVariant(cells[i]));
    }
    res = QVariant(vars);
}
