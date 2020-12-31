#ifndef EXCEL_H
#define EXCEL_H
#include <qdebug.h>

#include<qstring.h>
#include <QMainWindow>
#include<ActiveQt/QAxObject>
#include <QStandardPaths>
/*excel操作*/
enum EXcel_ColumnType{
    ColumnA = 1,
    ColumnB = 2,
    ColumnC = 3,
    ColumnD = 4,
    ColumnE = 5,
    ColumnF = 6,
    ColumnG = 7,
    ColumnH = 8,
    ColumnI = 9
};


class excel :public QWidget
{
public:
    excel();
    bool create_excel();
    QString saveas();
    void Excel_SetCell(QAxObject *worksheet,EXcel_ColumnType column,int row,QColor color,QString text);
    void castListListVariant2Variant(const QList<QList<QVariant> > &cells, QVariant &res);
    void castVariant2ListListVariant(const QVariant &var, QList<QList<QVariant> > &res);

};

#endif // EXCEL_H
