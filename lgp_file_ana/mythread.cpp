#include "mythread.h"
#include "mainwindow.h"
mythread::mythread(QString _file_dir,QString excel_file,int _current_column)
{
    file_dir = _file_dir;
    current_column = _current_column;
    tmp_lgp = new lgp_ana();
    excelfile = excel_file;
}
//excelfile
void mythread::produce_excel(QString vfile_name)
{
    QAxObject *excel = new QAxObject();//建立excel操作对象
    excel->setControl("Excel.Application");//连接Excel控件
    excel->setProperty("Visible", true);//显示窗体看效果
    excel->setProperty("DisplayAlerts", false);//显示警告看效果
    QAxObject *workbooks = excel->querySubObject("WorkBooks");

    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)",QDir::toNativeSeparators(excelfile)); //打开

    excel->setProperty("Caption", "Qt Excel");      //标题为Qt Excel
    QAxObject *work_book = excel->querySubObject("ActiveWorkBook");

    QAxObject *worksheet = work_book->querySubObject("Sheets(int)",1);     //获取表单1


    //QMessageBox::about(NULL,  "生成中",  "生成excel文件中");

    for(int i = 0 ;i < 1 ;i++)
    {
        LGP_DATA tmp_lpg_data;
        tmp_lpg_data = tmp_lgp->find_data(qstringtochar(file_dir+QString("/")+vfile_name),qstringtochar(vfile_name));
        qDebug() << tmp_lpg_data.file_time;
        tmp_excel->Excel_SetCell(worksheet,ColumnA,current_column,QColor(0,0,0),tmp_lpg_data.file_time);
        tmp_excel->Excel_SetCell(worksheet,ColumnB,current_column,QColor(0,0,0),tmp_lpg_data.acceleration);
        tmp_excel->Excel_SetCell(worksheet,ColumnC,current_column,QColor(0,0,0),tmp_lpg_data.change);
        tmp_excel->Excel_SetCell(worksheet,ColumnD,current_column,QColor(0,0,0),tmp_lpg_data.ptop);
        tmp_excel->Excel_SetCell(worksheet,ColumnE,current_column,QColor(0,0,0),tmp_lpg_data.frequency);
        current_column++;
    }
    workbook->dynamicCall("Save()" );
    workbook->dynamicCall("Close()");  //关闭文件
    excel->dynamicCall("Quit()");//关闭excel
    //QMessageBox::about(NULL,  "成功",  "成功找到"+ QString::number(tmp_number)+"个lgp文件，已成功导出excel文件\n导出文件位于\n"+ExcelFile);

}
char* mythread::qstringtochar(QString qst)
{
    QString str1 = qst;
    QByteArray ba = str1.toLocal8Bit();
    char *c_str2 = ba.data();
    return c_str2;
}

void mythread::run()
{
    QFileInfoList last_dir_last;
    QString suffix("lgp");
    QString targetName("");
    QString new_lgp_file;

    // 获取目录文件列表
    QDir dir(file_dir);
    dir.setFilter(QDir::Files | QDir::NoSymLinks);
    dir.setSorting(QDir::Time | QDir::Reversed);

    QStringList filters;
    filters << QString("*.%1").arg(suffix);
    dir.setNameFilters(filters);
    last_dir_last = dir.entryInfoList();
    while(1)
    {
        sleep(2);

        // 获取目录文件列表
        QDir dir(file_dir);
        dir.setFilter(QDir::Files | QDir::NoSymLinks);
        dir.setSorting(QDir::Time | QDir::Reversed);

        QStringList filters;
        filters << QString("*.%1").arg(suffix);
        dir.setNameFilters(filters);

        QFileInfoList listTmp = dir.entryInfoList();
        if(listTmp != last_dir_last)
        {
            new_lgp_file = listTmp[listTmp.size()-1].absoluteFilePath();
            qDebug() <<  new_lgp_file ;
            produce_excel(listTmp[listTmp.size()-1].fileName());
        }
        last_dir_last = listTmp;
    }
}
