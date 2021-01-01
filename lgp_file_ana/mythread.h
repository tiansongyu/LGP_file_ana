#ifndef MYTHREAD_H
#define MYTHREAD_H
#include "lgp_ana.h"
#include "excel.h"
#include <QThread>
class mythread: public QThread
{
    Q_OBJECT
public:
    mythread(QString _file_dir,QString excel_file,int _current_column);
    void produce_excel(QString vfile_name);
    char* qstringtochar(QString qst);
    excel *tmp_excel;
    QString excelfile;
protected:
    void run() override;

private:
    QString file_dir;
    int current_column;
    lgp_ana* tmp_lgp;
};

#endif // MYTHREAD_H
