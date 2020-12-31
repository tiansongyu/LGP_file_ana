#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "fileapi.h"


MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    tmp_excel = new excel();
    tmp_lgp = new lgp_ana();
}

MainWindow::~MainWindow()
{
    delete ui;
}


bool MainWindow::produce_excel()
{
    QString templatePath = "./template.xlsx";
       QFileInfo info(templatePath);

       if(!info.exists())
       {
               qDebug()<<"template.xlsx is NULL";
               return 0;
       }

       templatePath = info.absoluteFilePath();                   //获取模板的绝地路径
       templatePath = QDir::toNativeSeparators(templatePath);   //转换一下路径,让windows能够识别

       QString ExcelFile = QDir::toNativeSeparators(tmp_excel->saveas());  //打开文件保存对话框,找到要保存的位置

       if(ExcelFile=="")
               return  0;

       QFile::copy(templatePath, ExcelFile);                   //将模板文件复制到要保存的位置去

       info.setFile(ExcelFile);
       info.setFile(info.dir().path()+"/~$"+info.fileName());

       if(info.exists())          //判断一下,有没有"~$XXX.xlsx"文件存在,是不是为只读
       {
           qDebug()<<"报表属性为只读,请检查文件是否已打开!";
            return   0;
       }


       QAxObject *excel = new QAxObject();//建立excel操作对象
       excel->setControl("Excel.Application");//连接Excel控件
       excel->setProperty("Visible", true);//显示窗体看效果
       excel->setProperty("DisplayAlerts", false);//显示警告看效果
       QAxObject *workbooks = excel->querySubObject("WorkBooks");

       QAxObject* workbook = workbooks->querySubObject("Open(const QString&)",QDir::toNativeSeparators(ExcelFile) ); //打开

       excel->setProperty("Caption", "Qt Excel");      //标题为Qt Excel
       QAxObject *work_book = excel->querySubObject("ActiveWorkBook");

       QAxObject *worksheet = work_book->querySubObject("Sheets(int)",1);     //获取表单1

       tmp_excel->Excel_SetCell(worksheet,ColumnB,2,QColor(74,51,255),"12345");     //设置B2单元格内容为12345

       tmp_excel->Excel_SetCell(worksheet,ColumnB,3,QColor(255,255,0),"B3");     //设置B3单元格内容

       tmp_excel->Excel_SetCell(worksheet,ColumnB,4,QColor(255,0,0),"B4");     //设置B4单元格内容
       /////////////////////////////////////
       /// \brief file_number
       ///	std::cout << "选择.lgp文件所在目录" << std::endl;
       std::vector<QString> vfile_name;
       int file_number = 0;
           WIN32_FIND_DATA p;
           char* d = new char[40];
           sprintf(d, "%s/*.lgp", qstringtochar(ui->dir_name->text()));
           const size_t cSize = strlen(d)+1;
           wchar_t* wc = new wchar_t[cSize];
           mbstowcs (wc, d, cSize);
           qDebug() << QString("%1").arg(d);
           HANDLE h = FindFirstFile(wc, &p);
           if (h == INVALID_HANDLE_VALUE)
           {
                qDebug() << "没有找到任何 .lgp文件,请选择有.lgp文件存在的目录" ;
               system("pause");
               return 0;
           }
           //std::cout << "从" << dir << "/中找到如下文件" << std::endl;
           //puts(p.cFileName);
           vfile_name.push_back(QString::fromWCharArray(p.cFileName));
           file_number++;
           while (FindNextFile(h, &p))
           {
               //puts(p.cFileName);
               vfile_name.push_back(QString::fromWCharArray(p.cFileName));
               file_number++;
           }
           //std::cout << "共找到" << file_number << "个文件" << std::endl << std::endl;
           int tmp_number = file_number;

           /////////////////////////////////////

       /*批量一次性设置A6~I106所在内容*/
       //tmp_excel->Excel_SetCell(worksheet,ColumnB,4,QColor(255,0,0),"B4");     //设置B4单元格内容
       int current_column = 2;

       for(int i = 0 ;i < tmp_number ;i++)
       {
           LGP_DATA tmp_lpg_data;
           qDebug() << vfile_name[i];
           tmp_lpg_data = tmp_lgp->find_data(qstringtochar(ui->dir_name->text()+QString("/")+vfile_name[i]));
           tmp_excel->Excel_SetCell(worksheet,ColumnA,current_column,QColor(255,0,0),tmp_lpg_data.file_time);
           tmp_excel->Excel_SetCell(worksheet,ColumnB,current_column,QColor(255,0,0),tmp_lpg_data.acceleration);
           tmp_excel->Excel_SetCell(worksheet,ColumnC,current_column,QColor(255,0,0),tmp_lpg_data.change);
           tmp_excel->Excel_SetCell(worksheet,ColumnD,current_column,QColor(255,0,0),tmp_lpg_data.ptop);
           tmp_excel->Excel_SetCell(worksheet,ColumnE,current_column,QColor(255,0,0),tmp_lpg_data.frequency);
           current_column++;
       }


       QAxObject *user_range = worksheet->querySubObject("Range(const QString&)","A6:I106");

       QList<QList<QVariant> > datas;
       for(int i=1;i<101;i++)
       {
           QList<QVariant> rows;
           for(int j=1;j<10;j++)
           {
               rows.append(i*j);
           }
           datas.append(rows);
       }

       QVariant var;
       tmp_excel->castListListVariant2Variant(datas,var);

       user_range->setProperty("Value", var);


       workbook->dynamicCall("Save()" );


        workbook->dynamicCall("Close()");  //关闭文件
        excel->dynamicCall("Quit()");//关闭excel

}
void MainWindow::on_pushButton_clicked()
{

    QString dirpath = QFileDialog::getExistingDirectory(this, "选择目录", "./", QFileDialog::ShowDirsOnly);
    ui->dir_name->setText(dirpath);

}

void MainWindow::on_pushButton_2_clicked()
{
     produce_excel();
}
char* MainWindow::qstringtochar(QString qst)
{
    QString str1 = qst;
    QByteArray ba = str1.toLocal8Bit();
    char *c_str2 = ba.data();
    return c_str2;
}
