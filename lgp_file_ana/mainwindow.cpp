#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "fileapi.h"

#pragma execution_character_set("UTF-8")

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    tmp_excel = new excel();
    tmp_lgp = new lgp_ana();
    ui->progressBar->setVisible(false);

}

MainWindow::~MainWindow()
{
    delete ui;
}


bool MainWindow::produce_excel()
{
    std::vector<QString> vfile_name;
    QFileInfoList lst;;
    QString tmp_file_name = ui->dir_name->text();
    if (!GetSpecifiedFormatFiles(tmp_file_name, "", lst,"lgp"))
    {
        qDebug() << "没有找到任何 .lgp文件,请选择有.lgp文件存在的目录" ;
        QMessageBox::about(NULL,  "错误",  "所选文件夹没有找到lgp文件，请确定所选路径中没有中文");
        return false;
    }
    else
    {
        qDebug() << "成功找到文件";
    }
    for (QFileInfo data : lst)
    {
        //qDebug() << "data=" <<data.fileName();
        vfile_name.push_back(data.fileName());
    }
    int tmp_number = lst.size();
    qDebug() << tmp_number;
    QString templatePath = "./template.xlsx";
    QFileInfo info(templatePath);

    if(!info.exists())
    {
        qDebug()<<"template.xlsx is NULL";
        return 0;
    }

    templatePath = info.absoluteFilePath();                   //获取模板的绝对路径
    templatePath = QDir::toNativeSeparators(templatePath);   //转换一下路径,让windows能够识别

    QString ExcelFile = QDir::toNativeSeparators(tmp_excel->saveas());  //打开文件保存对话框,找到要保存的位置
    ui->excel_dir->setText(ExcelFile);
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
    excel->setProperty("Visible", false);//显示窗体看效果
    excel->setProperty("DisplayAlerts", false);//显示警告看效果
    QAxObject *workbooks = excel->querySubObject("WorkBooks");

    QAxObject* workbook = workbooks->querySubObject("Open(const QString&)",QDir::toNativeSeparators(ExcelFile) ); //打开

    excel->setProperty("Caption", "Qt Excel");      //标题为Qt Excel
    QAxObject *work_book = excel->querySubObject("ActiveWorkBook");

    QAxObject *worksheet = work_book->querySubObject("Sheets(int)",1);     //获取表单1

    int current_column = 2;

    //QMessageBox::about(NULL,  "生成中",  "生成excel文件中");
    ui->progressBar->setRange(0,tmp_number);
    ui->progressBar->setValue(0);
    ui->progressBar->setVisible(true);

    for(int i = 0 ;i < tmp_number ;i++)
    {
        LGP_DATA tmp_lpg_data;
        tmp_lpg_data = tmp_lgp->find_data(qstringtochar(ui->dir_name->text()+QString("/")+vfile_name[i]),qstringtochar(vfile_name[i]));
        tmp_excel->Excel_SetCell(worksheet,ColumnA,current_column,QColor(0,0,0),tmp_lpg_data.file_time);
        tmp_excel->Excel_SetCell(worksheet,ColumnB,current_column,QColor(0,0,0),tmp_lpg_data.acceleration);
        tmp_excel->Excel_SetCell(worksheet,ColumnC,current_column,QColor(0,0,0),tmp_lpg_data.change);
        tmp_excel->Excel_SetCell(worksheet,ColumnD,current_column,QColor(0,0,0),tmp_lpg_data.ptop);
        tmp_excel->Excel_SetCell(worksheet,ColumnE,current_column,QColor(0,0,0),tmp_lpg_data.frequency);
        current_column++;
        ui->progressBar->setValue(i+1);
    }
    ui->label_3->setText("生成excel文件成功...");
    workbook->dynamicCall("Save()" );
    workbook->dynamicCall("Close()");  //关闭文件
    excel->dynamicCall("Quit()");//关闭excel
    QMessageBox::about(NULL,  "成功",  "成功找到"+ QString::number(tmp_number)+"个lgp文件，已成功导出excel文件\n导出文件位于\n"+ExcelFile);
    ui->label_3->clear();
    ui->progressBar->setVisible(false);
    return true;

}
void MainWindow::on_pushButton_clicked()
{

    QString dirpath = QFileDialog::getExistingDirectory(this, "选择lgp文件所在目录", "./", QFileDialog::ShowDirsOnly);
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

/**
 * @brief 获取指定目录下特定格式的文件列表
 * @param dstDir: 目标文件夹
 * @param targetName: 文件名前缀,eg:"AutoUpdate"
 * @param list: 得到的文件列表
 * @param szSuffix: 文件后缀名
 */
bool MainWindow::GetSpecifiedFormatFiles(
        const QString & dstDir,
        const QString & targetName,
        QFileInfoList & list,
        QString suffix = "lgp")
{
    // 获取目录文件列表
    QDir dir(dstDir);
    dir.setFilter(QDir::Files | QDir::NoSymLinks);
    dir.setSorting(QDir::Name );

    QStringList filters;
    filters << QString("*.%1").arg(suffix);
    dir.setNameFilters(filters);

    QFileInfoList listTmp = dir.entryInfoList();
    foreach(QFileInfo item, listTmp)
    {
        //qDebug() << "item.absoluteFilePath()=" << item.absoluteFilePath();
        //qDebug() << "item.completeBaseName()=" << item.completeBaseName();
        if (targetName.toLower() == item.completeBaseName().left(targetName.length()).toLower())
        {
            list.append(item);
        }
    }
    qDebug() << listTmp[listTmp.size()-1].completeBaseName();
    return !list.isEmpty();
}
void MainWindow::listen_thread()
{

}

void MainWindow::on_pushButton_3_clicked()
{
    listen_dirpath = QFileDialog::getExistingDirectory(this, "选择lgp文件所在目录", "./", QFileDialog::ShowDirsOnly);
    ui->lineEdit->setText(listen_dirpath);

}

void MainWindow::on_pushButton_4_clicked()
{
    QString ExcelFile = QDir::toNativeSeparators(tmp_excel->saveas());  //打开文件保存对话框,找到要保存的位置
    ui->lineEdit_2->setText(ExcelFile);

}

void MainWindow::on_pushButton_5_clicked()
{
    int current_column = produce_excel_ano();
    tmp_thread = new mythread(listen_dirpath,ui->lineEdit_2->text(), current_column);
    tmp_thread->start();
    qDebug() << "线程开始" ;
}

void MainWindow::on_pushButton_6_clicked()
{


    tmp_thread->quit();
    //delete tmp_thread;
    qDebug() << "线程停止" ;


}
int MainWindow::produce_excel_ano()
{
    std::vector<QString> vfile_name;
    QFileInfoList lst;;
    QString tmp_file_name = ui->lineEdit->text();
    GetSpecifiedFormatFiles(tmp_file_name, "", lst,"lgp");
    for (QFileInfo data : lst)
    {
        //qDebug() << "data=" <<data.fileName();
        vfile_name.push_back(data.fileName());
    }
    int tmp_number = lst.size();
    qDebug() << tmp_number;
    QString templatePath = "./template.xlsx";
    QFileInfo info(templatePath);

    if(!info.exists())
    {
        qDebug()<<"template.xlsx is NULL";
        return 0;
    }

    templatePath = info.absoluteFilePath();                   //获取模板的绝对路径
    templatePath = QDir::toNativeSeparators(templatePath);   //转换一下路径,让windows能够识别

    QString ExcelFile = ui->lineEdit_2->text();  //打开文件保存对话框,找到要保存的位置
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

    int current_column = 2;

    //QMessageBox::about(NULL,  "生成中",  "生成excel文件中");
    ui->progressBar->setRange(0,tmp_number);
    ui->progressBar->setValue(0);
    ui->progressBar->setVisible(true);

    for(int i = 0 ;i < tmp_number ;i++)
    {
        LGP_DATA tmp_lpg_data;
        tmp_lpg_data = tmp_lgp->find_data(qstringtochar(ui->lineEdit->text()+QString("/")+vfile_name[i]),qstringtochar(vfile_name[i]));
        qDebug() << tmp_lpg_data.file_time;
        tmp_excel->Excel_SetCell(worksheet,ColumnA,current_column,QColor(0,0,0),tmp_lpg_data.file_time);
        tmp_excel->Excel_SetCell(worksheet,ColumnB,current_column,QColor(0,0,0),tmp_lpg_data.acceleration);
        tmp_excel->Excel_SetCell(worksheet,ColumnC,current_column,QColor(0,0,0),tmp_lpg_data.change);
        tmp_excel->Excel_SetCell(worksheet,ColumnD,current_column,QColor(0,0,0),tmp_lpg_data.ptop);
        tmp_excel->Excel_SetCell(worksheet,ColumnE,current_column,QColor(0,0,0),tmp_lpg_data.frequency);
        current_column++;
        ui->progressBar->setValue(i+1);
    }
    ui->label_3->setText("生成excel文件成功...");
    workbook->dynamicCall("Save()" );
    workbook->dynamicCall("Close()");  //关闭文件
    excel->dynamicCall("Quit()");//关闭excel
    QMessageBox::about(NULL,  "成功",  "成功找到"+ QString::number(tmp_number)+"个lgp文件，已成功导出excel文件\n导出文件位于\n"+ExcelFile);
    ui->label_3->clear();
    ui->progressBar->setVisible(false);
    return current_column;

}

