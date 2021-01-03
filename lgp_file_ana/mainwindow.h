#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "excel.h"
#include <qdebug.h>
#include<ActiveQt/QAxObject>
#include <QFileDialog>
#include <QString>
#include <qmessagebox.h>

#include <QtCore/QCoreApplication>
#include <QFileInfoList>
#include <QDir>
#include <QDebug>
#include <QDesktopServices>

#include <Windows.h>

#include <fstream>
#include <iostream>
#include <string.h>
#include <iomanip>
#include <vector>
#include <string>
#include <ShlObj.h>

#include "mytype.h"
#include "lgp_ana.h"
#include "mythread.h"
QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    Ui::MainWindow *ui;
    excel *tmp_excel;
    lgp_ana *tmp_lgp;
    char* qstringtochar(QString qst);
    ERROR_CODE GetSpecifiedFormatFiles(
            const QString & dstDir,
            const QString & targetName,
            QFileInfoList & list,
            QString suffix);
    void listen_thread();
    bool produce_excel();
    int produce_excel_ano();

    mythread* tmp_thread;
    QString listen_dirpath;
private slots:
    void on_pushButton_clicked();
    void on_pushButton_2_clicked();
    void on_pushButton_3_clicked();

    void on_pushButton_4_clicked();

    void on_pushButton_5_clicked();

    void on_pushButton_6_clicked();

    void on_pushButton_7_clicked();

private:


};
#endif // MAINWINDOW_H
