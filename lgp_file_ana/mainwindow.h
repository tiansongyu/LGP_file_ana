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

#include <Windows.h>

#include <fstream>
#include <iostream>
#include <string.h>
#include <iomanip>
#include <vector>
#include <string>
#include <ShlObj.h>

#include "lgp_ana.h"
QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    bool produce_excel();
    ~MainWindow();

private slots:
    void on_pushButton_clicked();
    void on_pushButton_2_clicked();
private:
    Ui::MainWindow *ui;
    excel *tmp_excel;
    lgp_ana *tmp_lgp;
    char* qstringtochar(QString qst);
    bool GetSpecifiedFormatFiles(
            const QString & dstDir,
            const QString & targetName,
            QFileInfoList & list,
            QString suffix);

};
#endif // MAINWINDOW_H
