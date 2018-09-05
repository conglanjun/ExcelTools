#include "mainwindow.h"
#include <QApplication>
#include <QPushButton>
#include <QPushButton>
#include "pushBtn.h"
#include "excel.h"

#include <QDialog>
#include <QRect>
#include <QFont>
#include <QLineEdit>
#include <QGridLayout>
#include <QProgressBar>

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
//    MainWindow w;

    QDialog *mainWindow = new QDialog;
    mainWindow->resize(300, 150);
    QSizePolicy sizePolicy(QSizePolicy::Fixed, QSizePolicy::Fixed);
    sizePolicy.setHorizontalStretch(0);
    sizePolicy.setVerticalStretch(0);
    sizePolicy.setHeightForWidth(mainWindow->sizePolicy().hasHeightForWidth());
    mainWindow->setSizePolicy(sizePolicy);
    mainWindow->setMinimumSize(QSize(400, 300));
    mainWindow->setMaximumSize(QSize(400, 300));
    mainWindow->setSizeGripEnabled(false);

    QGridLayout *mainLayout = new QGridLayout(mainWindow);

    QProgressBar *progressBar = new QProgressBar();

    excel *e = new excel(progressBar);
    pushbtn btn(e, mainWindow);
    btn.setText("选择输入文件");
    btn.setGeometry(QRect(20, 40, 80, 40));

    mainLayout->addWidget(&btn, 0, 0);
    mainLayout->addWidget(progressBar, 0, 1);

    mainWindow->setWindowTitle("财务助手");
    mainWindow->show();

    return a.exec();
}
