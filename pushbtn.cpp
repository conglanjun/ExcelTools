#include "pushbtn.h"

void pushbtn::OnClicked()
{
//    QString str;
//    str = "You press " + this->text();
//    QMessageBox::information(this, tr("Information"), str);
//    excel e;
//    if(QString::compare(this->text(),"导出文件位置",Qt::CaseSensitive) == 0){
//        e.excelExport();
//    }else if(QString::compare(this->text(),"选择输入文件",Qt::CaseSensitive) == 0){
//        e.excelImport();
//    }
//    e.excelImport();

//    WorkThread *workThread = new WorkThread();
//    workThread->start();
    e->excelImport();

}

