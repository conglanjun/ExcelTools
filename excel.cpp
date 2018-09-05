#include "excel.h"
#include "workthread.h"
void excel::excelImportDemo()
{
    QFileDialog *fileDialog = new QFileDialog(this);//创建一个QFileDialog对象，构造函数中的参数可以有所添加。
    fileDialog->setWindowTitle(tr("导入文件"));//设置文件保存对话框的标题
//    fileDialog->setAcceptMode(QFileDialog::AcceptSave);//设置文件对话框为保存模式
    fileDialog->setFileMode(QFileDialog::AnyFile);//设置文件对话框弹出的时候显示任何文件，不论是文件夹还是文件
    fileDialog->setViewMode(QFileDialog::Detail);//文件以详细的形式显示，显示文件名，大小，创建日期等信息；
    if(fileDialog->exec() == QDialog::Accepted) {//注意使用的是QFileDialog::Accepted或者QDialog::Accepted,不是QFileDialog::Accept
        QString path = fileDialog->selectedFiles()[0];//得到用户选择的文件名
        QMessageBox::information(this, tr("Information"), path);
        QAxObject excel("Excel.Application");
        excel.setProperty("Visible", true);
        QAxObject *work_books = excel.querySubObject("WorkBooks");
        work_books->dynamicCall("Open(const QString&)", path);
        excel.setProperty("Caption", "Qt Excel");
        QAxObject *work_book = excel.querySubObject("ActiveWorkBook");
        QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
        //删除工作表（删除第一个）
//        QAxObject *first_sheet = work_sheets->querySubObject("Item(int)", 1);
//        first_sheet->dynamicCall("delete");
        //插入工作表（插入至最后一行）
        int sheet_count = work_sheets->property("Count").toInt();
        QAxObject *last_sheet = work_sheets->querySubObject("Item(int)", sheet_count);
        QAxObject *work_sheet = work_sheets->querySubObject("Add(QVariant)", last_sheet->asVariant());
        last_sheet->dynamicCall("Move(QVariant)", work_sheet->asVariant());

        work_sheet->setProperty("Name", "Qt Sheet");  //设置工作表名称
        //操作单元格（第2行第2列）
        QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", 2, 2);
        cell->setProperty("Value", "Java C++ C# PHP Perl Python Delphi Ruby");  //设置单元格值
        cell->setProperty("RowHeight", 50);  //设置单元格行高
        cell->setProperty("ColumnWidth", 30);  //设置单元格列宽
        cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152
        cell->setProperty("VerticalAlignment", -4108);  //上对齐（xlTop）-4160 居中（xlCenter）：-4108  下对齐（xlBottom）：-4107
        cell->setProperty("WrapText", true);  //内容过多，自动换行
        //cell->dynamicCall("ClearContents()");  //清空单元格内容
        QAxObject* interior = cell->querySubObject("Interior");
        interior->setProperty("Color", QColor(0, 255, 0));   //设置单元格背景色（绿色）

        QAxObject* border = cell->querySubObject("Borders");
        border->setProperty("Color", QColor(0, 0, 255));   //设置单元格边框色（蓝色）

        QAxObject *font = cell->querySubObject("Font");  //获取单元格字体
        font->setProperty("Name", QStringLiteral("华文彩云"));  //设置单元格字体
        font->setProperty("Bold", true);  //设置单元格字体加粗
        font->setProperty("Size", 20);  //设置单元格字体大小
        font->setProperty("Italic", true);  //设置单元格字体斜体
        font->setProperty("Underline", 2);  //设置单元格下划线
        font->setProperty("Color", QColor(255, 0, 0));  //设置单元格字体颜色（红色）

        //设置单元格内容，并合并单元格（第5行第3列-第8行第5列）
        QAxObject *cell_5_6 = work_sheet->querySubObject("Cells(int,int)", 5, 3);
        cell_5_6->setProperty("Value", "Java");  //设置单元格值
        QAxObject *cell_8_5 = work_sheet->querySubObject("Cells(int,int)", 8, 5);
        cell_8_5->setProperty("Value", "C++");

        QString merge_cell;
        merge_cell.append(QChar(3 - 1 + 'A'));  //初始列
        merge_cell.append(QString::number(5));  //初始行
        merge_cell.append(":");
        merge_cell.append(QChar(5 - 1 + 'A'));  //终止列
        merge_cell.append(QString::number(8));  //终止行
        QAxObject *merge_range = work_sheet->querySubObject("Range(const QString&)", merge_cell);
        merge_range->setProperty("HorizontalAlignment", -4108);
        merge_range->setProperty("VerticalAlignment", -4108);
        merge_range->setProperty("WrapText", true);
        merge_range->setProperty("MergeCells", true);  //合并单元格
        //merge_range->setProperty("MergeCells", false);  //拆分单元格

        work_book->dynamicCall("Save()");  //保存文件（为了对比test与下面的test2文件，这里不做保存操作） work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
        work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
        excel.dynamicCall("Quit(void)");  //退出

//        filePath=listWidget_File->item(listWidget_File->currentRow())->text();//这个是得到在ListWidget中点击查看的图片，得到这个图片的名字
//        filePath=QString("/media/sd/PICTURES")+filePath;//将路径和文件名连接起来
//        QImage iim(filePath);//创建一个图片对象,存储源图片
//        QPainter painter(&iim);//设置绘画设备
//        QFile file(path);//创建一个文件对象，存储用户选择的文件
//        if (!file.open(QIODevice::ReadWrite)){以只读的方式打开用户选择的文件，如果失败则返回
//            return;
    }
}

void excel::excelImport()
{
    QFileDialog *fileDialog = new QFileDialog(this);//创建一个QFileDialog对象，构造函数中的参数可以有所添加。
    fileDialog->setWindowTitle(tr("导入文件"));//设置文件保存对话框的标题
//    fileDialog->setAcceptMode(QFileDialog::AcceptSave);//设置文件对话框为保存模式
    fileDialog->setFileMode(QFileDialog::AnyFile);//设置文件对话框弹出的时候显示任何文件，不论是文件夹还是文件
    fileDialog->setViewMode(QFileDialog::Detail);//文件以详细的形式显示，显示文件名，大小，创建日期等信息；
    if(fileDialog->exec() == QDialog::Accepted) {//注意使用的是QFileDialog::Accepted或者QDialog::Accepted,不是QFileDialog::Accept

        const QString path = fileDialog->selectedFiles()[0];//得到用户选择的文件名

        wt = new WorkThread(path);
//        connect(wt, SIGNAL(send_export_signal()),this,SLOT(send_cmd()));
        connect(wt, &WorkThread::send_export_signal, this, &excel::send_cmd);
        connect(wt, &WorkThread::send_excel_row_count, this, &excel::receive_row_count);
        connect(wt, &WorkThread::send_excel_row_done, this, &excel::receive_row_done);


//        qpDialog = new QProgressDialog(this);
//        qpDialog->setWindowModality(Qt::WindowModal);
//        qpDialog->setMinimumDuration(0);
//        qpDialog->setAttribute(Qt::WA_DeleteOnClose, true);
//        qpDialog->setWindowTitle(tr("处理文件"));
//        qpDialog->setLabelText(tr("正在处理......"));
//        qpDialog->setCancelButtonText(tr("取消"));

        wt->start();
//        btn->setEnabled(false);
        // 进度条

//        qpDialog->setRange(0,80000);
//        for(int i=1;i<80000;i++)//注意这里是从1开始的
//        {
//          //do something;
//          qpDialog->setValue(i);
//          if(qpDialog->wasCanceled())
//          {
//             //do something;
//          }
//        }
//        qpDialog->close();
    }
}

void excel::send_cmd(QString path)
{
    QMessageBox::information(this, tr("Information"), "完事啦，去：" + path + "，查看文件吧！");
    progressBar->setValue(0);
//    btn->setEnabled(true);
}

void excel::receive_row_count(int row_count)
{
    this->row_count = row_count;
//    qpDialog->setRange(0, row_count);
//    qpDialog->setValue(1);

    // bar
    progressBar->setRange(0, row_count);
    progressBar->setValue(1);
}

void excel::receive_row_done()
{
    int currentValue = progressBar->value();
    if(currentValue < this->row_count){
//        qpDialog->setValue(++currentValue);
        // bar
        progressBar->setValue(++currentValue);
    }
}

void excel::setCellValue(QAxObject *work_sheet, int row, QAxObject *data_sheet, int data_row, bool isDouble, int index)
{
    int i;
    for(i = 1; i < 9; ++i)
    {
        QAxObject *cell = work_sheet->querySubObject("Cells(int,int)", row, i);
        if(row == 1)
        {
            float column_length = 13.86;
            switch(i){
                case 1:
                    column_length = 26.57;
                    break;
                case 2:
                    column_length = 8.43;
                    break;
                case 3:
                    column_length = 18.57;
                    break;
                case 4:
                    column_length = 14.29;
                    break;
                case 5:
                    column_length = 15.86;
                    break;
                case 6:
                    column_length = 12.71;
                    break;
            }
            cell->setProperty("ColumnWidth", (int)column_length);  //设置单元格列宽
            cell->setProperty("HorizontalAlignment", -4108); //左对齐（xlLeft）：-4131  居中（xlCenter）：-4108  右对齐（xlRight）：-4152
            cell->setProperty("VerticalAlignment", -4108);  //上对齐（xlTop）-4160 居中（xlCenter）：-4108  下对齐（xlBottom）：-4107
            QAxObject *font = cell->querySubObject("Font");  //获取单元格字体
            font->setProperty("Bold", true);  //设置单元格字体加粗
        }
        QString data = data_sheet->querySubObject("Cells(int,int)", data_row, i)->dynamicCall("Value2()").toString();
        if(i == 2 || i == 4 || i == 5){
            cell->setProperty("Value", "'" + data);  //设置单元格值
        }else{
            cell->setProperty("Value", data);  //设置单元格值
        }
        QAxObject *font = cell->querySubObject("Font");  //获取单元格字体
        if(isDouble)
        {
            switch(index % 5){
                case 0:
                    font->setProperty("Color", QColor(0, 255, 0));  //设置单元格字体颜色（绿色）
                    break;
                case 1:
                    font->setProperty("Color", QColor(0, 127, 255));  //设置单元格字体颜色（淡蓝）
                    break;
                case 2:
                    font->setProperty("Color", QColor(184, 115, 51));  //设置单元格字体颜色（铜色）
                    break;
                case 3:
                    font->setProperty("Color", QColor(107, 35, 142));  //设置单元格字体颜色（深石板蓝）
                    break;
                case 4:
                    font->setProperty("Color", QColor(255, 36, 0));  //设置单元格字体颜色（橙红色）
                    break;
                default:
                    break;
            }
        }else{
            font->setProperty("Color", QColor(0, 0, 0));
        }
    }
}

void excel::excelExport()
{
    QString filepath=QFileDialog::getSaveFileName(this,QObject::tr("保存路径"),".",QObject::tr("Microsoft Office 2007 (*.xlsx)"));//获取保存路径
    if(!filepath.isEmpty()){
        QAxObject *excel = new QAxObject(this);
        excel->setControl("Excel.Application");//连接Excel控件
        excel->dynamicCall("SetVisible (bool Visible)","false");//不显示窗体
        excel->setProperty("DisplayAlerts", false);//不显示任何警告信息。如果为true那么在关闭是会出现类似“文件已修改，是否保存”的提示

        QAxObject *workbooks = excel->querySubObject("WorkBooks");//获取工作簿集合
        workbooks->dynamicCall("Add");//新建一个工作簿
        QAxObject *workbook = excel->querySubObject("ActiveWorkBook");//获取当前工作簿
        QAxObject *worksheets = workbook->querySubObject("Sheets");//获取工作表集合
        QAxObject *worksheet = worksheets->querySubObject("Item(int)",1);//获取工作表集合的工作表1，即sheet1

        QAxObject *cellA,*cellB,*cellC,*cellD;

        //设置标题
        int cellrow=1;
        QString A="A"+QString::number(cellrow);//设置要操作的单元格，如A1
        QString B="B"+QString::number(cellrow);
        QString C="C"+QString::number(cellrow);
        QString D="D"+QString::number(cellrow);
        cellA = worksheet->querySubObject("Range(QVariant, QVariant)",A);//获取单元格
        cellB = worksheet->querySubObject("Range(QVariant, QVariant)",B);
        cellC=worksheet->querySubObject("Range(QVariant, QVariant)",C);
        cellD=worksheet->querySubObject("Range(QVariant, QVariant)",D);
        cellA->dynamicCall("SetValue(const QVariant&)",QVariant("流水号"));//设置单元格的值
        cellB->dynamicCall("SetValue(const QVariant&)",QVariant("用户名"));
        cellC->dynamicCall("SetValue(const QVariant&)",QVariant("金额"));
        cellD->dynamicCall("SetValue(const QVariant&)",QVariant("日期"));
        cellrow++;

//        int rows=this->model->rowCount();
//        for(int i=0;i<rows;i++){
//            QString A="A"+QString::number(cellrow);//设置要操作的单元格，如A1
//            QString B="B"+QString::number(cellrow);
//            QString C="C"+QString::number(cellrow);
//            QString D="D"+QString::number(cellrow);
//            cellA = worksheet->querySubObject("Range(QVariant, QVariant)",A);//获取单元格
//            cellB = worksheet->querySubObject("Range(QVariant, QVariant)",B);
//            cellC=worksheet->querySubObject("Range(QVariant, QVariant)",C);
//            cellD=worksheet->querySubObject("Range(QVariant, QVariant)",D);
//            cellA->dynamicCall("SetValue(const QVariant&)",QVariant(this->model->item(i,0)->data(Qt::DisplayRole).toString()));//设置单元格的值
//            cellB->dynamicCall("SetValue(const QVariant&)",QVariant(this->model->item(i,1)->data(Qt::DisplayRole).toString()));
//            cellC->dynamicCall("SetValue(const QVariant&)",QVariant(this->model->item(i,2)->data(Qt::DisplayRole).toString()));
//            cellD->dynamicCall("SetValue(const QVariant&)",QVariant(this->model->item(i,3)->data(Qt::DisplayRole).toString()));
//            cellrow++;
//        }

        workbook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(filepath));//保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
        workbook->dynamicCall("Close()");//关闭工作簿
        excel->dynamicCall("Quit()");//关闭excel
        delete excel;
        excel=NULL;
    }
}
