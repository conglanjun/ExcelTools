#include "workthread.h"
#include "qt_windows.h"
#include <QDebug>

void setCellValue(QAxObject *work_sheet, int row, QAxObject *data_sheet, int data_row, bool isDouble, int index)
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


void WorkThread::run ()
{
    CoInitializeEx(NULL, COINIT_MULTITHREADED);

    QString path = getPath();//得到用户选择的文件名
    QAxObject * excel = new QAxObject("Excel.Application");
//        excel.setProperty("Visible", false);
    QAxObject *work_books = excel->querySubObject("WorkBooks");
    work_books->dynamicCall("Open(const QString&)", path);
    excel->setProperty("Caption", "Qt Excel");
    QAxObject *work_book = excel->querySubObject("ActiveWorkBook");
    QAxObject *work_sheets = work_book->querySubObject("Sheets");  //Sheets也可换用WorkSheets
    //删除工作表（删除第一个）
//        QAxObject *first_sheet = work_sheets->querySubObject("Item(int)", 1);
//        first_sheet->dynamicCall("delete");

    // 获取发票元数据
    QAxObject *data_sheet = work_sheets->querySubObject("Item(int)", 2);
    // 排序发票号
    QAxObject *order_sheet = work_sheets->querySubObject("Item(int)", 3);
    // 计算发票号个数
    // 排序后的发票号
    int code_row = 1;// 记录发票号的行数，为了显示进度条
    int code_column = 1;
    while(true)
    {
        QAxObject *order_code_cell = order_sheet->querySubObject("Cells(int,int)", code_row, code_column);
        QString order_code = order_code_cell->dynamicCall("Value2()").toString();
        if(order_code.isEmpty()) break;
        ++code_row;
    }
    emit send_excel_row_count(code_row);

    //插入工作表（插入至最后一行）
    int sheet_count = work_sheets->property("Count").toInt();
    QAxObject *last_sheet = work_sheets->querySubObject("Item(int)", sheet_count);
    QAxObject *work_sheet = work_sheets->querySubObject("Add(QVariant)", last_sheet->asVariant());
    last_sheet->dynamicCall("Move(QVariant)", work_sheet->asVariant());

    work_sheet->setProperty("Name", "排序后的抵扣联登记");  //设置工作表名称



    setCellValue(work_sheet, 1, data_sheet, 1, false, 0);

    // 根据排序后的发票号循环拿出每个元素，去元数据的发票号中进行对比。
    // 记录已经访问过的行。避免重复读取数据。
    QVector<int> indexs;
    // 记录data与排序code匹配的集合
    QVector<int> patterns;
    // 记录word sheet中的行数。
    int word_row = 2;
    code_row = 1;
    code_column = 1;
    int index = 0;// 记录重复匹配数，为了用颜色区分。
    while(true)
    {
        // 操作单元格
        // 排序后的发票号
        QAxObject *order_code_cell = order_sheet->querySubObject("Cells(int,int)", code_row, code_column);
        QString order_code = order_code_cell->dynamicCall("Value2()").toString();
//            QMessageBox::information(this, tr("order_code info"), order_code);
        if(order_code.isEmpty()) break;
        patterns.clear();
        // 元数据发票号
        int data_row = 2;
        int data_colum = 4;
        while(true)
        {
            bool isExist = indexs.contains(data_row);
            if (isExist)
            {
                ++data_row;
                continue;
            }
            // 获取data中的值。
            QAxObject *data_code_cell = data_sheet->querySubObject("Cells(int,int)", data_row, data_colum);
            QString data_code = data_code_cell->dynamicCall("Value2()").toString();
            if(data_code.isEmpty()) break;
//                QMessageBox::information(this, tr("data_code info"), data_code);
            if(data_code.endsWith(order_code))
            {
                // 记录访问过的data sheet
                indexs.append(data_row);
                // 记录data中哪些行和排序code匹配
                patterns.append(data_row);
            }
            ++ data_row;
        }

        // set data值到work，如果set一条数据，则行数据不变色。否则设置成其他颜色。
        bool isDouble = patterns.size() > 1 ? true : false;
        if(isDouble) ++index;
        for(int i = 0; i < patterns.size(); ++i)
        {
            setCellValue(work_sheet, word_row, data_sheet, patterns.at(i), isDouble, index);
            ++word_row;
        }
        ++code_row;
        emit send_excel_row_done();
    }
    work_book->dynamicCall("Save()");  //保存文件（为了对比test与下面的test2文件，这里不做保存操作） work_book->dynamicCall("SaveAs(const QString&)", "E:\\test2.xlsx");  //另存为另一个文件
    work_book->dynamicCall("Close(Boolean)", false);  //关闭文件
    excel->dynamicCall("Quit(void)");  //退出
    emit send_export_signal(m_path);
//    send();

//    delete excel;

//    QMessageBox::information(getExcel(), tr("Information"), "完事啦，去：" + path + "，查看文件吧！");
}

//void WorkThread::send()
//{
//    emit send_export_signal();
//}

