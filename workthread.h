#ifndef WORKTHREAD_H
#define WORKTHREAD_H
#include <QThread>
#include "excel.h"
class WorkThread : public QThread
{
    Q_OBJECT
public:
    WorkThread(const QString & path) : m_path(path){}

    QString getPath()
    {
        return m_path;
    }
//自定义信号
signals:
    void send_export_signal(QString path);
    void send_excel_row_done();
    void send_excel_row_count(int row_count);
    void send_btn_enable(bool flag);

protected:
    void run();
private:
    QString m_path;
};
#endif // WORKTHREAD_H
//class Newspaper : public QThread
//{
//    Q_OBJECT
//public :
//    Newspaper(const QString & name) : m_name(name){}
//    void send()
//    {
//        emit newPaper(m_name);
//    }
//signals:
//    void newPaper(const QString &name);

//private :
//    QString m_name;
//};
