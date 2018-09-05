#ifndef PUSHBTN_H
#define PUSHBTN_H

#include <QPushButton>
#include <QMessageBox>
#include "excel.h"

class pushbtn : public QPushButton
{
public :
    void OnClicked();

public:
    pushbtn(excel *e_param, QWidget *parent = NULL):QPushButton(parent), e(e_param)
    {
        connect(this, &QPushButton::clicked, this, &OnClicked);
    }
private:
    excel *e;
};

#endif // PUSHBTN_H
