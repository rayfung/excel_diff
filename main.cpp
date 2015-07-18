#include <QCoreApplication>
#include <QString>
#include <QList>
#include <QFile>
#include <cstdio>
#include "xlsxdocument.h"

int GetWidth(QString s)
{
    int width = 0;

    for(int i = 0; i < s.length(); ++i)
    {
        if(s[i].unicode() < 0x2000)
            width++;
        else
            width += 2;
    }
    return width;
}

QString convertExcel(const QString &path)
{
    QXlsx::Document doc(path);
    QString s;
    QXlsx::CellRange range = doc.dimension();
    QList<QList<QString>> stringTable;
    QList<int> columnWidth;
    int i, j;

    for(i = 0; i < range.lastRow(); ++i)
    {
        stringTable.append(QList<QString>());

        for(j = 0; j < range.lastColumn(); ++j)
        {
            QVariant var = doc.read(i + 1, j + 1);
            QString item;
            if(var.type() == QVariant::Double)
                item = QString::number(var.toDouble());
            else
                item = var.toString();

            stringTable[i].append(item);
            if(j >= columnWidth.length())
                columnWidth.append(GetWidth(item));
            else
            {
                int width = GetWidth(item);
                if(width > columnWidth.at(j))
                    columnWidth[j] = width;
            }
        }
    }

    for(i = 0; i < stringTable.length(); ++i)
    {
        for(j = 0; j < stringTable.at(i).length(); ++j)
        {
            int width = GetWidth(stringTable.at(i).at(j));

            s += stringTable.at(i).at(j);
            s += QString(columnWidth.at(j) - width + 2, ' ');
        }
        s += "\n";
    }
    return s;
}

int main(int argc, char *argv[])
{
    for(int i = 1; i < argc; ++i)
    {
        QString path = QString::fromLocal8Bit(argv[i]);

        printf("%s", convertExcel(path).toUtf8().constData());
    }
    return 0;
}
