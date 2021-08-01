#include <QApplication>
#include <QFileInfo>
#include <QFile>
#include <QDebug>
#include <QString>
#include <QDir>
#include <QFileDialog>
#include <QFileInfoList>
#include <QList>
#include <ActiveQt/QAxObject>
#include <QObject>
#include <QStringList>
#include <QThread>

#pragma execution_character_set("utf-8")

//遍历一个文件夹下的所有docx文件，包含子目录中的docx文件
void getDocFile(QFileInfoList &list, QDir dir)
{
    QFileInfoList files = dir.entryInfoList(QDir::Filter::NoDotAndDotDot | QDir::Filter::AllEntries);
    for (int i = 0; i < files.size(); ++i)
    {
        QFileInfo info = files[i];
        if (info.isDir())
            getDocFile(list, info.absoluteFilePath());
        else if (info.absoluteFilePath().endsWith(".docx"))
            list << info;
    }
}
int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    QFileInfoList filelist;
    //只能选取文件夹！就算只有一个docx文件，也只能选取文件夹，比如 /test/ 目录下有个a.docx, 那你就要选择/test目录
    getDocFile(filelist, QFileDialog::getExistingDirectory(nullptr, QString("选择文件夹"), ""));

    QAxObject *wordApplication = new QAxObject("Word.Application");   //新建一个word应用
    QAxObject *wordDoc = wordApplication->querySubObject("Documents");    //并设置为处理文档模式
    QAxObject* doc = nullptr;   //后续用来执行读取、转换动作
    QString OutputFileName;   //存放生成的pdf路径

    for (int i = 0; i < filelist.size(); ++i)
    {
        //读取一个docx文件
        doc = wordDoc->querySubObject("Open(string,bool,bool)",filelist.at(i).absoluteFilePath(),false, true);
        //下一步要生成的pdf放到当前路径
        OutputFileName = static_cast<QString>(filelist.at(i).absoluteFilePath()).replace(".docx", ".pdf");
        qDebug() << i + 1 << OutputFileName;
        //开始转换pdf
        doc->querySubObject("ExportAsFixedFormat(string, int ,bool)", OutputFileName, 17, false);
        doc->dynamicCall("Close(bool)", false);
    }
    qDebug() << "全部完成";
    return a.exec();
}
