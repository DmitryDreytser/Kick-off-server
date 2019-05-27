/* 
   buildinc.js
   Скрипт для автоматического увеличения номера ревизии/сборки
   для проектов Visual C# (VS 2005 Express)
   файл должен содержать строку вида
   [assembly: AssemblyVersion("0.1.0.0")]
   эта строка ищется и изменяется с инкрементированием номера ревизии
*/

// чтение содержимого файла
function fnLoadContent(filename)
{
    fileObj = new ActiveXObject("Scripting.FileSystemObject");
    var ForReading = 1, ForWriting = 2, ForAppending = 8;
    var TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0;
    
    f = fileObj.GetFile(filename);
    ts = f.OpenAsTextStream(ForReading, TristateUseDefault);
    var s = ts.ReadAll();         // прочитать и вывести весь файл
    ts.Close();                   // закрыть текстовый поток
    return s;
}

// sContent - содержимое файла. В случае обнаружения искомой 
// строки оно изменяется на новое значение
function fnIncrementVerInfo(sContent, sAssemblyOrFile)
{
    var searchString = sAssemblyOrFile;
    var begin = sContent.indexOf(searchString);
    if (begin < 0)
        return sContent;
    begin += searchString.length;
    // найти конец строки с информацией о версии
    var end = sContent.indexOf("\")]", begin);
    if (end < 0)
        return sContent;
    var verMajor = 0;
    var verMinor = 0;
    var build = 0;
    var revision = 0;
    
    // эту часть текста мы вырежем и заменим новой информацией о версии
    var cutBegin = begin;
    var cutEnd = end;
    
    // копируем строку с версией
    var info = sContent.substring(begin, end);
    // разбираем
    
    // старший номер версии
    begin = 0;
    end = info.indexOf(".");
    if (end < 0)
        return sContent;
    verMajor = info.substring(begin, end);

    info = info.substr(end+1);

    // младший номер версии
    begin = 0;
    end = info.indexOf(".");
    if (end < 0)
        return sContent;
    verMinor = info.substring(begin, end);

    info = info.substr(end+1);
    
    // сборка
    begin = 0;
    end = info.indexOf(".");
    if (end < 0)
        return sContent;
    build = info.substring(begin, end);

    info = info.substr(end+1);
    
    // ревизия
    begin = 0;
    end = info.length;
    if (end < 0)
        return sContent;
    revision = info;

    // увеличиваем номер
    revision++;
    var info = verMajor + "." + verMinor + "." + build + "." + revision;
    
    // собираем новый файл
    var s = sContent.substring(0, cutBegin) + info + sContent.substring(cutEnd, sContent.length);
    return s;
}

// основная программа
function main()
{
    var args = WScript.Arguments;
		
    var sFileToParse = args.item(0) + "Properties\\AssemblyInfo.cs";
    // открыть файл
    var content = fnLoadContent(sFileToParse);
    // увеличить номер ревизии
    content = fnIncrementVerInfo(content, "[assembly: AssemblyVersion(\"");
    content = fnIncrementVerInfo(content, "[assembly: AssemblyFileVersion(\"");
    // Создать объект FileSystemObject.
    var myFileSysObj = new ActiveXObject("Scripting.FileSystemObject")
    // Создать объект TextStream.
    var myTextStream = myFileSysObj.OpenTextFile(sFileToParse, 2, true)
    // Записать в файл результат
    myTextStream.Write(content);
}

// пуск
main();