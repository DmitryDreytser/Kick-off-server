/* 
   buildinc.js
   ������ ��� ��������������� ���������� ������ �������/������
   ��� �������� Visual C# (VS 2005 Express)
   ���� ������ ��������� ������ ����
   [assembly: AssemblyVersion("0.1.0.0")]
   ��� ������ ������ � ���������� � ������������������ ������ �������
*/

// ������ ����������� �����
function fnLoadContent(filename)
{
    fileObj = new ActiveXObject("Scripting.FileSystemObject");
    var ForReading = 1, ForWriting = 2, ForAppending = 8;
    var TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0;
    
    f = fileObj.GetFile(filename);
    ts = f.OpenAsTextStream(ForReading, TristateUseDefault);
    var s = ts.ReadAll();         // ��������� � ������� ���� ����
    ts.Close();                   // ������� ��������� �����
    return s;
}

// sContent - ���������� �����. � ������ ����������� ������� 
// ������ ��� ���������� �� ����� ��������
function fnIncrementVerInfo(sContent, sAssemblyOrFile)
{
    var searchString = sAssemblyOrFile;
    var begin = sContent.indexOf(searchString);
    if (begin < 0)
        return sContent;
    begin += searchString.length;
    // ����� ����� ������ � ����������� � ������
    var end = sContent.indexOf("\")]", begin);
    if (end < 0)
        return sContent;
    var verMajor = 0;
    var verMinor = 0;
    var build = 0;
    var revision = 0;
    
    // ��� ����� ������ �� ������� � ������� ����� ����������� � ������
    var cutBegin = begin;
    var cutEnd = end;
    
    // �������� ������ � �������
    var info = sContent.substring(begin, end);
    // ���������
    
    // ������� ����� ������
    begin = 0;
    end = info.indexOf(".");
    if (end < 0)
        return sContent;
    verMajor = info.substring(begin, end);

    info = info.substr(end+1);

    // ������� ����� ������
    begin = 0;
    end = info.indexOf(".");
    if (end < 0)
        return sContent;
    verMinor = info.substring(begin, end);

    info = info.substr(end+1);
    
    // ������
    begin = 0;
    end = info.indexOf(".");
    if (end < 0)
        return sContent;
    build = info.substring(begin, end);

    info = info.substr(end+1);
    
    // �������
    begin = 0;
    end = info.length;
    if (end < 0)
        return sContent;
    revision = info;

    // ����������� �����
    revision++;
    var info = verMajor + "." + verMinor + "." + build + "." + revision;
    
    // �������� ����� ����
    var s = sContent.substring(0, cutBegin) + info + sContent.substring(cutEnd, sContent.length);
    return s;
}

// �������� ���������
function main()
{
    var args = WScript.Arguments;
		
    var sFileToParse = args.item(0) + "Properties\\AssemblyInfo.cs";
    // ������� ����
    var content = fnLoadContent(sFileToParse);
    // ��������� ����� �������
    content = fnIncrementVerInfo(content, "[assembly: AssemblyVersion(\"");
    content = fnIncrementVerInfo(content, "[assembly: AssemblyFileVersion(\"");
    // ������� ������ FileSystemObject.
    var myFileSysObj = new ActiveXObject("Scripting.FileSystemObject")
    // ������� ������ TextStream.
    var myTextStream = myFileSysObj.OpenTextFile(sFileToParse, 2, true)
    // �������� � ���� ���������
    myTextStream.Write(content);
}

// ����
main();