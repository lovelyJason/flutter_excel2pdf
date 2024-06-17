import 'dart:io';
// import 'package:archive/archive.dart'
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:excel/excel.dart' as excelLib;
import 'package:path/path.dart' as path;
import 'package:docx_template/docx_template.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/services.dart';
import 'package:desktop_window/desktop_window.dart' as window_size;

// import 'package:flutter_dropzone/flutter_dropzone.dart';
// import 'package:flutter_dropzone_platform_interface/flutter_dropzone_platform_interface.dart';

void main() async {
  WidgetsFlutterBinding.ensureInitialized();
  if (Platform.isWindows || Platform.isLinux || Platform.isMacOS) {
    await window_size.DesktopWindow.setWindowSize(Size(400, 380));
    await window_size.DesktopWindow.setMinWindowSize(Size(400, 380));
    await window_size.DesktopWindow.setMaxWindowSize(Size(4000, 380));
  }
  runApp(MyApp());
}

class MyApp extends StatelessWidget {
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      home: HomePage(),
    );
  }
}

class HomePage extends StatefulWidget {
  @override
  _HomePageState createState() => _HomePageState();
}

class _HomePageState extends State<HomePage> {
  String fileContent = 'Excel Content: None';
  late Directory outputDir;
  // late DropzoneViewController controller; // 添加这个变量

  @override
  void initState() {
    super.initState();
    _initOutputDir();
  }

  Future testGenerateDocxFromTemplate() async {
    try {
      // 从模板文件中创建 DocxTemplate 实例
      final docx = await DocxTemplate.fromBytes(
          await File('assets/template.docx').readAsBytes());

      Content content = Content();
      // 生成 Word 文档，根据传入的 Content 对象进行替换占位符
      content
        ..add(TextContent('docname', 'Jason Huang'))
        ..add(TextContent("passport", "Passport NE0323 4456673"))
        ..add(TextContent('{{G-header}}', 'replacement1'))
        ..add(TextContent('{H-header}', 'replacement2'))
        ..add(TextContent('header', 'replacement2'));

      final docGenerated = await docx.generate(content);
      print('docGenerated的类型是：${docGenerated.runtimeType}');

      // 获取存储目录
      final directory = await getApplicationDocumentsDirectory();
      // 生成路径
      final outputFile =
          File('${directory.path}/generated_docx_with_replaced_content.docx');

      if (docGenerated != null) {
        // 写入文件
        print('生成成功: ${directory.path}');
        await outputFile.writeAsBytes(docGenerated);
      }
    } catch (e, stackTrace) {
      print('Exception: $e');
      print('StackTrace: $stackTrace'); // 调用方法获取堆栈跟踪
    }
  }

  Future<Directory> getDesktopPath() async {
    Directory desktopDir;

    if (Platform.isMacOS) {
      desktopDir = Directory('/Users/${Platform.environment['USER']}/Desktop');
    } else if (Platform.isWindows) {
      String userProfile =
          Platform.environment['USERPROFILE'] ?? ''; // 使用??提供默认空字符串
      desktopDir = Directory(userProfile + '\\Desktop');
    } else {
      throw UnsupportedError('Unsupported platform for getting desktop path.');
    }

    // Optionally, you can check if the directory exists before returning it
    // though the createSync in _initOutputDir will handle non-existent dirs
    // if (await desktopDir.exists()) {
    //   return desktopDir;
    // } else {
    //   throw StateError('Desktop directory not found.');
    // }

    return desktopDir;
  }

  Future<void> _initOutputDir() async {
    // outputDir = await getApplicationDocumentsDirectory();
    // Try to get the desktop path for macOS and Windows
    if (Platform.isMacOS || Platform.isWindows) {
      // Custom logic since `path_provider` doesn't directly support desktop paths
      outputDir = await getDesktopPath();
    } else {
      // For other platforms like mobile, use application documents directory
      outputDir = await getApplicationDocumentsDirectory();
    }

    final ayeshaDir = Directory(path.join(outputDir.path, 'ayesha'));
    if (!ayeshaDir.existsSync()) {
      ayeshaDir.createSync(recursive: true);
    }
    outputDir = ayeshaDir;
    print('outputDir: $outputDir');
  }

  Future<List<List<dynamic>>> readExcel(String filePath) async {
    var fileBytes = File(filePath).readAsBytesSync();
    var excel = excelLib.Excel.decodeBytes(fileBytes);

    List<List<dynamic>> rows = [];
    for (var table in excel.tables.keys) {
      var sheet = excel.tables[table];
      if (sheet != null) {
        for (var row in sheet.rows) {
          rows.add(row);
        }
      }
    }

    return rows;
  }

  Future<void> generateWordDocuments(List<List<dynamic>> excelData,
      List<int> templateBytes, String prefix) async {
    try {
      // 获取表头行数据，，从索引6，第7列，G开始获取这一行中所有后续列的值。这意味着headers变量将包含第7列及其之后所有列的表头。
      var headers = excelData[0].sublist(6);
      // excelData.sublist(1)：从Excel数据的第2行（索引为1）开始，获取所有后续行的数据。这表示去掉第一行表头，保留实际数据行。
      var rows = excelData.sublist(1).map((row) => row.sublist(6)).toList();

      List<double> sums = List<double>.filled(headers.length, 0);
      List<int> counts = List<int>.filled(headers.length, 0);
      List<List<String>> nonNumericContents =
          List.generate(headers.length, (_) => []);

      // 遍历每一行的数据
      for (var data in rows) {
        data.asMap().forEach((index, cell) {
          var cellValue = cell?.value;

          // 检查 cellValue 是否是数字
          if (cellValue is num && cellValue != 1 && cellValue != 2) {
            sums[index] += cellValue.toDouble();
            counts[index]++;
          } else {
            nonNumericContents[index].add(cellValue.toString());
          }
        });
      }

      // 计算平均数
      List<dynamic> averagesOrContents = sums.asMap().entries.map((entry) {
        int index = entry.key;
        double sum = entry.value;
        int count = counts[index];
        return count > 0
            ? double.parse((sum / count).toStringAsFixed(1))
            : nonNumericContents[index].join('\n');
      }).toList();
      // .cast<double>();

      var templateData = <String, dynamic>{};

      headers.asMap().forEach((index, header) {
        // 一个cell对象，包含了单元格的内容以及其他属性，如样式、对齐方式等。以下是对问题的详细解答：
        // print(header.value.toString());
        // print(data[index].toString());
        var columnLetter = String.fromCharCode(71 + index);
        templateData['${columnLetter}-header'] = header?.value?.toString();
        templateData['${columnLetter}-content'] =
            averagesOrContents[index].toString();
      });
      var docx =
          await DocxTemplate.fromBytes(Uint8List.fromList(templateBytes));
      Content c = Content();

      // templateData是一个{G-header:表头值，G-content:内容值, H-header:表头值，H-content:内容值}的map
      templateData.forEach((key, value) {
        c.add(TextContent(key, value.toString()));
      });

      //   // 生成 Word 文档，根据传入的 Content 对象进行替换占位符
      var generatedDocBuffer = await docx.generate(c);
      if (generatedDocBuffer != null) {
        var modifiableBuffer =
            List<int>.from(generatedDocBuffer); // Create a modifiable list
        var outputFilePath =
            path.join(outputDir.path, 'result_${prefix}_平均数.docx');
        File(outputFilePath).writeAsBytesSync(modifiableBuffer);
        print('Word 文件已成功生成: $outputFilePath');
      } else {
        print('生成 Word 文件时发生错误');
      }

      // for (int i = 0; i < rows.length; i++) {
      //   var data = rows[i];
      // }
    } catch (e) {
      throw Exception('发生错误: $e');
    }
  }

  void showMessageBox(BuildContext context, String message) {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        content: Text(message),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(context),
            child: Text('确认'),
          ),
        ],
      ),
    );
  }

  void pickFile() async {
    // testGenerateDocxFromTemplate();
    // return;
    var result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
        allowMultiple: true
    );

    if (result != null) {
      // var filePath = result.files.single.path;
      try {
        var templateContent = await rootBundle.load('assets/template.docx');
        // buffer.asUint8List()不可修改。您可以通过将其转换为普通列表来修改它
        var templateBytes = templateContent.buffer.asUint8List().toList();
        for (var file in result.files) {
          var filePath = file.path;
          if (filePath != null) {
            String fileName = path.basenameWithoutExtension(filePath);
            print('filename: $fileName');
            var excelData = await readExcel(filePath);
            await generateWordDocuments(excelData, templateBytes, fileName);
          }
        }
        var outputFiles =
            outputDir.listSync().map((file) => file.path).join('\n');
        setState(() {
          fileContent = outputFiles;
        });

        showMessageBox(context, '老板，已为您处理完成\n已经放到了桌面下的ayesha目录');
      } catch (e, stackTrace) {
        print('发生错误: $e');
        showMessageBox(context, '发生错误: $e, $stackTrace');
      }
    }
  }

  void showUsageDialog() {
    showDialog(
      context: context,
      builder: (context) => AlertDialog(
        title: Text('使用说明', style: TextStyle(fontSize: 16)),
        content: Text(
          '1. 点击"选择Excel"按钮选择一个Excel文件,目前只能单个文件。\n'
          '2. 系统会自动读取Excel文件并生成对应的Word文档，理论上任意的excel都可以处理为word,自定义模板格式即可。\n'
          '3. 生成的文档会保存在桌面的"ayesha"目录中。\n'
          '4. 如果遇到任何问题，请联系开发者。',
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.pop(context),
            child: Text('确认'),
          ),
        ],
      ),
    );
  }

  Future<void> downloadTemplate() async {
    try {
      final templateContent = await rootBundle.load('assets/template.docx');
      final templateBytes = templateContent.buffer.asUint8List();
      final desktopDir = await getDesktopPath();
      final templateFile = File(path.join(desktopDir.path, 'template.docx'));
      await templateFile.writeAsBytes(templateBytes);
      showMessageBox(context, '模板文件已成功下载到桌面');
    } catch (e) {
      showMessageBox(context, '下载模板文件时发生错误: $e');
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text(
          'Excel2PDF2.0 by jasonhuang QQ315945659',
          style: TextStyle(fontSize: 12),
        ),
        actions: [
          IconButton(
            icon: Icon(Icons.info_outline),
            onPressed: showUsageDialog,
          ),
        ],
      ),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            ElevatedButton(
              onPressed: pickFile,
              child: Text('选择Excel(暂不支持拖拽excel)'),
            ),
            SizedBox(height: 20),
            Expanded(
                child: SingleChildScrollView(
              child: Text(
                fileContent,
                style: TextStyle(fontSize: 16),
              ),
            ))
          ],
        ),
      ),
      floatingActionButton: Container(
        margin: EdgeInsets.only(bottom: 30.0), // 调整距离底部的距离
        child: FloatingActionButton.extended(
          onPressed: downloadTemplate,
          tooltip: '下载模板',
          label: Text('下载模板'),
          icon: Icon(Icons.download),
        ),
      ),
      floatingActionButtonLocation: FloatingActionButtonLocation.centerDocked,
    );
  }
}
