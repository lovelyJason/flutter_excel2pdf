import 'dart:io';
import 'package:flutter/material.dart';
import 'package:path_provider/path_provider.dart';
import 'package:excel/excel.dart';
import 'package:path/path.dart' as path;
import 'package:docx_template/docx_template.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/services.dart';

void main() {
  WidgetsFlutterBinding.ensureInitialized();
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

  @override
  void initState() {
    super.initState();
    _initOutputDir();
  }

  Future<void> _initOutputDir() async {
    outputDir = await getApplicationDocumentsDirectory();
    final ayeshaDir = Directory(path.join(outputDir.path, 'ayesha'));
    if (!ayeshaDir.existsSync()) {
      ayeshaDir.createSync(recursive: true);
    }
    outputDir = ayeshaDir;
  }

  Future<List<List<dynamic>>> readExcel(String filePath) async {
    var fileBytes = File(filePath).readAsBytesSync();
    var excel = Excel.decodeBytes(fileBytes);

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

  Future<void> generateWordDocuments(
      List<List<dynamic>> excelData, List<int> templateBytes) async {
    try {
      var headers = excelData[0].sublist(7);
      var rows = excelData.sublist(1).map((row) => row.sublist(7)).toList();

      for (int i = 0; i < rows.length; i++) {
        var data = rows[i];
        var templateData = <String, dynamic>{};

        headers.asMap().forEach((index, header) {
          var columnLetter = String.fromCharCode(71 + index);
          templateData['${columnLetter}-header'] = header;
          templateData['${columnLetter}-content'] = data[index];
        });

        var docx =
            await DocxTemplate.fromBytes(Uint8List.fromList(templateBytes));
        Content c = Content();

        templateData.forEach((key, value) {
          c.add(TextContent(key, value.toString()));
        });

        var generatedDocBuffer = await docx.generate(c);
        if (generatedDocBuffer != null) {
          var modifiableBuffer =
              List<int>.from(generatedDocBuffer); // Create a modifiable list
          var outputFilePath =
              path.join(outputDir.path, 'result_${i + 1}.docx');
          File(outputFilePath).writeAsBytesSync(modifiableBuffer);
          print('Word 文件已成功生成: $outputFilePath');
        } else {
          print('生成 Word 文件时发生错误');
        }
      }
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
    var result = await FilePicker.platform.pickFiles(
      type: FileType.custom,
      allowedExtensions: ['xlsx'],
    );

    if (result != null) {
      var filePath = result.files.single.path;
      if (filePath != null) {
        try {
          var excelData = await readExcel(filePath);
          var templateContent = await rootBundle.load('assets/template.docx');
          // buffer.asUint8List()不可修改。您可以通过将其转换为普通列表来修改它
          var templateBytes = templateContent.buffer.asUint8List().toList();
          await generateWordDocuments(excelData, templateBytes);
          var outputFiles =
              outputDir.listSync().map((file) => file.path).join('\n');
          setState(() {
            fileContent = outputFiles;
          });

          showMessageBox(context, '老板，已为您处理完成\n已经放到了文档目录的ayesha目录');
        } catch (e) {
          print('发生错误: $e');
          showMessageBox(context, '发生错误: $e');
        }
      }
    }
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('Excel2PDF'),
      ),
      body: Center(
        child: Column(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            ElevatedButton(
              onPressed: pickFile,
              child: Text('选择Excel(支持拖拽excel进来)'),
            ),
            SizedBox(height: 20),
            SingleChildScrollView(
              child: Text(
                fileContent,
                style: TextStyle(fontSize: 16),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
