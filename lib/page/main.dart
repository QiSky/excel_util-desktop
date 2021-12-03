import 'dart:io';

import 'package:excel/excel.dart';
import 'package:excel_util/page/check_item.dart';
import 'package:file_picker/file_picker.dart';
import 'package:flutter/cupertino.dart';
import 'package:flutter/material.dart';
import 'package:flutter/widgets.dart';
import 'package:flutter_easyloading/flutter_easyloading.dart';
import 'package:oktoast/oktoast.dart';
import 'package:path_provider/path_provider.dart';

class MainPage extends StatefulWidget {
  @override
  State<StatefulWidget> createState() => MainPageState();
}

class MainPageState extends State<MainPage> {
  String? fileName;

  List<String> selectedTables = [];

  List<String> totalTables = [];

  Excel? _excel;

  String _radioGroupValue = "manual";

  late TextEditingController _startController = TextEditingController();

  late TextEditingController _endController = TextEditingController();

  late TextEditingController _outFileController = TextEditingController();

  late TextEditingController _outTableController = TextEditingController();

  late TextEditingController _outAutoSizeController = TextEditingController();

  late String outPath = '';

  @override
  void initState() {
    EasyLoading.instance
      ..maskType = EasyLoadingMaskType.clear
      ..indicatorType = EasyLoadingIndicatorType.pouringHourGlass
      ..contentPadding = const EdgeInsets.symmetric(vertical: 15.0, horizontal: 15.0)
      ..userInteractions = false;
    Future.delayed(Duration.zero,() {
      getApplicationDocumentsDirectory().then((value) {
        setState(() {
          outPath = value.absolute.path;
        });
      });
    });
    super.initState();
  }
  @override
  Widget build(BuildContext context) {
    return Scaffold(
      appBar: AppBar(
        title: Text('excel转换工具'),
        actions: [
          IconButton(
              onPressed: () async{
                EasyLoading.show(status: "执行中。。。");
                if(selectedTables.isEmpty) {
                  showToast('请至少选中一个表格');
                  EasyLoading.dismiss();
                  return;
                }
                if(_radioGroupValue == "manual") {
                  if (_startController.text.isEmpty) {
                    showToast('开始行不能为空');
                    EasyLoading.dismiss();
                    return;
                  }
                  if (_endController.text.isEmpty) {
                    showToast('结束行不能为空');
                    EasyLoading.dismiss();
                    return;
                  }
                  int? startNumber = int.tryParse(_startController.text);
                  if(startNumber == null || startNumber <= 0) {
                    showToast('开始行输入不合法');
                    EasyLoading.dismiss();
                    return;
                  }
                  int? endNumber = int.tryParse(_endController.text);
                  if(endNumber == null || endNumber <= 0) {
                    showToast('结束行输入不合法');
                    EasyLoading.dismiss();
                    return;
                  }
                } else {
                  if(_outAutoSizeController.text.isEmpty) {
                    showToast('输出行数量不能为空');
                    EasyLoading.dismiss();
                    return;
                  }
                  int? number = int.tryParse(_outAutoSizeController.text);
                  if(number == null || number <= 0) {
                    showToast('结束行输入不合法');
                    EasyLoading.dismiss();
                    return;
                  }
                }
                if(_outTableController.text.isEmpty) {
                  showToast('输出表名不能为空');
                  EasyLoading.dismiss();
                  return;
                }
                if(_outFileController.text.isEmpty) {
                  showToast('输出文件名不能为空');
                  EasyLoading.dismiss();
                  return;
                }
                if(_radioGroupValue == "manual") {
                  var excel = Excel.createExcel();
                  excel.rename("Sheet1", _outTableController.text);
                  totalTables.forEach((element) {
                    // List<List<dynamic>?>? data = _excel?.tables[element]?.selectRangeValues(
                    //     CellIndex.indexByColumnRow(rowIndex: int.tryParse(_startController.text), columnIndex: 1),
                    //   end: CellIndex.indexByColumnRow(rowIndex: int.tryParse(_endController.text), columnIndex: _excel?.tables[element]?.maxCols)
                    // );
                    int? cols = _excel?.tables[element]?.maxCols;
                    List<List<dynamic>?>? data = _excel?.tables[element]?.selectRangeValues(
                        CellIndex.indexByColumnRow(rowIndex: int.tryParse(_startController.text)!-1, columnIndex: 0),
                        end: CellIndex.indexByColumnRow(rowIndex: int.tryParse(_endController.text)!-1, columnIndex: cols!-1)
                    );
                    if(data?.isNotEmpty == true) {
                      data?.forEach((element1) {
                        excel.appendRow(element, element1!);
                      });
                    }
                  });
                  var fileBytes = excel.save(fileName: _outFileController.text+".xlsx");
                  File("$outPath/${_outFileController.text+".xlsx"}")
                  ..createSync(recursive: true)
                  ..writeAsBytesSync(fileBytes!);
                  showToast("任务完成");
                } else {
                  var ii = 0;
                  totalTables.forEach((element) {
                    int? maxRows = _excel?.tables[element]?.maxRows;
                    int? cols = _excel?.tables[element]?.maxCols;
                    if(int.parse(_outAutoSizeController.text) > maxRows!) {
                      var excel = Excel.createExcel();
                      excel.rename("Sheet1", _outTableController.text);
                      List<List<dynamic>?>? data = _excel?.tables[element]?.selectRangeValues(
                          CellIndex.indexByColumnRow(rowIndex: 0, columnIndex: 0),
                          end: CellIndex.indexByColumnRow(rowIndex: maxRows-1, columnIndex: cols!-1)
                      );
                      // for(var i = 0;i<maxRows;i++) {
                      //   excel.appendRow(_outTableController.text, _excel!.tables[element]!.row(i));
                      // }
                      data?.forEach((element1) {
                        excel.appendRow(element, element1!);
                      });
                      var fileBytes = excel.save(fileName: _outFileController.text+"_${ii++}.xlsx");
                      File("$outPath/${_outFileController.text+"_${ii}.xlsx"}")
                        ..createSync(recursive: true)
                        ..writeAsBytesSync(fileBytes!);
                      showToast("任务完成");
                    } else {
                      int limit = maxRows%int.parse(_outAutoSizeController.text) != 0?
                      maxRows~/int.parse(_outAutoSizeController.text)+1:maxRows~/int.parse(_outAutoSizeController.text);
                      for(var i = 0;i< limit;i++) {
                        var excel = Excel.createExcel();
                        excel.rename("Sheet1", _outTableController.text);
                        for(var j = i*int.parse(_outAutoSizeController.text);j< ((i+1)*int.parse(_outAutoSizeController.text) > maxRows ? maxRows:(i+1)*int.parse(_outAutoSizeController.text)); j++) {
                          List<List<dynamic>?>? data = _excel?.tables[element]?.selectRangeValues(
                              CellIndex.indexByColumnRow(rowIndex: j, columnIndex: 0),
                              end: CellIndex.indexByColumnRow(rowIndex: j, columnIndex: cols!-1)
                          );
                          data?.forEach((element1) {
                            excel.appendRow(element, element1!);
                          });
                          // excel.appendRow(_outTableController.text, _excel!.tables[element]!.row(j));
                        }
                        var fileBytes = excel.save(fileName: _outFileController.text+"_${ii++}.xlsx");
                        File("$outPath/${_outFileController.text+"_${ii}.xlsx"}")
                          ..createSync(recursive: true)
                          ..writeAsBytesSync(fileBytes!);
                      }
                      showToast("任务完成");
                    }
                  });
                  showToast("任务完成");
                }
                EasyLoading.dismiss();
              },
              icon: Icon(Icons.event))
        ],
      ),
      body: Container(
        height: MediaQuery.of(context).size.height,
        padding: EdgeInsets.only(left: 10, right: 10),
        child: SingleChildScrollView(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              SizedBox(
                height: 10,
              ),
              Row(
                children: [
                  Text('选择文件名：${fileName ?? ''}'),
                  Spacer(),
                  RaisedButton(
                    child: Text("选择文件"),
                    onPressed: () async {
                      FilePickerResult? result =
                          await FilePicker.platform.pickFiles(
                        type: FileType.custom,
                        allowedExtensions: ['xls', 'xlsx'],
                      );
                      if (result != null) {
                        var file = File(result.files.single.path!);
                        var bytes = file.readAsBytesSync();
                        _excel = Excel.decodeBytes(bytes);
                        for (var table in _excel!.tables.keys) {
                          totalTables.add(table);
                        }
                        setState(() {
                          fileName = result.files.single.name;
                        });
                        _outFileController.text =
                            (fileName?.split(".").first ?? 'data') + "_output";
                        _outTableController.text =
                            totalTables.isEmpty ? 'Sheet1' : totalTables[0];
                      }
                    },
                  )
                ],
              ),
              SizedBox(
                height: 10,
              ),
              Text('请选择表名：'),
              SizedBox(
                height: 10,
              ),
              ListView.builder(
                  shrinkWrap: true,
                  itemCount: totalTables.length,
                  itemBuilder: (context, index) {
                    return CheckItem(
                        totalTables[index] +
                            "  总行数：${_excel![totalTables[index]].maxRows}",
                        (String name, bool value) {
                      if (value) {
                        if (!selectedTables.contains(name.split("  ").first)) {
                          selectedTables.add(name.split("  ").first);
                        }
                      } else {
                        if (selectedTables.contains(name.split("  ").first)) {
                          selectedTables.removeWhere(
                              (element) => element == name.split("  ").first);
                        }
                      }
                    });
                  }),
              Text('输出选项：'),
              SizedBox(
                height: 10,
              ),
              Row(
                children: [
                  Expanded(
                    child: RadioListTile(
                      title: Text('手动模式'),
                      value: 'manual',
                      groupValue: _radioGroupValue,
                      onChanged: (value) {
                        setState(() {
                          _radioGroupValue = value.toString();
                        });
                      },
                    ),
                  ),
                  Expanded(
                      child: RadioListTile(
                    title: Text('自动模式'),
                    value: 'auto',
                    groupValue: _radioGroupValue,
                    onChanged: (value) {
                      setState(() {
                        _radioGroupValue = value.toString();
                      });
                    },
                  )),
                  Spacer()
                ],
              ),
              SizedBox(
                height: 10,
              ),
              _radioGroupValue == "manual" ?Column(
                children: [
                  Row(
                    children: [
                      Expanded(
                          child: Row(
                            children: [
                              Text('开始行（大于0且不能为小数)：'),
                              Container(
                                  width: 300,
                                  child: CupertinoTextField(
                                    controller: _startController,
                                    clearButtonMode: OverlayVisibilityMode.editing,
                                    keyboardType: TextInputType.number,
                                  )),
                              Spacer(),
                            ],
                          ))
                    ],
                  ),
                  SizedBox(
                    height: 10,
                  ),
                  Row(
                    children: [
                      Expanded(
                          child: Row(
                            children: [
                              Text('结束行（大于0且不能为小数)：'),
                              Container(
                                  width: 300,
                                  child: CupertinoTextField(
                                    controller: _endController,
                                    clearButtonMode: OverlayVisibilityMode.editing,
                                    keyboardType: TextInputType.number,
                                  )),
                              Spacer(),
                            ],
                          ))
                    ],
                  ),
                ],
              ) : Column(
                children: [
                  Row(
                    children: [
                      Expanded(
                          child: Row(
                            children: [
                              Text('输出行数量（从1开始累加）：'),
                              Container(
                                  width: 300,
                                  child: CupertinoTextField(
                                    controller: _outAutoSizeController,
                                    clearButtonMode: OverlayVisibilityMode.editing,
                                    keyboardType: TextInputType.number,
                                  )),
                              Spacer(),
                            ],
                          ))
                    ],
                  ),
                ],
              ),
              SizedBox(
                height: 10,
              ),
              Row(
                children: [
                  Expanded(
                      child: Row(
                    children: [
                      Text('输出文件名.xlsx(自动模式下输出文件自动增加后缀名_index）：'),
                      Container(
                          width: 200,
                          child: CupertinoTextField(
                            controller: _outFileController,
                            clearButtonMode: OverlayVisibilityMode.editing,
                          )),
                      Spacer(),
                    ],
                  ))
                ],
              ),
              SizedBox(
                height: 10,
              ),
              Row(
                children: [
                  Expanded(
                      child: Row(
                    children: [
                      Text('输出的表名（不建议写中文）：'),
                      Container(
                          width: 300,
                          child: CupertinoTextField(
                            controller: _outTableController,
                            clearButtonMode: OverlayVisibilityMode.editing,
                          )),
                      Spacer(),
                    ],
                  ))
                ],
              ),
              SizedBox(
                height: 10,
              ),
              Row(
                children: [
                  Expanded(
                      child: Row(
                        children: [
                          Text('输出的路径：$outPath'),
                        ],
                      ))
                ],
              ),
            ],
          ),
        ),
      ),
    );
  }
}
