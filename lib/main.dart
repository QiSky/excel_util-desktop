import 'dart:async';

import 'package:excel_util/page/main.dart';
import 'package:flutter/foundation.dart';
import 'package:flutter/material.dart';
import 'package:flutter_easyloading/flutter_easyloading.dart';
import 'package:oktoast/oktoast.dart';

void main() {
  runZoned(() {
    runApp(MyApp());
  });
}

class MyApp extends StatelessWidget {
  // This widget is the root of your application.
  @override
  Widget build(BuildContext context) {
    return OKToast(child: MaterialApp(
      debugShowCheckedModeBanner: false,
      title: 'excel工具',
      theme: ThemeData(
        primarySwatch: Colors.blue,
      ),
      home: MainPage(),
      builder: EasyLoading.init(),
    ));
  }
}
