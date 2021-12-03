
import 'package:flutter/material.dart';
import 'package:flutter/widgets.dart';

class CheckItem extends StatefulWidget {

  String name;

  Function callback;

  CheckItem(this.name, this.callback);

  @override
  State<StatefulWidget> createState() => CheckItemState();

}

class CheckItemState extends State<CheckItem> {

  bool isChecked = false;

  @override
  Widget build(BuildContext context) {
    return CheckboxListTile(
      title: Text(widget.name),
      value: isChecked,
      onChanged: (value){
        setState(() {
          isChecked = value!;
        });
        widget.callback(widget.name, value!);
      },
    );
  }

}