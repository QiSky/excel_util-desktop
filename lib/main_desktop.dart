import 'package:flutter/foundation.dart';

import 'main.dart' as original_main;

// This file is the default main entry-point for go-flutter application.
void main() {
  debugDefaultTargetPlatformOverride = TargetPlatform.windows;
  original_main.main();
}
