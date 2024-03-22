import 'package:flutter/material.dart';
import 'package:ui/feature/faq/domain/presentation/faq_widget.dart';
import 'package:ui/feature/gseet_input/presentation/gsheet_input_widget.dart';
import 'package:ui/feature/screen/landing_screen.dart';

class MyApp extends StatelessWidget {
  const MyApp({super.key});
  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'Doculaboration',
      theme: ThemeData(
        colorScheme: ColorScheme.fromSeed(seedColor: Colors.deepPurple),
        useMaterial3: true,
      ),
      home: LandingScreen(items: [
        DrawerItem(
          title: "Gsheet Input",
          widget: const GsheetNameInputPage(),
          icon: const Icon(Icons.home),
        ),
        DrawerItem(
          title: "FAQ",
          widget: const FAQWidget(),
          icon: const Icon(Icons.question_mark),
        )
      ]),
    );
  }
}
