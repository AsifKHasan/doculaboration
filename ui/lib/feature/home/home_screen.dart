import 'dart:math';

import 'package:flutter/material.dart';
import 'package:gap/gap.dart';
import 'package:ui/core/dependency.dart';
import 'package:ui/feature/home/data/home_api_client.dart';

//Local imports
import 'package:ui/utility/save_file_mobile.dart'
    if (dart.library.html) 'package:ui/utility/save_file_web.dart';

class MyHomePage extends StatefulWidget {
  const MyHomePage({super.key, required this.title});

  final String title;

  @override
  State<MyHomePage> createState() => _MyHomePageState();
}

class _MyHomePageState extends State<MyHomePage> {
  static const _exampleGsheetName =
      "BEC-AI-Chatbot__eoi-submission__book-3__proposed-solution";
  final _gsheetNameController = TextEditingController();
  bool _shouldEnableButton = false;

  bool _isLoading = false;

  @override
  void dispose() {
    _gsheetNameController.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final maxWidth = min(MediaQuery.of(context).size.width, 800).toDouble();
    final themeColor = Theme.of(context).colorScheme.primary;
    return Scaffold(
      appBar: AppBar(
        backgroundColor: Theme.of(context).colorScheme.inversePrimary,
        title: Text(widget.title),
      ),
      body: Align(
        alignment: Alignment.center,
        child: SizedBox(
          width: maxWidth,
          child: Center(
            child: Container(
              padding: const EdgeInsets.all(40),
              decoration: BoxDecoration(
                boxShadow: [
                  BoxShadow(
                    color: themeColor.withOpacity(0.05),
                  ),
                ],
                borderRadius: BorderRadius.circular(10),
              ),
              child: Column(
                mainAxisSize: MainAxisSize.min,
                mainAxisAlignment: MainAxisAlignment.center,
                children: <Widget>[
                  TextFormField(
                    controller: _gsheetNameController,
                    decoration: const InputDecoration(labelText: "Gsheet Name"),
                    onChanged: (value) {
                      _isLoading = false;
                      if (value.trim().isEmpty) {
                        _shouldEnableButton = false;
                      } else {
                        _shouldEnableButton = true;
                      }
                      setState(() {});
                    },
                    onSaved: (newValue) {
                      if (newValue == null) {
                        _shouldEnableButton = false;
                      } else if (newValue.isEmpty) {
                        _shouldEnableButton = false;
                      } else {
                        _shouldEnableButton = true;
                        _downloadAndOpenPDF(newValue.trim());
                      }
                      setState(() {});
                    },
                  ),
                  const Gap(20),
                  _isLoading
                      ? const CircularProgressIndicator()
                      : ElevatedButton(
                          onPressed: !_shouldEnableButton
                              ? null
                              : () => _downloadAndOpenPDF(
                                  _gsheetNameController.text.trim()),
                          child: const Text("Download PDF"),
                        )
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }

  void _downloadAndOpenPDF(String gsheetName) async {
    _isLoading = true;
    setState(() {});
    final homeClient = getIt<HomeApiClient>();
    try {
      final response = await homeClient.downloadGsheetPDF(
        _gsheetNameController.text.trim(),
      );
      await saveAndLaunchFile(
        response,
        "${_gsheetNameController.text}.pdf",
      );
    } catch (error, stackTrace) {
      if (!mounted) return;
      ScaffoldMessenger.of(context).showSnackBar(
        SnackBar(
          content: Column(
            children: [
              const Text("Error"),
              Text(error.toString()),
              const Text("StackTrace"),
              Text(stackTrace.toString()),
            ],
          ),
        ),
      );
    }
    _isLoading = false;
    setState(() {});
  }
}
