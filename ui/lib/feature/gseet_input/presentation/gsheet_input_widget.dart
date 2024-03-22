import 'package:flutter/material.dart';
import 'package:gap/gap.dart';
import 'package:ui/core/dependency.dart';

//Local imports
import 'package:ui/utility/save_file_mobile.dart'
    if (dart.library.html) 'package:ui/utility/save_file_web.dart';

import '../../screen/landing_screen.dart';
import '../data/home_api_client.dart';

class GsheetNameInputPage extends StatefulWidget {
  const GsheetNameInputPage({
    super.key,
  });

  @override
  State<GsheetNameInputPage> createState() => _GsheetNameInputPageState();
}

class _GsheetNameInputPageState extends State<GsheetNameInputPage> {
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
    final themeColor = Theme.of(context).colorScheme.primary;
    return ElevatedCardWidget(
      child: Padding(
        padding: const EdgeInsets.symmetric(
          vertical: 32.0,
          horizontal: 16.0,
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
