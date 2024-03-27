import 'package:flutter/material.dart';
import 'package:gap/gap.dart';
import 'package:ui/feature/gseet_input/presentation/download_option_selector_widget.dart';
import 'package:ui/feature/gseet_input/util/file_extension.dart';
import 'package:ui/feature/shared/presentation/widget/elevated_card_widget.dart';

class GsheetNameInputPage extends StatefulWidget {
  final Function(String, List<FileExtension>) onSubmitGsheetName;
  const GsheetNameInputPage({
    super.key,
    required this.onSubmitGsheetName,
  });

  @override
  State<GsheetNameInputPage> createState() => _GsheetNameInputPageState();
}

class _GsheetNameInputPageState extends State<GsheetNameInputPage> {
  static const _exampleGsheetName =
      "BEC-AI-Chatbot__eoi-submission__book-3__proposed-solution";
  final _gsheetNameController = TextEditingController();
  bool _shouldEnableButton = false;

  final downloadOptions = [
    FileExtension.json,
    FileExtension.odt,
    FileExtension.pdf,
  ];
  List<int> _userSelection = <int>[];

  @override
  void dispose() {
    _gsheetNameController.dispose();
    super.dispose();
  }

  @override
  void initState() {
    super.initState();

    /// default selection
    _userSelection = [
      downloadOptions.indexWhere(
        (element) => element == FileExtension.pdf,
      )
    ];
    setState(() {});
  }

  @override
  Widget build(BuildContext context) {
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
            Row(
              mainAxisAlignment: MainAxisAlignment.end,
              crossAxisAlignment: CrossAxisAlignment.center,
              children: [
                Expanded(
                  child: SizedBox(
                    height: 40,
                    width: 300,
                    child: DownloadOptionSelectorWidget(
                      downloadOptions: downloadOptions,
                      onChangedSelection: (intList) {
                        setState(() {
                          _userSelection = intList;
                        });
                      },
                      initialSelection: _userSelection,
                    ),
                  ),
                ),
                ElevatedButton(
                  onPressed: !_shouldEnableButton
                      ? null
                      : () => _downloadAndOpenPDF(
                          _gsheetNameController.text.trim()),
                  child: const Text("Add To Queue"),
                )
              ],
            )
          ],
        ),
      ),
    );
  }

  void _downloadAndOpenPDF(String gsheetName) async {
    widget.onSubmitGsheetName(
      gsheetName,
      _userSelection.map((e) => downloadOptions[e]).toList(),
    );
  }
}
