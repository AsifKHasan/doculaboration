import 'package:flutter/material.dart';
import 'package:flutter_markdown/flutter_markdown.dart';
import 'package:ui/feature/screen/landing_screen.dart';

import '../../../../gen/assets.gen.dart';

class FAQWidget extends StatefulWidget {
  const FAQWidget({super.key});

  @override
  State<FAQWidget> createState() => _FAQWidgetState();
}

class _FAQWidgetState extends State<FAQWidget> {
  String markdownString = "";
  @override
  Widget build(BuildContext context) {
    return _shouldShowMarkdownWidget()
        ? SelectionArea(
            child: ElevatedCardWidget(
              child: Markdown(
                data: markdownString,
              ),
            ),
          )
        : const CircularProgressIndicator();
  }

  bool _shouldShowMarkdownWidget() {
    return markdownString.isNotEmpty;
  }

  @override
  void initState() {
    super.initState();

    DefaultAssetBundle.of(context).loadString(Assets.faq).then(
      (value) {
        markdownString = value;
        setState(() {});
      },
    );
  }
}
