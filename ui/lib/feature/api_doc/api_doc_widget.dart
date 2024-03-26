import 'package:flutter/material.dart';
import 'package:flutter_webview_plugin/flutter_webview_plugin.dart';

class ApiDocWidget extends StatelessWidget {
  final String url;

  const ApiDocWidget({
    super.key,
    required this.url,
  });

  @override
  Widget build(BuildContext context) {
    return WebviewScaffold(
      url: url, // Specify the URL you want to load
      withZoom: true, // Enable zooming
      withLocalStorage: true, // Enable local storage
      hidden: true, // Hide the webview initially until it's fully loaded
      initialChild: const Center(
        child: CircularProgressIndicator(), // Show a loading indicator
      ),
    );
  }
}
