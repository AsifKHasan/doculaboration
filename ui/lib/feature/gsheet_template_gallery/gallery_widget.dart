import 'package:flutter/material.dart';
import 'package:url_launcher/url_launcher_string.dart';

import '../shared/presentation/widget/elevated_card_widget.dart';

class GalleryWidget extends StatelessWidget {
  final List<GalleryItem> items;
  const GalleryWidget({
    super.key,
    required this.items,
  });

  @override
  Widget build(BuildContext context) {
    return GridView.count(
      crossAxisCount: 3,
      mainAxisSpacing: 16,
      crossAxisSpacing: 16,
      childAspectRatio: 16 / 9,
      children: items
          .map(
            (e) => InkWell(
              hoverColor: Theme.of(context).colorScheme.surface,
              child: ElevatedCardWidget(
                child: Center(child: Text(e.title)),
              ),
              onTap: () => launchUrlString(e.url),
            ),
          )
          .toList(),
    );
  }
}

class GalleryItem {
  final String url;
  final String title;

  GalleryItem({
    required this.url,
    required this.title,
  });
}
