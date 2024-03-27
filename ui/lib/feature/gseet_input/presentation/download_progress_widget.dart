import 'package:flutter/material.dart';

class DownloadProgressWidget extends StatelessWidget {
  final double progressPercentage;
  final double speed;
  final String speedUnit;

  const DownloadProgressWidget({
    super.key,
    required this.progressPercentage,
    required this.speed,
    required this.speedUnit,
  });

  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        const Text("Download"),
        Row(
          children: [
            LinearProgressIndicator(
              value: progressPercentage / 100,
            ),
            Text("$speed $speedUnit")
          ],
        )
      ],
    );
  }
}
