import 'package:flutter/material.dart';
import 'package:gap/gap.dart';

class LogWidget extends StatelessWidget {
  final String log;
  final bool started;
  const LogWidget({
    super.key,
    required this.log,
    required this.started,
  });

  @override
  Widget build(BuildContext context) {
    final scrollController = ScrollController();
    WidgetsBinding.instance.addPostFrameCallback((duration) {
      scrollController.animateTo(
        scrollController.position.maxScrollExtent,
        duration: const Duration(milliseconds: 500),
        curve: Curves.ease,
      );
    });
    return SingleChildScrollView(
      controller: scrollController,
      child: Align(
        alignment: Alignment.topLeft,
        child: Column(
          children: [
            Text(
              log,
            ),
            !started
                ? const SizedBox.shrink()
                : const SizedBox(
                    width: 20,
                    child: LinearProgressIndicator(),
                  ),
            const Gap(8.0),
          ],
        ),
      ),
    );
  }
}
