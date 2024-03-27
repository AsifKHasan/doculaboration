import 'package:flutter/material.dart';
import 'package:gap/gap.dart';
import 'package:ui/feature/gseet_input/util/file_extension.dart';

class DownloadOptionSelectorWidget extends StatelessWidget {
  final List<FileExtension> downloadOptions;
  final List<int> initialSelection;
  final void Function(List<int>) onChangedSelection;

  const DownloadOptionSelectorWidget({
    super.key,
    required this.downloadOptions,
    required this.onChangedSelection,
    required this.initialSelection,
  });

  @override
  Widget build(BuildContext context) {
    final list = initialSelection;
    return ListView.separated(
      scrollDirection: Axis.horizontal,
      itemBuilder: (context, index) {
        final job = downloadOptions[index];
        return Row(
          children: [
            InputChip(
              label: Text(job.name),
              selected: initialSelection.contains(index),
              onSelected: (isSelected) {
                if (isSelected) {
                  list.add(index);
                } else {
                  list.remove(index);
                }
                onChangedSelection(list);
              },
            ),
          ],
        );
      },
      separatorBuilder: (context, index) => const Gap(8.0),
      itemCount: downloadOptions.length,
    );
  }
}
