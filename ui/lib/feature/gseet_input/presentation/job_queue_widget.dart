import 'package:flutter/material.dart';
import 'package:gap/gap.dart';
import 'package:ui/feature/gseet_input/data/model/job_model.dart';
import 'package:ui/feature/gseet_input/presentation/job_queue_item_widget.dart';

class JobQueueWidget extends StatelessWidget {
  final List<JobQueueItem> jobList;
  final void Function(int, int) onTap;
  const JobQueueWidget({
    super.key,
    required this.jobList,
    required this.onTap,
  });

  @override
  Widget build(BuildContext context) {
    final scrollController = ScrollController();
    return ListView.separated(
      shrinkWrap: true,
      controller: scrollController,
      itemBuilder: (context, index) {
        final jobQueueItem = jobList[index];
        return JobQueueItemWidget(
          name: jobQueueItem.name,
          jobList: jobQueueItem.jobList,
          onTap: (jobIndex) {
            onTap.call(index, jobIndex);
          },
        );
      },
      separatorBuilder: (context, index) => const Gap(8.0),
      itemCount: jobList.length,
    );
  }
}
