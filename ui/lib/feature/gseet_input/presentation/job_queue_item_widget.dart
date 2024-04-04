import 'package:flutter/material.dart';
import 'package:ui/feature/gseet_input/presentation/job_list_widget.dart';

import '../../shared/presentation/widget/elevated_card_widget.dart';
import '../data/model/job_state_model.dart';

class JobQueueItemWidget extends StatelessWidget {
  final String name;
  final List<JobModel> jobList;
  final void Function(int) onTap;

  const JobQueueItemWidget({
    super.key,
    required this.name,
    required this.onTap,
    required this.jobList,
  });

  @override
  Widget build(BuildContext context) {
    return ElevatedCardWidget(
      child: Padding(
        padding: const EdgeInsets.all(8.0),
        child: Row(
          crossAxisAlignment: CrossAxisAlignment.center,
          mainAxisAlignment: MainAxisAlignment.spaceBetween,
          children: [
            SizedBox(
              width: 250,
              child: Text(
                name,
              ),
            ),
            Expanded(
              child: Container(
                alignment: Alignment.centerRight,
                height: 80,
                child: JobListWidget(
                  jobList: jobList,
                  onTap: onTap,
                ),
              ),
            ),
          ],
        ),
      ),
    );
  }
}
