import 'dart:async';

import 'package:flutter/material.dart';
import 'package:ui/feature/gseet_input/data/model/job_state_model.dart';
import 'package:ui/feature/gseet_input/presentation/job_state_widget.dart';
import 'package:ui/feature/gseet_input/util/job_state_stage.dart';

class JobListWidget extends StatelessWidget {
  final List<JobModel> jobList;
  final void Function(int) onTap;
  const JobListWidget({
    super.key,
    required this.jobList,
    required this.onTap,
  });

  @override
  Widget build(BuildContext context) {
    return ListView.builder(
      shrinkWrap: true,
      scrollDirection: Axis.horizontal,
      itemBuilder: (context, index) {
        final job = jobList[index];
        return Center(
          child: InkWell(
            child: StatefulBuilder(builder: (
              context,
              setState,
            ) {
              String elapsedTime = job.getElapsedTime();
              if (job.jobStateStage == JobStateStage.started) {
                Timer.periodic(const Duration(seconds: 1), (timer) {
                  setState(() {
                    elapsedTime = job.getElapsedTime();
                  });
                });
              }
              return JobStateWidget(
                stateName: switch (job) {
                  ProcessJobModel() => job.name,
                  DownloadJobModel() =>
                    "${job.name} ${job.downloadFileExtension.name}",
                },
                jobStateStage: job.jobStateStage,
                isFirst: index == 0,
                isLast: index == jobList.length - 1,
                elapsedTime: elapsedTime,
              );
            }),
            onTap: () {
              onTap.call(index);
            },
          ),
        );
      },
      itemCount: jobList.length,
    );
  }
}
