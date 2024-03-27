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
    return ListView.separated(
      shrinkWrap: true,
      scrollDirection: Axis.horizontal,
      itemBuilder: (context, index) {
        final job = jobList[index];
        return Center(
          child: InkWell(
            child: JobStateWidget(
              stateName: switch (job) {
                ProcessJobModel() => job.name,
                DownloadJobModel() =>
                  "${job.name} ${job.downloadFileExtension.name}",
              },
              jobStateStage: job.jobStateStage,
            ),
            onTap: () {
              onTap.call(index);
            },
            onHover: (value) {
              if (value) {
                onTap.call(index);
              }
            },
          ),
        );
      },
      separatorBuilder: (context, index) {
        final job = jobList[index];
        return Center(
          child: Container(
            color: job.jobStateStage == JobStateStage.success
                ? Colors.green
                : Colors.grey,
            height: 2,
            width: 20,
          ),
        );
      },
      itemCount: jobList.length,
    );
  }
}
