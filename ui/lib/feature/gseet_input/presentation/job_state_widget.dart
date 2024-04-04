import 'package:flutter/material.dart';
import 'package:ui/feature/gseet_input/util/job_state_stage.dart';
import 'package:ui/feature/shared/presentation/widget/box_container.dart';

class JobStateWidget extends StatelessWidget {
  final JobStateStage jobStateStage;
  final String stateName;
  final bool isFirst;
  final bool isLast;
  final String elapsedTime;

  const JobStateWidget({
    super.key,
    required this.stateName,
    required this.jobStateStage,
    required this.isFirst,
    required this.isLast,
    required this.elapsedTime,
  });

  @override
  Widget build(BuildContext context) {
    return Tooltip(
      message: "Job($stateName): ${jobStateStage.name}",
      child: Row(
        crossAxisAlignment: CrossAxisAlignment.center,
        children: [
          isFirst
              ? const SizedBox.shrink()
              : Container(
                  width: 20,
                  height: 2,
                  color: jobStateStage == JobStateStage.success
                      ? Colors.green
                      : Colors.grey,
                ),
          BoxContainer(
            child: Padding(
              padding: const EdgeInsets.all(8.0),
              child: Column(
                mainAxisAlignment: MainAxisAlignment.end,
                mainAxisSize: MainAxisSize.min,
                children: [
                  getIcon(context),
                  Text(elapsedTime),
                ],
              ),
            ),
          ),
          isLast
              ? const SizedBox.shrink()
              : Container(
                  width: 20,
                  height: 2,
                  color: jobStateStage == JobStateStage.success
                      ? Colors.green
                      : Colors.grey,
                ),
        ],
      ),
    );
  }

  Widget getIcon(BuildContext context) {
    final theme = Theme.of(context);
    final erroColor = theme.colorScheme.error;
    final primary = theme.primaryColor;
    switch (jobStateStage) {
      case JobStateStage.pending:
        return const Icon(
          Icons.circle,
          color: Colors.grey,
        );
      case JobStateStage.started:
        return Icon(
          Icons.run_circle_rounded,
          color: primary,
        );
      case JobStateStage.error:
        return Icon(
          Icons.error_rounded,
          color: erroColor,
        );
      case JobStateStage.success:
        return const Icon(
          Icons.check_circle,
          color: Colors.green,
        );
    }
  }
}
