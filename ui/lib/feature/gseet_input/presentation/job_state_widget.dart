import 'package:flutter/material.dart';
import 'package:ui/feature/gseet_input/util/job_state_stage.dart';

class JobStateWidget extends StatelessWidget {
  final JobStateStage jobStateStage;
  final String stateName;

  const JobStateWidget({
    super.key,
    required this.stateName,
    required this.jobStateStage,
  });

  @override
  Widget build(BuildContext context) {
    return Tooltip(
      message: "Job($stateName): ${jobStateStage.name}",
      child: getIcon(context),
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
