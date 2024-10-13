import 'package:flutter/material.dart';
import 'package:gap/gap.dart';
import 'package:ui/feature/gseet_input/presentation/gsheet_input_widget.dart';
import 'package:flutter_bloc/flutter_bloc.dart';
import 'package:ui/feature/gseet_input/presentation/job_queue_widget.dart';
import 'package:ui/feature/shared/presentation/bloc/job_queue_bloc.dart';
import 'package:ui/feature/shared/presentation/bloc/selected_item_bloc.dart';

import '../util/file_extension.dart';

class HomeWidget extends StatelessWidget {
  final JobQueueBloc jobListBloc;
  final SelectedItemBloc selectedItembloc;

  const HomeWidget({
    super.key,
    required this.jobListBloc,
    required this.selectedItembloc,
  });

  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        GsheetNameInputPage(onSubmitGsheetName: _onSubmitGsheetName),
        const Gap(8.0),
        Expanded(
          child: BlocBuilder<JobQueueBloc, JobQueueState>(
            bloc: jobListBloc,
            builder: (context, state) {
              return state.when(
                initial: () => const SizedBox.shrink(),
                update: (jobQueueItemList) => JobQueueWidget(
                  onTap: (index, jobIndex) {
                    selectedItembloc.setSelection((index, jobIndex));
                  },
                  jobList: jobQueueItemList,
                ),
              );
            },
          ),
        )
      ],
    );
  }

  _onSubmitGsheetName(String gsheetName, List<FileExtension> list) {
    jobListBloc.addJob(
      gsheetName,
      downloadExtensions: list,
    );
  }
}
