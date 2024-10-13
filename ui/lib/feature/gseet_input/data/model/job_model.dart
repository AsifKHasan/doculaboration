import 'package:equatable/equatable.dart';
import 'package:ui/feature/gseet_input/data/model/job_state_model.dart';

class JobQueueItem extends Equatable {
  final String name;
  final List<JobModel> jobList;
  const JobQueueItem({
    required this.name,
    this.jobList = const [],
  });

  @override
  List<Object?> get props => [
        name,
        jobList,
      ];

  JobQueueItem copyWith({
    String? name,
    List<JobModel>? jobList,
  }) {
    return JobQueueItem(
      name: name ?? this.name,
      jobList: jobList ?? this.jobList,
    );
  }

  JobQueueItem copyWithUpdate(int index, JobModel jobModel) {
    final list = List<JobModel>.from(jobList);
    list.insert(index, jobModel);
    list.removeAt(index + 1);
    return copyWith(jobList: list);
  }
}
