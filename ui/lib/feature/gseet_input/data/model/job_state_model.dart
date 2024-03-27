// ignore_for_file: must_be_immutable

import 'package:equatable/equatable.dart';

import '../../util/file_extension.dart';
import '../../util/job_state_stage.dart';

sealed class JobModel extends Equatable {
  String log;
  JobStateStage jobStateStage;
  bool skippable;
  String name;

  JobModel({
    this.log = "No logs",
    this.skippable = true,
    this.jobStateStage = JobStateStage.pending,
    this.name = "",
  });

  @override
  List<Object?> get props;
}

class DownloadJobModel extends JobModel {
  double speed;
  String speedUnit;
  double progressPercentage;
  String pureFileName;
  FileExtension downloadFileExtension;

  DownloadJobModel({
    super.log,
    super.jobStateStage,
    super.skippable,
    super.name = "Download",
    this.speed = 0,
    this.speedUnit = "MB/s",
    this.progressPercentage = 0,
    required this.pureFileName,
    required this.downloadFileExtension,
  });

  @override
  List<Object?> get props => [
        log,
        jobStateStage,
        skippable,
        name,
        speed,
        speedUnit,
        progressPercentage,
        pureFileName,
        downloadFileExtension,
      ];
  String get fileNameForSave => "$pureFileName.${downloadFileExtension.name}";
}

class ProcessJobModel extends JobModel {
  ProcessJobModel({
    super.log,
    super.jobStateStage,
    super.skippable = true,
    super.name = "Process",
  });

  @override
  List<Object?> get props => [
        log,
        jobStateStage,
        skippable,
      ];
}
