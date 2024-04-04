// ignore_for_file: must_be_immutable

import 'package:equatable/equatable.dart';

import '../../util/file_extension.dart';
import '../../util/job_state_stage.dart';

sealed class JobModel extends Equatable {
  String log;
  JobStateStage jobStateStage;
  bool skippable;
  String name;
  late DateTime _startedAt;
  Duration? _elapsedTime;

  JobModel({
    this.log = "No logs",
    this.skippable = true,
    this.jobStateStage = JobStateStage.pending,
    this.name = "",
  });

  void start() {
    _startedAt = DateTime.now();
    _elapsedTime = Duration.zero;
  }

  String getElapsedTime() {
    if (_elapsedTime != null && jobStateStage == JobStateStage.started) {
      _elapsedTime = DateTime.now().difference(_startedAt);
    }

    var timeSpent = "";
    int d = _elapsedTime?.inDays ?? 0;
    int h = _elapsedTime?.inHours ?? 0;
    int m = _elapsedTime?.inMinutes ?? 0;
    int s = _elapsedTime?.inSeconds ?? 0;

    if (d > 0) {
      h = h - d * 24;
      timeSpent += "${d}d";
    }
    if (h > 0) {
      m = m - h * 60;
      if (timeSpent.isNotEmpty) {
        timeSpent += " ";
      }
      timeSpent += "${h}h";
    }
    if (m > 0) {
      if (timeSpent.isNotEmpty) {
        timeSpent += " ";
      }
      s = s - m * 60;
      timeSpent += "${m}m";
    }
    if (s > 0) {
      if (timeSpent.isNotEmpty) {
        timeSpent += " ";
      }
      timeSpent += "${s}s";
    }
    if (d == 0 && h == 0 && m == 0 && s == 0) {
      timeSpent = "0s";
    }
    return timeSpent;
  }

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
        _startedAt,
        _elapsedTime,
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
        _startedAt,
        _elapsedTime,
      ];
}
