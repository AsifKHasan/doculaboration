import 'package:bloc/bloc.dart';
import 'package:freezed_annotation/freezed_annotation.dart';
import 'package:injectable/injectable.dart';
import 'package:ui/feature/gseet_input/data/model/job_model.dart';
import 'package:ui/feature/gseet_input/util/sse_event_mobile.dart'
    if (dart.library.html) 'package:ui/feature/gseet_input/util/sse_event_web.dart';
import 'package:ui/feature/gseet_input/util/job_state_stage.dart';

//Local imports
import 'package:ui/utility/save_file_mobile.dart'
    if (dart.library.html) 'package:ui/utility/save_file_web.dart';

import '../../../../core/api_module.dart';
import '../../../gseet_input/data/model/job_state_model.dart';
import '../../../gseet_input/data/source/home_api_client.dart';
import '../../../gseet_input/util/file_extension.dart';

part 'job_queue_bloc.freezed.dart';

@freezed
class JobQueueState with _$JobQueueState {
  const JobQueueState._();
  const factory JobQueueState.initial() = _initial;

  const factory JobQueueState.update({
    required List<JobQueueItem> jobQueueItemList,
  }) = _update;
}

@Injectable()
class JobQueueBloc extends Cubit<JobQueueState> {
  final queueItemIndexMap = <String, int>{};
  final jobNameCount = <String, int>{};
  final jobQueue = <JobQueueItem>[];
  final HomeApiClient homeApiClient;
  final String baseUrl;

  @factoryMethod
  JobQueueBloc(
    this.homeApiClient, {
    @Named(kBaseUrl) required this.baseUrl,
  }) : super(const JobQueueState.initial());

  void addJob(
    String jobName, {
    List<FileExtension> downloadExtensions = const [
      FileExtension.pdf,
      FileExtension.odt,
      FileExtension.json,
    ],
  }) {
    if (queueItemIndexMap[jobName] != null) {
      jobNameCount.update(
        jobName,
        (count) => count + 1,
        ifAbsent: () => 1,
      );
    }
    final name = jobName +
        (jobNameCount[jobName] == null ? "" : jobNameCount[jobName].toString());

    final jobQueueItem = JobQueueItem(name: name, jobList: [
      ProcessJobModel(),
      ...downloadExtensions.map(
        (extension) => DownloadJobModel(
          pureFileName: jobName,
          downloadFileExtension: extension,
        ),
      ),
    ]);
    jobQueue.add(jobQueueItem);
    queueItemIndexMap.putIfAbsent(name, () => jobQueue.length - 1);

    /// Send the original user input to the server
    jobQueueItem.jobList.first.start();
    emit(JobQueueState.update(jobQueueItemList: List.from(jobQueue)));
    sse(
      "$baseUrl/process/$jobName",
      onEvent: (event) => _onSseEvent(
        jobQueueItem.jobList.first as ProcessJobModel,
        event,
      ),
      onDone: () {
        final jobList = jobQueueItem.jobList;
        final processJob = jobList[0];
        if (processJob.jobStateStage == JobStateStage.error) return;
        processJob.jobStateStage = JobStateStage.success;
        emit(
          JobQueueState.update(
            jobQueueItemList: List.from(jobQueue),
          ),
        );
        for (int index = 0; index < jobList.length; index++) {
          final jobModel = jobList[index];
          if (jobModel is DownloadJobModel) {
            jobModel.start();
            homeApiClient
                .downloadFile(
                  jobName,
                  jobModel.downloadFileExtension.name,
                  onProgress: (count, total) => _onDownloadProgress(
                    jobModel,
                    (count, total),
                  ),
                )
                .then(
                  (fileData) => _onFileDataReceived(jobModel, fileData),
                )
                .onError(
                  (e, s) => _onDownloadError(jobModel, (e, s)),
                );
          }
        }
      },
      onError: (e) => _onSseError(
        jobQueueItem.jobList.first as ProcessJobModel,
        e,
      ),
    );
  }

  void _onSseError(ProcessJobModel jobModel, e) {
    final processJob = jobModel;
    processJob.jobStateStage = JobStateStage.error;
    processJob.log = "${processJob.log}\n${e.toString()}";
    emit(
      JobQueueState.update(
        jobQueueItemList: List.from(jobQueue),
      ),
    );
  }

  void _onSseEvent(ProcessJobModel jobModel, String? event) {
    final processJob = jobModel;
    processJob.jobStateStage = JobStateStage.started;
    processJob.log = event.toString();
    emit(
      JobQueueState.update(
        jobQueueItemList: List.from(jobQueue),
      ),
    );
  }

  void _onDownloadProgress(JobModel jobModel, (int, int) progressPair) {
    final prevJobModel = jobModel;
    prevJobModel.log =
        "Download percentage: ${progressPair.$1 * 100 ~/ progressPair.$2}";
    prevJobModel.jobStateStage = JobStateStage.started;
    emit(
      JobQueueState.update(
        jobQueueItemList: List.from(jobQueue),
      ),
    );
  }

  void _onFileDataReceived(
    DownloadJobModel jobModel,
    List<int> fileData,
  ) {
    saveAndLaunchFile(fileData, jobModel.fileNameForSave).then(
      (value) {
        final prevJobModel = jobModel;
        prevJobModel.jobStateStage = JobStateStage.success;
        emit(
          JobQueueState.update(
            jobQueueItemList: List.from(jobQueue),
          ),
        );
      },
    ).onError((e, s) {
      jobModel.jobStateStage = JobStateStage.error;
      jobModel.log = "${jobModel.log}\n$e\n$s";
      emit(
        JobQueueState.update(
          jobQueueItemList: List.from(jobQueue),
        ),
      );
    });
  }

  void _onDownloadError(
    DownloadJobModel jobModel,
    (Object?, StackTrace) errorPair,
  ) {
    final e = errorPair.$1;
    final s = errorPair.$2;
    jobModel.jobStateStage = JobStateStage.error;
    jobModel.log = "${jobModel.log}\n$e\n$s";
    emit(
      JobQueueState.update(
        jobQueueItemList: List.from(jobQueue),
      ),
    );
  }
}
