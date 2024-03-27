import 'package:equatable/equatable.dart';
import 'package:ui/feature/gseet_input/data/model/job_model.dart';

class SelectedJobModel extends Equatable {
  final Map<String, JobQueueItem> modelMap;
  final (int, int) selectedForLogging;

  const SelectedJobModel({
    required this.modelMap,
    required this.selectedForLogging,
  });

  @override
  List<Object?> get props => [
        modelMap,
        selectedForLogging,
      ];

  SelectedJobModel copyWith({
    Map<String, JobQueueItem>? modelMap,
    (int, int)? selectedForLogging,
  }) {
    return SelectedJobModel(
      modelMap: modelMap ?? this.modelMap,
      selectedForLogging: selectedForLogging ?? this.selectedForLogging,
    );
  }
}
