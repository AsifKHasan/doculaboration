import 'package:flutter_bloc/flutter_bloc.dart';
import 'package:injectable/injectable.dart';

@injectable
class SelectedItemBloc extends Cubit<(int, int)?> {
  SelectedItemBloc() : super(null);
  setSelection((int, int) jobQuery) {
    emit(jobQuery);
  }
}
