import 'package:get_it/get_it.dart';
import 'package:injectable/injectable.dart';
import 'package:ui/core/dependency.config.dart';

final getIt = GetIt.instance;

@InjectableInit()
configureDependencies() => getIt.init();
