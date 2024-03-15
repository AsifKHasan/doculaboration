import 'package:dio/dio.dart';
import 'package:injectable/injectable.dart';
import 'package:pretty_dio_logger/pretty_dio_logger.dart';

const kBaseUrl = 'BASEURL';

@module
abstract class ApiModule {
  Dio get dio => Dio()
    ..interceptors.addAll(
      [
        PrettyDioLogger(),
      ],
    );

  @Named(kBaseUrl)
  String get baseUrl => 'http://localhost:8000';
}
