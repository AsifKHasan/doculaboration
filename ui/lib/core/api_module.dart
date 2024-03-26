import 'package:dio/dio.dart';
import 'package:flutter/material.dart';
import 'package:injectable/injectable.dart';
import 'package:pretty_dio_logger/pretty_dio_logger.dart';

const kBaseUrl = 'BASEURL';

@module
abstract class ApiModule {
  Dio get dio {
    return Dio()
      ..interceptors.addAll(
        [
          PrettyDioLogger(),
        ],
      );
  }

  @Named(kBaseUrl)
  String get baseUrl {
    const baseUrl = String.fromEnvironment('api_base_url');
    if (baseUrl.isEmpty) {
      throw "Environment not setup correctly. Read the ui/README.md carefully";
    }
    debugPrint("Expecting the api service to run on this url: $baseUrl");
    return baseUrl;
  }
}
