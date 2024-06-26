import 'package:dio/dio.dart';
import 'package:injectable/injectable.dart';
import 'package:retrofit/retrofit.dart';

import '../../../../core/api_module.dart';

part 'home_api_client.g.dart';

@RestApi()
@Injectable()
abstract class HomeApiClient {
  @factoryMethod
  factory HomeApiClient(
    Dio dio, {
    @Named(kBaseUrl) String baseUrl,
  }) = _HomeApiClient;

  @DioResponseType(ResponseType.bytes)
  @GET("/{extension}/{gsheetName}/download")
  Future<List<int>> downloadFile(
    @Path() String gsheetName,
    @Path() String extension, {
    @ReceiveProgress() ProgressCallback? onProgress,
  });
}
