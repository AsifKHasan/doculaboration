import 'dart:html';

///GET REQUEST

void sse(
  String url, {
  Function(String?)? onEvent,
  Function(dynamic)? onError,
  Function()? onDone,
}) {
  // This is for Flutter Web
  final httpReq = HttpRequest();

  httpReq
    ..open('GET', url)
    ..setRequestHeader('Content-Type', 'text/event-stream')
    ..onProgress.listen((event) {
      onEvent?.call(httpReq.responseText);
    })
    ..onLoadEnd.listen((event) {
      onDone?.call();
    })
    ..onError.listen((event) {
      final log = httpReq.status == 0
          ? "Could not connect to the server"
          : httpReq.statusText;
      onError?.call(log);
    })
    ..send();
}
