# ui
This is the frontend application for Doculaboration.

## Environment Setup
Prerequisites:
- See the [Setup project credential and config](../README.md)
- Install and setup flutter [https://docs.flutter.dev/get-started/install](https://docs.flutter.dev/get-started/install)
- Make sure [`api service is running on port 8200`](../api/README.md)
```bash
cd ui
flutter pub get
flutter pub run build_runner build -d
```

## Run the application
```bash
cd ui
flutter run -d chrome --web-port=8300 --dart-define="api_base_url=http://localhost:8200"
```
