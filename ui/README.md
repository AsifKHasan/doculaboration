# ui

This is the ui application for Doculaboration.

## Environment Setup

1. Install fvm
2. Generate files
```bash
fvm flutter pub get
fvm flutter pub run build_runner build -d
```
## Run the application

```bash
fvm flutter run -d chrome --web-port=8080 # backend application is configured to allow requests from only 8080 port
```
