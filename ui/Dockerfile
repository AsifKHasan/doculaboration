ARG flutter_version

FROM ghcr.io/cirruslabs/flutter:${flutter_version} AS build-env

# Start to run Flutter commands
# doctor to see if all was installes ok
RUN flutter doctor -v

# stup new folder as the working directory
ARG api_base_url
ENV WORKDIR=/doculaboration
WORKDIR $WORKDIR

# copy source code to folder
COPY . .

# Run build: 1 - clean, 2 - pub get, 3 - build web
RUN cd $WORKDIR && \
  flutter clean && \
  flutter pub get && \
  flutter pub run build_runner build -d && \
  flutter build web --dart-define=api_base_url=$api_base_url

# use nginx to deploy
FROM nginx:1.25.2-alpine

# copy the info of the builded web app to nginx
COPY --from=build-env /doculaboration/build/web /usr/share/nginx/html

CMD ["nginx", "-g", "daemon off;"]
