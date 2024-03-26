# Environemnt to install flutter and build web
FROM debian:latest AS build-env

# install all needed stuff
RUN apt-get update
RUN apt-get install -y curl git unzip

# define variables
ARG FLUTTER_SDK=/usr/local/flutter
ARG FLUTTER_VERSION=3.19.3
ARG api_base_url

#clone flutter
ENV GIT_HTTP_LOW_SPEED_LIMIT=1000
ENV GIT_HTTP_LOW_SPEED_TIME=60
RUN git clone https://github.com/flutter/flutter.git $FLUTTER_SDK
# change dir to current flutter folder and make a checkout to the specific version
RUN cd $FLUTTER_SDK && git fetch && git checkout $FLUTTER_VERSION

# setup the flutter path as an enviromental variable
ENV PATH="$FLUTTER_SDK/bin:$FLUTTER_SDK/bin/cache/dart-sdk/bin:${PATH}"

# Start to run Flutter commands
# doctor to see if all was installes ok
RUN flutter doctor -v

# stup new folder as the working directory
ENV WORKDIR=/doculaboration
WORKDIR $WORKDIR
# copy source code to folder
COPY . .

# Run build: 1 - clean, 2 - pub get, 3 - build web
RUN flutter clean
RUN flutter pub get
RUN flutter pub run build_runner build -d
RUN flutter build web --dart-define="api_base_url=${api_base_url}"

# use nginx to deploy
FROM nginx:1.25.2-alpine

# copy the info of the builded web app to nginx
COPY --from=build-env /doculaboration/build/web /usr/share/nginx/html

CMD ["nginx", "-g", "daemon off;"]
