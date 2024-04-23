# Stage 1: Build Python 3.8 environment
FROM python:3.8 AS builder

# Stage 1: Install necessary system dependencies
RUN apt-get update && \
    apt-get install -y \
    build-essential \
    libpq-dev \
    vim

# Stage 1: Install Python dependencies
WORKDIR /app
COPY requirements.txt .
RUN pip install --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# Stage 2: Create final image
FROM ubuntu:latest

# Set environment variable to prevent interactive prompts during package installation
ENV DEBIAN_FRONTEND noninteractive


# Stage 2: Copy binaries
COPY --from=builder /usr/local/ /usr/local/
COPY --from=builder /usr/bin/python3 /usr/bin/python3
COPY --from=builder /usr/bin/vim /usr/bin/vim

# Stage 1: Install necessary system dependencies
RUN apt-get update && \
    apt-get install -y \
    libreoffice \
    libreoffice-script-provider-python \
    python3-uno

ARG api_port
ENV api_port=${api_port}

# Stage 2: Set working directory and copy application code
ENV WORKDIR=/doculaboration
WORKDIR ${WORKDIR}
COPY . .
# setup config.yml file
RUN mv $WORKDIR/gsheet-to-json/conf/config.yml.dist $WORKDIR/gsheet-to-json/conf/config.yml
RUN mv $WORKDIR/json-to-context/conf/config.yml.dist $WORKDIR/json-to-context/conf/config.yml
RUN mv $WORKDIR/json-to-docx/conf/config.yml.dist $WORKDIR/json-to-docx/conf/config.yml
RUN mv $WORKDIR/json-to-latex/conf/config.yml.dist $WORKDIR/json-to-latex/conf/config.yml
RUN mv $WORKDIR/json-to-odt/conf/config.yml.dist $WORKDIR/json-to-odt/conf/config.yml

# Give executable permission for .sh files
RUN chmod +x **/*.sh

# FIXME: put the credentail and config folder to the correct directory
