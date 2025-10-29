# Use Node.js 18 with Debian base for better package support
FROM node:20-bookworm

# Install system dependencies for image processing and Python
RUN apt-get update && apt-get install -y --no-install-recommends \
    libvips-dev \
    libraw-bin \
    dcraw \
    ghostscript \
    imagemagick \
    libmagick++-dev \
    libreoffice \
    libreoffice-impress \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-java-common \
    default-jre-headless \
    unoconv \
    calibre \
    exiftool \
    pandoc \
    psmisc \
    fonts-dejavu fonts-liberation locales \
    python3 \
    python3-pip \
    python3-venv \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables for Python and ensure Ghostscript is in PATH
ENV PATH="/opt/venv/bin:/usr/bin:$PATH"
ENV GS_PROG=/usr/bin/gs

# Install Python dependencies in virtual environment
COPY requirements.txt ./
RUN python3 -m venv /opt/venv
RUN pip install --no-cache-dir -r requirements.txt

# Set working directory
WORKDIR /app

# Copy package files
COPY package.json ./
COPY tsconfig.json ./

# Install all dependencies (including dev dependencies for build)
RUN npm install

# Copy source code and Python scripts
COPY src/ ./src/
COPY scripts/ ./scripts/
COPY viewers/ ./viewers/

# Build the application
RUN npm run build

# Remove dev dependencies to reduce image size
RUN npm prune --production

# Create non-root user for security with proper home directory
RUN groupadd -r appuser && useradd -r -g appuser -d /home/appuser -m appuser
RUN chown -R appuser:appuser /app

# Create necessary directories for LibreOffice and set permissions
RUN mkdir -p /home/appuser/.cache/dconf \
    && mkdir -p /home/appuser/.config/libreoffice/4/user \
    && mkdir -p /tmp/libreoffice \
    && chown -R appuser:appuser /home/appuser \
    && chmod -R 755 /home/appuser \
    && chmod -R 777 /tmp/libreoffice

# Set environment variables for LibreOffice and Java
ENV HOME=/home/appuser
ENV TMPDIR=/tmp
ENV DCONF_PROFILE=/dev/null
ENV SAL_USE_VCLPLUGIN=svp
ENV JAVA_HOME=/usr/lib/jvm/default-java
ENV PATH="$JAVA_HOME/bin:$PATH"

# Initialize LibreOffice user profile as root before switching to appuser
RUN libreoffice --headless --invisible --nocrashreport --nodefault --nofirststartwizard --nologo --norestore --accept='socket,host=localhost,port=2002;urp;' & \
    sleep 5 && \
    pkill -9 soffice || true && \
    pkill -9 oosplash || true

USER appuser

# Expose port
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:3000/health || exit 1

# Start the application with garbage collection enabled
CMD ["node", "--expose-gc", "dist/server.js"]
