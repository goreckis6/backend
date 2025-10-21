#!/bin/bash
# Manual Python environment setup script

echo "Setting up Python environment manually..."

# Install Python and dependencies
sudo apt update
sudo apt install -y python3 python3-pip python3-venv python3-dev
sudo apt install -y libjpeg-dev libpng-dev libtiff-dev libraw-dev
sudo apt install -y build-essential libffi-dev

# Create virtual environment
sudo mkdir -p /opt/venv
sudo chown -R debian:debian /opt/venv
python3 -m venv /opt/venv

# Install packages
/opt/venv/bin/pip install --upgrade pip
/opt/venv/bin/pip install Pillow rawpy

# Test the environment
echo "Testing Python environment..."
/opt/venv/bin/python --version
/opt/venv/bin/python -c "import rawpy; print('rawpy OK')" || echo "rawpy failed"
/opt/venv/bin/python -c "import PIL; print('Pillow OK')" || echo "Pillow failed"

echo "Python environment setup complete!"
