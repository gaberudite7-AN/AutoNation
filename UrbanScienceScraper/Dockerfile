FROM mcr.microsoft.com/azure-functions/python:4-python3.13

# Install system dependencies for Playwright
RUN apt-get update && apt-get install -y \
    wget gnupg libnss3 libatk1.0-0 libatk-bridge2.0-0 libcups2 libxss1 \
    libasound2 libxshmfence-dev libgbm-dev libgtk-3-0 libdrm2 libxcomposite1 \
    libxrandr2 libxdamage1 libx11-xcb1 libxinerama1 libgl1-mesa-glx libgl1-mesa-dri

# Install Python dependencies
COPY requirements.txt /
RUN pip install -r /requirements.txt

# Install Playwright browsers
RUN python -m playwright install

# Copy your function app code
COPY . /home/site/wwwroot
WORKDIR /home/site/wwwroot