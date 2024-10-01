# Use the official windows (servercore) image
FROM mcr.microsoft.com/windows/servercore:ltsc2019

# Set the working directory
WORKDIR /app

# Copy the local files to the container
COPY . .

# Install Python
RUN powershell -Command \
    Invoke-WebRequest -Uri  https://www.python.org/ftp/python/3.9.6/python-3.9.6.exe -OutFile python-installer.exe; \
    Start-Process python-installer.exe -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1' -NoNewWindow -Wait; \
    Remove-Item python-installer.exe;

# Install pip
RUN powershell -Command \
    python --version; \
    python -m ensurepip --default-pip; \
    pip --version;

# Install dependencies (if any)
RUN powershell -Command \
    pip install -r requirements.txt;

# Set the command to run your script
CMD ["python", "refresh_excel.py"]
