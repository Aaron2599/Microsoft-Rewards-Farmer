An undetected Microsoft Rewards Farmer that collects about 60% of the points you can earn on PC when you go idle for more than 5 minutes.

Everything is built with the average Windows user in mind, but you can still download and run it manually with Python. The MSI file installs it as a Windows application, binds it to startup, and automatically handles the installation of [UV](https://github.com/astral-sh/uv). 

Microsoft-Point-Farmer.vbs is the entry point that runs installer.bat. The installer ensures [UV](https://github.com/astral-sh/uv) is installed, then calls start.bat to run main.py with [UV](https://github.com/astral-sh/uv) using updated environment variables.
