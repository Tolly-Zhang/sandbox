# Startup Apps Using Task Scheduler

[![Download ZIP](https://img.shields.io/badge/Download-ZIP-blue?style=for-the-badge&logo=github)](https://download-directory.github.io/?url=https://github.com/Tolly-Zhang/sandbox/tree/main/configs/task-scheduler-startup/)

This configuration is for using Task Scheduler to start applications on startup.

## Setup

1. Open Task Scheduler > Import Task > Select the `startup.xml` file.
2. In the **General** tab:
   1. Under **Name**, give the task a name.
   2. If you want to run the task at login instead of startup, click **Change User of Group** > under **Enter the object name to select (examples)**, enter in your username.
3. Under **Actions**, add a path to the application you want to start on startup. You can also add arguments if needed.
4. Make any edits under the **Conditions** and **Settings** tabs if needed.
5. Save and exit Task Scheduler.


## Usage
The application runs automatically on startup or login (depending on your settings). 
You can see the task's run history in Task Scheduler to confirm it is running as expected.