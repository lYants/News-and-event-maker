# Dwengo News/Event Object Maker

This application allows users to create news and event objects for the Dwengo website. It provides a graphical interface that automatically generates the required Markdown header, so users do not need to edit the code manually.

The program also includes a built-in text editor for writing the content of the Markdown file.

After selecting the folder where the Dwengo website repository is cloned, the application automatically places newly created news and event objects in the correct directories within the local repository.


## Installation

Navigate to "Releases", select the latest release and download ObjectMaker.exe and updater.exe. These files are to be kept in the same folder in order for the auto-updater to work. Run ObjectMaker.exe to launch the application. If a newer version is available, it will update automatically. During this process, a terminal may pop up. Not to worry, it should close on its own and the updated app should launch.

Build app from code: pyinstaller --onefile --windowed --add-data "styles;styles" objectmaker.py