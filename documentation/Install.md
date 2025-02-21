Install bnkSEPA on your computer
================================

`This document is shows you how to install bnkSEPA for the Windows x64 platform. It should work with other OSs`  

To install bnkSEPA you need to:

- [Install bnkSEPA on your computer](#install-bnksepa-on-your-computer)
  - [Install Python v3.x](#install-python-v3x)
  - [Install Microsoft C++ Build Tools](#install-microsoft-c-build-tools)
  - [Install 7-Zip](#install-7-zip)
  - [Download bnkSEPA](#download-bnksepa)
    - [READ CAREFULLY](#read-carefully)
  - [Download support modules](#download-support-modules)
  - [Configure your settings](#configure-your-settings)
  - [Configure the *secrets.py* file](#configure-the-secretspy-file)


<a name="Python"></a>Install Python v3.x
-------------------

![Download Python from www.python.org](../images/InstallPython-001.PNG "Download Python").

Go to the [Python Website](https://www.python.org) and from the Download tab select the latest version of Python for your platform.

![Install the Python installer](../images/InstallPython-002.PNG "Install Python").

Click the installer. Before selecting the **Install Now** option, check the *Add Python 3.x to PATH*.  This will avoid you having to type the path to the Python interpreter in order to execute the program.



<a name="MSBldTools"></a>Install Microsoft C++ Build Tools
---------------------------------

![Download MS Build Tools](../images/InstallMSC++BuildTools--001.PNG "Download MS Build Tools").

Go to the [Microsoft C++ Build Tools page](https://visualstudio.microsoft.com/visual-cpp-build-tools/) and download the installer.

![Install the Desktop Development with C++ option](../images/InstallMSC++BuildTools--002.PNG "Install Desktop Development with C++").

Execute the installer and check the option **Desktop Development with C++**.   
Click install.  

![Exit the installer](../images/InstallMSC++BuildTools--003.PNG "Exit MS Built Tools installer").

Once the selected options have been installed exit.



<a name="7Zip"></a>Install 7-Zip
---------------------------------

Part of the process of generating the SCT file requires that you zip the XLSM file.

While there are a number of programs that generate zip file, 7Zip is the one that has been tested with this solution.

7-Zip is free software with open source and can be downloaded from [the 7-Zip web site](https://www.7-zip.org).



<a name="DLbnkSEPA"></a>Download bnkSEPA
----------------

![Login to bnkSEPA GitHub page](../images/InstallbnkSEPA--001.PNG "Login to bnkSEPA GitHub page").

Go to the [bnkSEPA GitHub repository](https://www.github.com/chribonn/bnkSEPA).

Click the Watch option on the page. You will be asked to register with GitHub.  *Registration is Free*.

>> The Watch option will inform you whenever bnkSEPA is enhanced and improved with new features.

>> Registering also allows you to contribute the the project and request new options.

Click the **Starred** option to help spread the word about this project.


![Download bnkSEPA code archive](../images/InstallbnkSEPA--002.PNG "Download bnkSEPA code archive").

Click the **Code** button and select the option **Download ZIP**. This will save the program to your computer.

*This tutorial will install the solution on the computer desktop.*


![Extract bnkSEPA code archive](../images/InstallbnkSEPA--003.PNG "Extract bnkSEPA code archive").

In *File Explorer* right click on the zip archive and select the option **Extract All...**.


![Extract bnkSEPA code archive](../images/InstallbnkSEPA--004.PNG "Extract bnkSEPA code archive").

### READ CAREFULLY

If you accept the directory suggested by the **Extract All...** option the code will be extracted into `Desktop ==> bnkSEPA ==> bnkSEPA`.

Backspace as shown above to get the installation to be at `Desktop ==> bnkSEPA`

Click the **Extact** button.



<a name="DLModules"></a>Download support modules
----------------


![Open Command Prompt](../images/InstallModules--001.PNG "Open command prompt").

1. Click on the bnkSEPA folder to open it.
2. Type **cmd** in the address bar.
3. This will cause a command window to open. The prompt should end with the bnkSEPA folder name.


![Install dependent libraries](../images/InstallModules--002.PNG "Install dependent libraries").

1. Type **Installer\install.cmd** and press enter. The environment and the modules this solution uses will be setup and installed.
2. Type Close the command prompt window.


![Perform recommended upgrades](../images/InstallModules--003.PNG "Install dependent libraries")

If prompted to update any components, copy the command into the command window.


<a name="UserConfig"></a>Configure your settings
----------------


## Configure the *secrets.py* file


![Edit secrets.py](../images/Configure-001.PNG "Open secrets.py for editing").

Right click on **secrets.py** file and choose *Edit with IDLE*. (You can also edit the file with a text editor such as *Notepad*).


![Edit secrets.py](../images/Configure-002.PNG "What to edit in secrets.py").

There are 4 settings you can edit in **secrets.py**:

  * **zip_file** - name of the zip file you will use to zip the XL workbook with your transactions.
  * **zip_pass** - password you will use when archving the file
  * **xl_file** - name of the XL file with your transactions

![Secrets entries.py](../images/Configure-003.PNG "Where the different secrets.py codes are used").


## YouTube Video

[![Watch the video](http://img.youtube.com/vi/x1cGcz2AZdQ/0.jpg)][https://www.youtube.com/watch?v=x1cGcz2AZdQ)
