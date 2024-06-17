# VBox-PowerPortable

A Portable VirtualBox instllation created with Batch Scripts / Powershell.

## Installation

> Note: Most of the source code is opensource so you can run the scripts completely using iwr and iex.

1. Download the latest release from the releases page.
2. Extract the zip file to a folder.
3. Run the script in order to install the Portable VirtualBox.
4. Done!

## Usage

You can start the VirtualBox by running below listed command in order:

### Getting VirtualBox

To get a copy of latest version of VirtualBox, run the following command:

```powershell
.\get_vbox.ps1
```

### Installing VirtualBox

VirtualBox requires a few services to be running in order to work properly. This script will install the required services and dependencies.

> Note: ⚠️ This script requires administrative privileges.

```powershell
.\setup_vbox.ps1
```

### Starting VirtualBox

To start VirtualBox, run the following command:

```powershell
.\start_vbox.ps1
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
