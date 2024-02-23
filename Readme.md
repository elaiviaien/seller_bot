# Telegram Channel Manager

The Telegram Channel Manager is a Python-based toolset designed to assist in the automation and management of Telegram channels and groups. By leveraging the Pyrogram library, this project offers functionalities such as channel joining, ID management, message handling, and administrative operations, all through a series of interconnected scripts.

## Features

- **Channel and Group ID Management**: Automate the extraction and handling of Telegram channel/group IDs from an Excel file.
- **Administrative Tools**: Perform administrative tasks such as joining channels, sending messages, and managing channel information.
- **Message Handling**: Automate the process of sending messages and managing interactions within channels or groups.
- **Excel Integration**: Seamlessly manage and track channel information using an Excel document (`list.xlsx`), enabling easy updates and retrieval of data.
- **Image Manipulation**: Additional support for image manipulation for channel/group content preparation.

## Components

- **pyro_main.py**: The core script containing the main functionalities for channel management, Excel file interaction, and image manipulation.
- **pyro_admin.py**: A dedicated script for administrative purposes, including channel joining, message management, and handling Excel data.
- **pyro_ceck.py**: A utility script aimed at checking, verifying, and joining channels, primarily focusing on maintenance and verification tasks.

## Getting Started

### Prerequisites

- Python 3.6 or above.
- Pyrogram library.
- PIL (Python Imaging Library) for image manipulation.
- Openpyxl for Excel file management.

### Installation

1. Clone this repository to your local machine.
2. Install the required Python packages by running `pip install -r requirements.txt`.
3. Make sure to have a valid `api_id` and `api_hash` from [Telegram's API Development Tools](https://my.telegram.org/auth).
4. Update the `list.xlsx` file with the Telegram channels/groups you wish to manage.

### Usage

Each script can be run independently based on the task at hand:

- For main channel management operations, run `python pyro_main.py`.
- For administrative tasks, execute `python pyro_admin.py`.
- For channel checking and verification, use `python pyro_ceck.py`.

## Disclaimer

This project is intended for educational and administrative assistance in managing Telegram channels and groups. The developers hold no liability for misuse or actions taken by users.
