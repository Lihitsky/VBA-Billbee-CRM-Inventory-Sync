# VBA-Billbee-CRM-Inventory-Sync

## Description
This VBA project is designed to fetch current inventory data from the Billbee CRM and insert it into an Excel table. The project includes three main components:
- `MainModule`: The main module that initiates the data fetching and insertion process.
- `ApiModule`: A module for making API requests to Billbee.
- `CRMClient` and `DataProcessor`: Classes for interacting with Billbee and processing the data.

## Dependencies
- [JsonConverter](https://github.com/VBA-tools/VBA-JSON): A library for handling JSON in VBA. Make sure it is added to the project.

## Installation
1. Download and add the JsonConverter library to your VBA project.
2. Insert the code from the `MainModule`, `ApiModule`, `CRMClient`, and `DataProcessor` files into the corresponding modules and classes in VBA.

## Usage
1. Open the Excel file and go to the VBA editor.
2. Ensure all modules and classes are added to the project.
3. Run the `Main` macro from the `MainModule`.

## Configuration
All configuration parameters (such as API_KEY, BILLBEE_USERNAME, and BILLBEE_API_PASSWORD) are set in the `CRMClient` class.

## Improvements
To improve performance, you can make asynchronous requests to the server.

## License
This project is licensed under the MIT License. See the LICENSE file for details.
