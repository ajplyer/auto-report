# Auto Report

Auto Report was written and designed for use by the production team at Stone Creek Coffee. 

Multiple batch reports are created daily through ShipStation as a .csv file, these reports are used to quantify ecommerce coffee fulfillment needs.
As these reports are first generated, they are entirely unsorted and contain line items not pertinent to coffee fulfillment. An example is pictured below:

![original batch report](https://github.com/ajplyer/auto-report/assets/119643565/2b2f95d8-a981-4d02-8b00-dd1f2f94776b)

These reports were then manually sorted and formatted to appear as pictured below:

![completed batch report](https://github.com/ajplyer/auto-report/assets/119643565/7a6079db-5ffc-4430-a71e-7181055a0945)

While not complex, the process was tedious and required some knowledge of excel. Auto Report fixes this by completely sorting and formatting the batch reports, including the translation of "Coffee Of The Month" subscriptions to their respective coffees.

## Installation

To install, simply download the "Auto Report" folder. No additional downloads are required.

## Usage

As mentioned, this program is designed for a highly specific use case. Other users will likely not find it applicable to their needs. However, several batch reports have been included as originally downloaded for demonstration purposes.

The exported .csv file, of any name, must be placed into the Auto Report folder. Then simply running the "auto_report" executable will produce a finished .xlsx file named "batch".

The names in all text files must be typed exactly as they appear in the original .csv sheet, anything not in coffee_list.txt will be deleted. The items will be sorted in the order they appear in 'coffee_list.txt'.
