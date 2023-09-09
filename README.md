# Auto Report

Auto Report was written and designed for use by the production team at Stone Creek Coffee. 

Multiple batch reports are created daily through ShipStation as a .csv file, these reports are used to quantify ecommerce coffee fulfillment needs.
As these reports are first generated, they are entirely unsorted and contain line items not pertinent to coffee fulfillment. An example is pictured below:

![original batch report](https://github.com/ajplyer/auto-report/assets/119643565/2b2f95d8-a981-4d02-8b00-dd1f2f94776b)

These reports were then manually sorted and formatted to appear as pictured below:

![completed batch report](https://github.com/ajplyer/auto-report/assets/119643565/962ac808-194f-470f-9247-ce8cfc92dd7f)

While not complex, the process was tedious and required some knowledge of excel. Auto Report fixes this by completely sorting and formatting the batch reports, including the translation of "Coffee Of The Month" subscriptions to their respective coffees.

## Installation

To install, simply download the "Auto Report" executable, and the three text files "coffee_list", "com_med_dark", and "com_light". No additional downloads are required.

Auto Report.exe was created using Auto PY to EXE

## Usage

As mentioned, this program is designed for a highly specific use case. Other users will likely not find it applicable to their needs. However, several batch reports have been included as originally downloaded for demonstration purposes.

Simply drag and drop the unsorted .csv batch report into the white entry box, you will now be able to click "Sort Batch Report". Once sorting is complete you will be able to click "Open Sorted Batch" which, as the text implies, will open the fully sorted and formatted batch report.
A "Clear Entry" button has been included for convenience, pressing it will clear the entry box and disable the other two previously mentioned buttons.

The names in all text files must be typed exactly as they appear in the original .csv sheet, anything not in coffee_list.txt will be deleted. The items will be sorted in the order they appear in 'coffee_list.txt'.

## TODO

- [ ] GUI for improved user experience
  - [x] Drag and Drop for batch to sort
  - [ ] Coffee List and Coffee Of The Month as lists in program
    - [ ] Edit lists in program
    - [ ] Rearrange Coffee List in program
  - [x] "Present" sorted batch report in program to avoid requiring user to find the file
  - [ ] Design UI
    - [ ] Icon
    - [ ] Stylize drag and drop entry box
    - [ ] Format colors to match Stone Creek Coffee color schemes

