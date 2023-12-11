
# TabulateQR

TabulateQR is a simple tool to create customized tables for organizing sample data linked with QR codes. 

It facilitates the loading of Excel data into the first sheet, with the first column designated as `QR Code`.
| QR Code                  | Col1  | Col2  |
|--------------------------|-------|-------|
| Id1401B102P3S11D30Y2024  | data1 | data2 |

Users can dynamically add more columns in the tool to record data associated with the QR Code. The tool utilizes a webcam to scan QR Codes and adds a new row for each scanned code in the data table. If it detects that the scanned QR code already exists, it seamlessly scrolls to the corresponding entry in the data table.

Additionally, users can incorporate a second sheet in the Excel file to provide information about QR Code decoding. For example, if the QR Code is `Id1401B102P3S11D30Y2024`, where `Id` corresponds to LabID, `B` to Bed, `P` to Plot, `S` to Sampling time, `D` to Depth, and `Y` to Year, the QR Code is decoded upon exporting the table to Excel.

| QR Code                  | LabID | Bed | Plot | Sampling Time | Depth | Year | Col1  | Col2  |
|--------------------------|-------|-----|------|---------------|-------|------|-------|-------|
| Id1401B102P3S11D30Y2024  | 1401  | 102 | 3    | 11            | 30    | 2024 | data1 | data2 |


In the second sheet of the Excel file, users can include two columns with headers `QR Code` and `Decoding` as follows:

| QR Code | Decoding       |
|---------|-----------------|
| Id      | LabID           |
| B       | Bed             |
| P       | Plot            |
| S       | Sampling Time   |
| D       | Depth           |
| Y       | Year            |
