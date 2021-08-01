# google-classroom-sheets-bidi

A set of bi-directional translation/synchronization utilities between Google Classroom and Google Spreadsheet.

# Installation

```
git clone https://github.com/kubohiroya/google-classroom-sheets-bidi.git
cd google-classroom-sheets-bidi
yarn install
clasp login
clasp create --title "MyClassrooms" --type sheets --rootDir ./dist
yarn deploy
clasp open
````

# Usage

Open *"Util"* menu to invoke bi-directional translation from/to the Google Spreadsheet.

## Bi-directional translation of a set of Google Classroom Course Definitions to/from rows of a Google Spreadsheet document.

![image](https://user-images.githubusercontent.com/1578247/118382152-7a812680-b62d-11eb-8ee7-dafd914a5e59.png)

## Bi-directional translation of a Google Form document(a form definition) to/from a Google Spreadsheet document.

![image](https://user-images.githubusercontent.com/1578247/118382215-f67b6e80-b62d-11eb-9f5d-1f31f718c8d4.png)

Enjoy!
