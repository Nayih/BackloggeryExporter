# Backloggery Exporter

This script fetches your entire game library from [Backloggery](https://www.backloggery.com), organizes it by platform, and exports the data to an **Excel spreadsheet**. It provides a structured, offline backup of your collection for easy access and archival purposes.
<small>**Note**: It will not save the game history introduced in the most recent versions of Backloggery, nor the lists to which each game belongs.</small>

![](https://i.imgur.com/MSK6SXb.jpg)

## Features

- Retrieves your complete game list from Backloggery.
- Organize games by platform, collection (used in older versions) and order of appearance similar to backloggery.
- Exports the data to an `.xlsx` file.
- Easy way to create a personal backup of your collection.

## Requirements

- Python 3.14+
- `Requests 2.32.4+`
- `XlsxWriter 3.2.5+`

You can install the dependencies with:

```bash
pip install -r requirements.txt
```