# Beer Festival Printing

Create a virtual environment and install reportlab.

## Cask Labels

See labelfields.csv for the required fields.

The following assumes that the input CSV file has one row per cask.

```
python3 casklabels.py \
    --labelfile labeltypes.csv \
    --labeltype <type from first column of labeltypes.csv> \
    --fieldfile labelfields.csv \
    --datafile <input csv file> \
    --outputfile <output pdf file>
```

The following assumes that the input CSV file has a column Quantity which
indicates the number of casks, this number of labels will be printed, with the
Cask column being generated.

```
python3 casklabels.py \
    --labelfile labeltypes.csv \
    --labeltype <type from first column of labeltypes.csv> \
    --fieldfile labelfields.csv \
    --quantity Quantity --enum Cask \
    --datafile <input csv file> \
    --outputfile <output pdf file>
```
