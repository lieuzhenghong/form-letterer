# form-letterer

## How to use
You need a `data.csv` with header row.
Each row corresponds to one document to be created.

![Image of csv](./img/data_csv.png)

and a specially marked-up docx file.

The .docx file supports csv variables, delineated with ``backticks``, and arbitrary python code with `{{brackets}}`.

![Image of letter generation](./img/letters.png)

Anything wrapped in ``backticks`` will be grabbed from the CSV.

Anything in `{{brackets}}` will be `eval()'ed`. I often find myself using `round()` and ternary expressions (`x if y else z`).

## Important notes

Files will be generated in the 'build' folder. **Please create a empty build folder in the directory first!**

File names will be generated using the 'name' column of `data.csv`. The name of the files will be `letter_{name}.docx`. **Please ensure that your data.csv has a column called 'name'!**

## Example
```
Dear `title`,

Based on your ${{‘{:,}’.format(`amount`)}} invested capital, you will own {{round(`amount`/3000000 * 100, 2)}}% of iGInt’s secondary common shares ownership in Nerdwallet (equivalent to {{‘{:,}’.format(int(round(`amount`/3000000 * 500000)))}} common shares).
```
