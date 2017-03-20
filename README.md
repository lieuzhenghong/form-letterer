# form-letterer

## How to use
You need a `data.csv` with header row.
Each row corresponds to one document to be created.

![Image of csv](./img/data_csv.png)

and a specially marked-up docx file.

The .docx file supports csv variables, delineated with ``backticks``, and arbitrary python code with `{{brackets}}`.

![Image of letter generation](./img/letters.png)

Files will be generated in the 'build' folder. **Please create a empty build folder in the directory first!**

Anything wrapped in ``backticks`` will be grabbed from the CSV.

Anything in `{{brackets}}` will be `eval()'ed`. I often find myself using `round()` and ternary expressions (`x if y else z`).

## Example
```
Dear `title`,

Based on your ${{‘{:,}’.format(`amount`)}} invested capital, you will own {{round(`amount`/3000000 * 100, 2)}}% of iGInt’s secondary common shares ownership in Nerdwallet (equivalent to {{‘{:,}’.format(int(round(`amount`/3000000 * 500000)))}} common shares).
```
