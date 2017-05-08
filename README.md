# form-letterer

## Introduction and motivation
I was given a task to generate twenty-five very similar Word documents. They
were letters which meant that each address, addressee, salutation etc. were
different. I didn't want to go through and manually edit each one so I built a
programmatic solution.

Form Letterer is a programmatic form letter generator. It takes a *template*
word document and a *data file* and automatically creates form letters.

![Explanatory image of how Form Letterer works](./img/overview.png)

Essentially, it's a version of Word's Mail Merge.

## Why not Mail Merge?

I didn't have Microsoft Word on my laptop at the time (lazy to dual boot). This
solution doesn't need Microsoft Word and uses only free software.

As the solution is deployed on a local server, multiple users can use it
without any prior configuration.

Additionally, this solution allows for things that can't be done in Mail Merge
like conditionals and arithmetic. Mail Merge is a simple search-and-replace but
you can do things here like 
```
{{`amount1` * `amount2`}} # multiplication of data values
# 'urgency' is a header row in your data file
{{'URGENT REMINDER' if `urgency` == 'urgent' else 'Reminder:'}} 
```

## How to use

1. `build/` folder
2. `data.csv` data file
3. `letter.docx` template file

![Explanatory image of how Form Letterer works with folder layout](./img/overview2.png)

First, **create an empty `build` folder in the same directory**, as the
letters will be generated there.  

Then, create a `data.csv` with header row.
Each row corresponds to one document to be created.

Here's an example of a `data.csv` data file.

![Image of csv](./img/data_csv.png)

File names will be generated using the 'name' column of `data.csv`. The name of
the files will be `letter_{name}.docx`. **Please ensure that your data.csv has
a column called 'name'!**

Lastly, create a `letter.docx` file in the same directory.

The .docx file supports csv variables, delineated with ``backticks``, and
arbitrary python code with `{{brackets}}`.

![Image of letter generation](./img/letters.png)

Anything wrapped in ``backticks`` will be grabbed from the CSV.

Anything in `{{brackets}}` will be `eval()'ed`. I often find myself using
`round()` and ternary expressions (`x if y else z`).


## Troubleshooting
TODO
