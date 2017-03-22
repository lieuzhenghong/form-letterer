import csv
import re
from docx import Document


def replace_text(old, rep):
    '''
    Replaces a string with another.
    Supports cross-run replacement.
    '''
# Get the start and end of the string
    idx = para.text.find(old)
    end = idx + len(old)
    i = 0
    runs = []
    added = False
    for run in para.runs:
        prev = i
        i += len(run.text)
        if (i >= end):
            if added is False:
                added = True
                # this is the first run. it starts before idx and ends after
                # end.  Therefore, we must keep the before idx part and the
                # after end part
                run.text = run.text[0:idx-prev] + \
                    rep + \
                    run.text[idx-prev+len(old):]
                break
            elif added is True:
                # This run contains the tail of the search string. We keep
                # everything after end.
                run.text = run.text[end-prev:]
                break
        if (i > idx):
            if added is False:
                added = True
                # this is the first run. it starts before idx and ends after
                # idx. we should replace this first run with rep and remove
                # the rest
                run.text = run.text[0:idx-prev] + rep
            else:
                # these are the runs that are not needed; they lie completely
                # within idx and end
                run.text = ''

def generate_letters(csv_filename, letter_filename):
    # Read the file and put it into an array
    with open('{}.csv'.format(csv_filename), newline='') as csvfile:
        reader = csv.reader(csvfile)
        data = []
        for idx, row in enumerate(reader):
            if idx == 0:
                headers = row
            else:
                data.append(row)

    for idx, row in enumerate(data):
        document = Document('{}.docx'.format(letter_filename))
        for para in document.paragraphs:
            m = re.findall(r'`(.+?)`', para.text)
            for match in m:
                # Find the start and end of the substring
                rep = str(row[headers.index(match)])
                n = '`'+match+'`'
                replace_text(n, rep)

        for para in document.paragraphs:
            m = re.findall(r'{{(.+?)}}', para.text)
            for match in m:
                match2 = match.replace('‘', "'")
                match2 = match2.replace('’', "'")
                x = str(eval(match2))
                replace_text('{{' + match + '}}', x)
        document.save('./build/{0}_{1}.docx'.format(letter_filename, row[headers.index('name')]))

if __name__ == '__main__':
    import sys
    generate_letters(sys.argv[1], sys.argv[2])
