import sys
import pandas as pd


def label_from_row(row):
    tel_nr = row['Mobiel / GSM']
    if type(tel_nr) == float:
        tel_nr = row['Telefoonnummer']
    bericht = row['Bericht voor uitgifte']
    label_str = row['Naam'] + '\n\n'
    if type(tel_nr) != float:
        label_str += tel_nr + '\n\n'
    label_str += row['Adres'] + '\n\n'
    label_str += row['Postcode'] + '\t' + row['Plaats'] + '\n\n'
    if type(bericht) != float:
        label_str += bericht + '\n\n'
    label_str += row['Pakket']
    return label_str


def generate_script(df):
    script = r"""
    \documentclass[12pt]{article}
    \usepackage[utf8]{inputenc}
    \pagenumbering{gobble}
    \begin{document}

    """
    for i, row in df.iterrows():
        if i % 2 == 0:
            script += '\n' + r'\vbox{' + '\n'
        if i % 2 == 1:
            script += '\n' + r'\nointerlineskip' + '\n'
        script += r'\begin{minipage}[t][0.5\textheight][t]{\textwidth}' + '\n'
        script += r'\Huge' + '\n'
        script += r'\begin{flushright}' + '\n'
        script += str(i + 1) + '\n'
        script += r'\end{flushright}' + '\n'
        script += r'\centering' + '\n'
        script += label_from_row(row) + '\n'
        script += r'\end{minipage}' + '\n'
        if i % 2 == 1:
            script += '}\n' + r'\newpage' + '\n'
    if i % 2 == 0:
        script += '}\n'
    script += '\n' + r'\end{document}' + '\n'
    return script


def load_xls(filename):
    df = pd.read_excel(filename)
    columns = df.iloc[3].values
    df = df[5:]
    df.columns = columns
    df = df[df['Kaartnummer'].notnull()]
    df = df.reset_index(drop=True)
    return df


def main():
    filename = sys.argv[1]
    df = load_xls(filename)
    script = generate_script(df)
    tex_filename = 'latex_labels.tex'
    with open(tex_filename, 'w+') as f:
        print(script, file=f)


if __name__ == '__main__':
    main()
