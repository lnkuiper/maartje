#!/bin/bash
python labels.py "$1"
pdflatex latex_labels.tex > /dev/null
rm -rf *.log *.aux *.tex
