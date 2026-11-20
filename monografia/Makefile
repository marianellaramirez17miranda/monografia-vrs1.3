.PHONY: all clean

all: main.pdf

main.pdf: main.tex $(wildcard capitulos/*.tex) $(wildcard config/*.tex) bibliografia/referencias.bib
	latexmk -pdf -interaction=nonstopmode main.tex

clean:
	latexmk -C
	rm -f *.bbl *.blg *.aux *.log *.toc *.out
