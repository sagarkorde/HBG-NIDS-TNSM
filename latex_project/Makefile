# ============================================================
#  Makefile for HBG-NIDS IEEE TNSM Paper
#  Usage:
#    make          → full build (pdf + bib + pdf×2)
#    make quick    → single pdflatex pass (no bib refresh)
#    make clean    → remove aux files
#    make cleanall → remove aux + pdf
# ============================================================

MAIN     = main
LATEX    = pdflatex
BIBTEX   = bibtex
LATEXOPT = -interaction=nonstopmode -halt-on-error

.PHONY: all quick clean cleanall

all: $(MAIN).pdf

$(MAIN).pdf: $(MAIN).tex references.bib
	$(LATEX) $(LATEXOPT) $(MAIN)
	$(BIBTEX) $(MAIN)
	$(LATEX) $(LATEXOPT) $(MAIN)
	$(LATEX) $(LATEXOPT) $(MAIN)

quick:
	$(LATEX) $(LATEXOPT) $(MAIN)

clean:
	rm -f $(MAIN).aux $(MAIN).bbl $(MAIN).blg $(MAIN).log \
	       $(MAIN).out $(MAIN).toc $(MAIN).lof $(MAIN).lot \
	       $(MAIN).fls $(MAIN).fdb_latexmk

cleanall: clean
	rm -f $(MAIN).pdf
