### Cleanup after style change in Zotero for Word

Many users encounter a problem after converting citations from an author-date (or any other inline) style to footnotes/endnotes and vice versa. In English these types of reference should be positioned differently respective the punctuation: footnote/endnote references are preceded by all punctuation marks except dashed and quotation marks and are not separated by a space from the preceding word. Author-date references should be preceded by a space and should precede any punctuation marks. 
Correct examples: 
> It was suggested in earlier literature,<sup>4</sup> but I disagree with that!!!<sup>5</sup>   
> It was suggested in earlier literature (Fu 2015, 5), but I disagree with that (Me 2016, 77)!!!

After changing the citation style with Zotero Word plugin the formatting becomes incorrect.
Examples after switching citation style: 
> It was suggested in earlier literature,(Fu 2015, 5) but I disagree with that!!!(Me 2016, 77)   
> It was suggested in earlier literature <sup>4</sup>, but I disagree with that <sup>5</sup>!!!

This issue was raised in at least two forum discussions, and no solution was proposed:  [Change citation style before-after periods](https://forums.zotero.org/discussion/38758/change-citation-style-before-after-periods) and [https://forums.zotero.org/discussion/56749/chicago-manual-of-style-citations-footnotes-appear-before-punctuation-marks](https://forums.zotero.org/discussion/56749/chicago-manual-of-style-citations-footnotes-appear-before-punctuation-marks).

To address this issue, two macros are proposed that search for all Zotero references in the document and correct them after switching from footnote/endnote references  to author-year references or from author-year references  to footnote/endnote references.

`CleanUpAfterChangingAuthorDateToNotes` should be run after author-date references were converted to footnotes/endnotes.
`CleanUpAfterChangingNotesToAuthorDate` should be run after footnotes/endnotes were converted to author-date references.

I only tested them in Word 2016 under Windows 7, but I guess they should word under any version of Word for Windows 2007-2016.
