# owa2pdf

Convert MS-Office (Word/Excel/PowerPoint) documents to PDF files via Office Online ([OneDrive](https://onedrive.live.com/)).

The application on non-Windows machine (e.g. Web Application) use something (e.g. [Apache OpenOffice](http://www.openoffice.org/)) to convert files. But, it is difficult to parse the layout of the MS-Office document.  
This script use the Office Online that is provided by Microsoft.

## Usage

Install [PhantomJS](http://phantomjs.org/) 1.9+ first.  
`owa2pdf.js` and `jquery-2.0.3.min.js` must be put on same directory. And, Microsoft account is needed. (see https://signup.live.com/)

```
phantomjs owa2pdf.js -u "user@example.com" -p "password" -i "/path/source.docx" -o "/path/dest.pdf"
```

The `--ignore-ssl-errors=true` option may be needed.

```
phantomjs --ignore-ssl-errors=true owa2pdf.js -u "user@example.com" -p "password" -i "/path/source.docx" -o "/path/dest.pdf"
```

## Notes

+ This script is slow.
+ If you have Windows machine and MS-Office, using them is better.
+ If Microsoft release Office Online API someday, using that is better. (This script is Web scraping.)
+ Your application might have to retry calling the script. The converting sometimes fail by various causes (e.g. network, MS server, etc.).
+ Your application might have to use the plural accounts, if it convert many files successively.

