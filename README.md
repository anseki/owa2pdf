# owa2pdf

Convert MS-Office (Word/Excel/PowerPoint) documents to PDF files via [Office Online](https://office.com/) (and [OneDrive](https://onedrive.live.com/)).

In some cases, the application on non-Windows machine (e.g. Web Application) uses something (e.g. [Apache OpenOffice](http://www.openoffice.org/)) to convert files. But, it is difficult to parse the layout of the MS-Office document.  
owa2pdf uses the Office Online that is provided by Microsoft, that working is high quality.

## Usage

Install [PhantomJS](http://phantomjs.org/) 1.9+ first.  
`owa2pdf.js` and `jquery-2.0.3.min.js` must be put on same directory. And, Microsoft account is needed. (see https://signup.live.com/)

```shell
phantomjs owa2pdf.js -u "user@example.com" -p "password" -i /path/source.docx -o /path/dest.pdf
```

The `--ignore-ssl-errors=true` option may be needed.

```shell
phantomjs --ignore-ssl-errors=true owa2pdf.js -u "user@example.com" -p "password" -i /path/source.docx -o /path/dest.pdf
```

## cleanpdf.pl

When the PDF file which is made by Office Online is opened by Adobe Reader, "Print" dialog-box is displayed. It's done by script which was embedded by Office Online. And some other things are embedded by Office Online.  
`cleanpdf.pl` removes some things which embedded by Office Online.  This is Perl script which needs [`PDF::API2`](http://search.cpan.org/perldoc?PDF%3A%3AAPI2) and [`CAM::PDF`](http://search.cpan.org/perldoc?CAM%3A%3APDF) modules. (e.g. `cpanm PDF::API2`, `cpanm CAM::PDF`)

At first, install 2 modules by your favorite installer.

```shell
cpanm PDF::API2 CAM::PDF
```

Then, clean files by `cleanpdf.pl`.

```shell
./cleanpdf.pl /path/dest.pdf
```

## Notes

+ owa2pdf is slow.
+ If you have Windows machine and MS-Office, using them is better.
+ If Microsoft releases Office Online API someday, using that is better. (owa2pdf is Web scraping.)
+ Your application might have to retry calling the script. The converting sometimes fails by various causes (e.g. network, MS server, etc.).
+ Your application might have to use the plural accounts, if it converts many files successively.

