A simple Stata project to test if it's possible to pre-format an Excel template
and then fill it with data via Stata.

Spoiler alert: It is possible — with limitations.


## The Good News

Formatting that is kept when Stata writes data to an existing Excel file:

- conditional formatting
- formatting of cells that are not empty (cells in the template file with text)
- column widths
- frozen rows & columns


## Limitations

Formatting that is lost when Stata writes to an existing Excel file (even when using `sheetmodify` and `keepcellfmt` options):

- formatting of cells that are empty (cells in the template file without text)

## Workaround to the Limitations

There are two ways to work around Stata's loss of "normal" Excel cell formatting.

#### Workaround 1

Style all cells using conditional formatting, using a custom condition such as
`=TRUE`. This always applies to the set styles to the specified range of the
conditional format.

However, conditional formatting options are limited.
For example, they do not include slanted/angled text. This is where
workaround 2 comes into play.

#### Workaround 2

Stata has a command for fine-grained manipulation of Excel files:
The `excelput` command.

Among many things, it can be used to apply (non-conditional) styling to Excel sheets. Here is how to bold and angle/slant the first row of a file:

```stata
putexcel set sdgindex.xlsx, modify sheet(sheetname)
putexcel (A1:ZZ1),                    /// target the entire first row
         overwritefmt                 /// remove any existing formatting
         bold                         /// bold the header row
         txtrotate(45)                //  angle/slant the text
putexcel clear
```


## Conclusion

Check out the code `master.do` for the full implementation details. Compare the
`template.xlsx` and `sdgindex.xlsx` files to see how the formatting changes
as Stata writes data to the file.
