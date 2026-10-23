# Optional Go editor

`core-go/` is reserved for a possible `excelize` editor. It remains inactive unless
the `umya-spreadsheet` comparison on sanitized real templates shows broad
preservation failures and a narrow OOXML patch is too costly. If enabled, Go is
limited to editing existing workbooks, uses `CGO_ENABLED=0`, and is not a reader
or new-report writer.
