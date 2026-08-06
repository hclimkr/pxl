# PXL (Unreleased)

> Draft notes for the next release. The version is fixed when the release is cut — at that point this file is
> renamed to `vX.Y.Z.md`, its heading takes the version, and the matching `## [Unreleased]` section in
> `CHANGELOG.md` is retitled the same way.

Diagnostic message release for **PXL** — a misconfigured CSV import now names the attribute it choked on instead of failing as an unclassified system error, an export that fails while parsing a string value says so in its own words, and the library's internal message keys are grouped by binding direction so that a message can no longer sit under a name that hides which side it comes from.

No API change: no public signature moved and no new type is exposed. Two things change for a caller — the exception type raised by an invalid `importCsvCharset` / `importCsvDelimiter`, and the wording of the export-side codec parse messages.

## Highlights

  - **An unusable CSV import charset or delimiter is reported as an argument error**: `importCsvCharset("UTF8-typo")` used to fail with a `PxlSystemException` that named neither the attribute nor the value, because `Charset.forName(...)` rejects its input with an unchecked exception that the surrounding `IOException`-only `try` never caught — the same held for `importCsvDelimiter('\n')`, and for `'"'`, which collides with the quote character. Both calls are now made and normalized before that block, so they raise `PxlArgumentException` naming the attribute, with the original exception kept as the cause. Catching `PxlException` still catches them
  - **Export parse failures are worded as export failures**: writing parses strings too — a `String` field value, or an `exportSample` on a typed column, is parsed into the target type before it is written back out — so a bad value used to be reported with the wording of an import failure: `'abc' is not a valid Duration value.` The export path now has its own messages, opening with the value's origin: `the export value 'abc' is not a valid Duration value.` (Korean: `출력할 값 'abc'은(는) 올바른 형식의 Duration 값이 아닙니다.`) This covers five messages — the generic parse failure and the enum/custom-object parse pairs — and reading is unaffected: import wording is byte-identical to 0.9.3
  - **A failure while writing no longer reads as one while reading**: the practical effect of the above is that a message finally distinguishes "the spreadsheet held a value I could not parse" from "the object I was asked to write held one". Both were phrased identically before, which pointed at the wrong end of the binding whenever an `exportSample` or a `String` field carried a malformed value

## Internal

  - **Diagnostic keys grouped by binding direction**: a key thrown on only one side now starts with that side — `builder.export.*`, `core.import.*`, `meta.export.*`, `codec.import.*` — and the format, where there is one, follows it (`core.import.csv.fileNameCountMismatch`). Keys genuinely shared by both sides keep a neutral name: `core.sheet.countExceeded` is thrown by the exporter and by both importers, `codec.columnType.unsupported` by all three codec entry points. The keys live in `internal/i18n`, are not public API, and none of the shared messages changed wording, so this is invisible to calling code
  - **Two keys had been hiding their direction**: `builder.workbookSheetExclusive` and `meta.maskingInvalid` read as direction-neutral although neither is reachable while importing — the export builders are the only callers of the first, and masking is an export-only attribute. Grouping the namespace surfaced both
  - **Codec parse keys split rather than renamed**: because export parses strings as well, `codec.parse.invalid` and the enum/object parse keys were being thrown from both directions under one name, so a rename could not have been correct either way. Each became a pair. The shared private helpers `importStringToEnum` / `importStringToObject` take a `forExport` flag to select the key; their entry-point overloads were already split by direction, so no call site changed
